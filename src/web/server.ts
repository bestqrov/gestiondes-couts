import { mkdirSync, readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import 'dotenv/config';
import express, { type Response } from 'express';
import multer from 'multer';
import { extractDocumentText } from '../ocr/documentTextExtractor.js';
import { detectAndParsePair } from '../parser/detectAndParsePair.js';
import { mergeDeclaration } from '../merge/declarationMerger.js';
import { validateArticle } from '../domain/validators.js';
import { generateCombinedExcel } from '../excel/combinedExcelGenerator.js';
import { renderResultsPage, renderResultsFragment } from './renderResultsPage.js';
import {
  renderSuperAdminOverview,
  renderSuperAdminUsers,
  renderSuperAdminPlaceholder,
  renderSuperAdminCosts,
  renderSuperAdminSettings,
  renderSuperAdminGenerate,
} from './renderSuperAdminDashboard.js';
import { calculateLandedCost } from '../domain/costCalculator.js';
import {
  createSession,
  requireAuth,
  requireSuperAdmin,
  setSessionCookie,
  destroySession,
} from './auth.js';
import { getDatabase } from '../db/database.js';
import {
  findUserByUsername,
  verifyPassword,
  seedSuperAdminIfEmpty,
  listUsers,
  createUser,
  setUserDisabled,
  type UserRole,
} from '../db/usersRepository.js';
import {
  saveDeclaration,
  listAllDeclarations,
  getArticlesForDeclaration,
  searchDeclarationsByRedevable,
} from '../db/declarationsRepository.js';
import { getAppSettings, updateAppSettings } from '../db/appSettingsRepository.js';
import { isValidHexColor } from '../domain/colorUtils.js';
import { FONT_OPTIONS, renderBrandOverrideStyle, renderLogoImg } from './brandingStyles.js';
import type { Declaration } from '../domain/types.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const UPLOAD_DIR = path.join(PROJECT_ROOT, '.tmp-uploads');
const OUTPUT_DIR = path.join(PROJECT_ROOT, '.tmp-output');
// These don't exist yet on a fresh checkout/container — created locally by
// hand during earlier development, but a fresh deployment has neither.
mkdirSync(UPLOAD_DIR, { recursive: true });
mkdirSync(OUTPUT_DIR, { recursive: true });

const db = getDatabase();
const superAdminUsername = process.env.SUPERADMIN_USERNAME ?? 'redwan';
const superAdminPassword = process.env.SUPERADMIN_PASSWORD ?? 'redwan2026';
if (!process.env.SUPERADMIN_USERNAME || !process.env.SUPERADMIN_PASSWORD) {
  console.warn(
    'SUPERADMIN_USERNAME/SUPERADMIN_PASSWORD not set — falling back to default credentials for initial setup. Set these in production.'
  );
}
seedSuperAdminIfEmpty(db, superAdminUsername, superAdminPassword);

// Fixed bcrypt hash (cost 10, matching usersRepository's hashing cost) with
// no corresponding real password — used so a login attempt against a
// nonexistent username still pays bcrypt's ~50-100ms cost, instead of
// returning near-instantly and leaking which usernames exist via timing.
const DUMMY_PASSWORD_HASH = '$2a$10$r2UyLAu1mdPlxnjaJGSsP.XJFU3ietRR/a.INq8TYvFMb4rSxugbC';

const uploadHtml = readFileSync(path.join(__dirname, 'views/upload.html'), 'utf-8');
const loginHtml = readFileSync(path.join(__dirname, 'views/login.html'), 'utf-8');

// Single-user local tool (no auth, no concurrency handling) — holds the most
// recently generated declaration in memory so /download and the results
// preview can reference it without a database.
let lastDeclaration: Declaration | undefined;
let lastGeneratedFilePath: string | undefined;

// multer's default disk storage strips the original file extension, but
// extractDocumentText() dispatches on extension (.pdf vs image formats) —
// preserve it explicitly so uploaded files are routed correctly.
const storage = multer.diskStorage({
  destination: UPLOAD_DIR,
  filename: (_req, file, callback) => {
    const uniqueSuffix = `${Date.now()}-${Math.round(Math.random() * 1e9)}`;
    callback(null, `${uniqueSuffix}${path.extname(file.originalname)}`);
  },
});
const upload = multer({ storage });

const LOGO_MAX_BYTES = 2 * 1024 * 1024;
const LOGO_ALLOWED_MIME_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/webp',
  'image/svg+xml',
]);
// Kept entirely in memory (never written to disk) — converted straight to
// a data: URI and stored in the app_settings row, so it persists in the
// same place as the database with no separate file/volume to configure.
const logoUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: LOGO_MAX_BYTES },
});

// Express's res.download()/res.sendFile() run the path through encodeURI()
// before using it as a filesystem path (see node_modules/express/lib/
// response.js) — harmless for a path made only of URL-safe characters, but
// this project's own directory name ("gestion de couts") contains spaces,
// which encodeURI turns into "%20", so Express ends up looking for a file
// at a path that doesn't exist on disk (confirmed: reproduced locally and
// matches the exact "NotFoundError: Not Found ... at
// ServerResponse.download" stack trace seen in production logs). Reading
// the file ourselves and sending the buffer directly sidesteps this
// entirely — no URI encoding is ever applied to the filesystem path.
async function sendXlsxFile(res: Response, filePath: string, downloadName: string): Promise<void> {
  const buffer = await readFile(filePath);
  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.setHeader('Content-Disposition', `attachment; filename="${downloadName}"`);
  res.send(buffer);
}

const app = express();
app.use(express.urlencoded({ extended: false }));

function renderLoginHtml(errorBlock: string): string {
  return loginHtml
    .replace('{{ERROR_BLOCK}}', errorBlock)
    .replace('{{BRAND_OVERRIDE}}', renderBrandOverrideStyle(getAppSettings(db)));
}

app.get('/login', (_req, res) => {
  res.send(renderLoginHtml(''));
});

app.post('/login', (req, res) => {
  const { username, password } = req.body as { username?: string; password?: string };
  const errorBlock =
    '<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>Identifiant ou mot de passe incorrect.</span></div>';

  if (!username || !password) {
    res.status(401).send(renderLoginHtml(errorBlock));
    return;
  }

  const user = findUserByUsername(db, username);
  // Always run a bcrypt compare, even for a nonexistent username, against a
  // fixed dummy hash — otherwise a missing user short-circuits before the
  // ~50-100ms bcrypt call a wrong-password attempt incurs, letting response
  // timing reveal which usernames exist.
  const passwordMatches = verifyPassword(user?.passwordHash ?? DUMMY_PASSWORD_HASH, password);
  if (!user || user.disabledAt || !passwordMatches) {
    res.status(401).send(renderLoginHtml(errorBlock));
    return;
  }

  const sessionId = createSession({ userId: user.id, username: user.username, role: user.role });
  setSessionCookie(res, sessionId);
  res.redirect(user.role === 'superadmin' ? '/superadmin/dashboard' : '/');
});

// Tolerant of a missing/already-expired session cookie — placed before the
// requireAuth gate below so a logout attempt never itself 401s/redirects
// into a loop; there's simply nothing to clear in that case.
app.post('/logout', (req, res) => {
  destroySession(req, res);
  res.redirect('/login');
});

app.use(requireAuth);

app.get('/', (req, res) => {
  const isSuperAdmin = req.session?.role === 'superadmin';
  const navLink = isSuperAdmin
    ? '<a href="/superadmin/dashboard" style="margin-left:auto;font-size:13px;color:var(--brand-600);text-decoration:none;font-weight:600;">Gestion des comptes &rarr;</a>'
    : '';
  const roleBadge = isSuperAdmin
    ? '<span class="role-pill role-pill-superadmin">Superadmin</span>'
    : '<span class="role-pill role-pill-admin">Admin</span>';
  const settings = getAppSettings(db);
  res.send(
    uploadHtml
      .replace('{{ERROR_BLOCK}}', '')
      .replace('{{NAV_LINK}}', navLink)
      .replace('{{ROLE_BADGE}}', roleBadge)
      .replace('{{LOGO_IMG}}', renderLogoImg(settings))
      .replace('{{BRAND_OVERRIDE}}', renderBrandOverrideStyle(settings))
  );
});

app.post(
  '/generate',
  upload.fields([
    { name: 'liquidation', maxCount: 1 },
    { name: 'dum', maxCount: 1 },
  ]),
  async (req, res) => {
    const files = req.files as Record<string, Express.Multer.File[]>;
    const liquidationFile = files.liquidation?.[0];
    const dumFile = files.dum?.[0];

    try {
      if (!liquidationFile || !dumFile) {
        throw new Error('Les deux fichiers (Liquidation et DUM) sont requis.');
      }

      const liquidationOcr = await extractDocumentText(liquidationFile.path);
      const dumOcr = await extractDocumentText(dumFile.path);

      const { liquidation, dum } = detectAndParsePair(liquidationOcr.text, dumOcr.text);
      const declaration = mergeDeclaration(liquidation, dum);
      for (const article of declaration.articles) {
        validateArticle(article);
      }

      // Each request writes to its own uniquely-named file rather than a
      // fixed shared path — two overlapping /generate requests (e.g. a
      // double-submit, or two users at once) were racing on the same fixed
      // "Declaration.xlsx" path, so a request could try to res.download() a
      // file that a concurrent request had just truncated/replaced,
      // producing a spurious "Not Found" (confirmed in production logs).
      const generatedFilePath = path.join(OUTPUT_DIR, `declaration-${randomUUID()}.xlsx`);
      await generateCombinedExcel(declaration, generatedFilePath);

      lastDeclaration = declaration;
      lastGeneratedFilePath = generatedFilePath;

      // Persisted (unlike the in-memory `last*` state above, which only
      // serves this admin's own immediate results/download/cost-preview and
      // is wiped on every restart) so the superadmin's "Coût de produit"
      // page survives redeploys instead of going blank until someone
      // generates again.
      const cost = calculateLandedCost(declaration, dum.shipmentCost ?? {});
      saveDeclaration(db, {
        ownerUserId: req.session!.userId,
        declaration,
        shipmentCostFields: dum.shipmentCost ?? {},
        articleCosts: cost.articleCosts,
        totalLandedCost: cost.totalLandedCost,
        costEstimatePartial: cost.partial,
        excelFilePath: generatedFilePath,
      });

      await sendXlsxFile(res, generatedFilePath, 'Declaration.xlsx');
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      res.status(400).json({ success: false, error: message });
    }
  }
);

app.get('/results', (_req, res) => {
  if (!lastDeclaration) {
    res.redirect('/');
    return;
  }
  res.send(renderResultsPage(lastDeclaration));
});

// Same tables as /results, but as an HTML fragment (no page wrapper) meant
// to be fetched and injected inline into the upload page's success panel —
// "Afficher résultats" shows this in place instead of navigating away.
app.get('/last-declaration-results', (_req, res) => {
  if (!lastDeclaration) {
    res.status(404).send('<p>Aucune déclaration générée pour le moment.</p>');
    return;
  }
  res.send(renderResultsFragment(lastDeclaration));
});

// Duty-only cost per unit (sum of this article's tax montants / quantite) —
// not full landed cost (purchase price + freight + insurance), which needs
// the shipment-cost fields and persistence work tracked separately. Labeled
// clearly on the client as "coût douanier" so it isn't mistaken for the
// complete landed cost.
app.get('/last-declaration-cost-summary', (_req, res) => {
  if (!lastDeclaration) {
    res.status(404).json({ success: false, error: 'Aucune déclaration générée pour le moment.' });
    return;
  }

  const articles = lastDeclaration.articles.map((article) => {
    const totalTaxes = article.taxes.reduce((sum, tax) => sum + tax.montant, 0);
    return {
      numero: article.numero,
      nomArticle: article.nomArticle,
      hsCode: article.hsCode,
      pays: article.pays,
      quantite: article.quantite,
      totalTaxes,
      dutyCostPerUnit: article.quantite > 0 ? totalTaxes / article.quantite : 0,
    };
  });

  res.json({
    success: true,
    code: lastDeclaration.code,
    redevable: lastDeclaration.redevable,
    articles,
  });
});

app.get('/download', async (_req, res) => {
  if (!lastGeneratedFilePath) {
    res.redirect('/');
    return;
  }
  await sendXlsxFile(res, lastGeneratedFilePath, 'Declaration.xlsx');
});

app.get('/superadmin/dashboard', requireSuperAdmin, (_req, res) => {
  res.send(
    renderSuperAdminOverview(listUsers(db), listAllDeclarations(db).length, getAppSettings(db))
  );
});

// Lets a superadmin generate a declaration without leaving the sidebar
// dashboard — reuses the exact same /generate, /download,
// /last-declaration-cost-summary, and /last-declaration-results endpoints
// as the standalone admin tool at "/", just wrapped in the sidebar shell.
app.get('/superadmin/generate', requireSuperAdmin, (_req, res) => {
  res.send(renderSuperAdminGenerate(getAppSettings(db)));
});

app.get('/superadmin/users', requireSuperAdmin, (req, res) => {
  res.send(renderSuperAdminUsers(listUsers(db), req.session!.userId, getAppSettings(db)));
});

app.get('/superadmin/costs', requireSuperAdmin, (req, res) => {
  // Reads from the database (most recent declaration across all admins),
  // not the in-memory `lastDeclaration` — that state is per-process and
  // wiped on every restart/redeploy, which made this page go blank even
  // though declarations had already been generated before the restart.
  const [mostRecent] = listAllDeclarations(db);
  if (!mostRecent) {
    res.send(
      renderSuperAdminPlaceholder(
        'Coût de produit',
        "Aucune déclaration n'a encore été générée sur l'application. Le coût par produit s'affichera ici après une génération.",
        getAppSettings(db)
      )
    );
    return;
  }
  const searchQuery = typeof req.query.q === 'string' ? req.query.q.trim() : '';
  const searchResults = searchQuery ? searchDeclarationsByRedevable(db, searchQuery) : undefined;
  const articles = getArticlesForDeclaration(db, mostRecent.id);
  res.send(
    renderSuperAdminCosts(mostRecent, articles, getAppSettings(db), searchQuery, searchResults)
  );
});

app.get('/superadmin/settings', requireSuperAdmin, (_req, res) => {
  res.send(renderSuperAdminSettings(getAppSettings(db)));
});

app.post(
  '/superadmin/settings',
  requireSuperAdmin,
  (req, res, next) => {
    logoUpload.single('logo')(req, res, (err) => {
      if (err) {
        // multer throws for oversized files (LIMIT_FILE_SIZE) — surface it
        // the same way other form errors on this page are shown, instead
        // of an unhandled 500.
        res.status(400).send(
          renderSuperAdminSettings(
            getAppSettings(db),
            err.code === 'LIMIT_FILE_SIZE'
              ? 'Le logo dépasse la taille maximale autorisée (2 Mo).'
              : "Échec de l'envoi du fichier."
          )
        );
        return;
      }
      next();
    });
  },
  (req, res) => {
    const { companyName, brandColor, fontFamily } = req.body as {
      companyName?: string;
      brandColor?: string;
      fontFamily?: string;
    };

    if (brandColor && !isValidHexColor(brandColor)) {
      res
        .status(400)
        .send(renderSuperAdminSettings(getAppSettings(db), 'Couleur invalide.'));
      return;
    }
    const allowedFonts = new Set(FONT_OPTIONS.map((opt) => opt.value));
    if (fontFamily && !allowedFonts.has(fontFamily)) {
      res.status(400).send(renderSuperAdminSettings(getAppSettings(db), 'Police invalide.'));
      return;
    }

    const logoFile = req.file;
    if (logoFile && !LOGO_ALLOWED_MIME_TYPES.has(logoFile.mimetype)) {
      res
        .status(400)
        .send(
          renderSuperAdminSettings(
            getAppSettings(db),
            'Format de logo non supporté (PNG, JPEG, WEBP ou SVG uniquement).'
          )
        );
      return;
    }

    const updated = updateAppSettings(db, {
      companyName: companyName?.trim() ? companyName.trim() : null,
      brandColor: brandColor || null,
      fontFamily: fontFamily || null,
      ...(logoFile
        ? { logoDataUri: `data:${logoFile.mimetype};base64,${logoFile.buffer.toString('base64')}` }
        : {}),
    });

    res.send(renderSuperAdminSettings(updated, undefined, 'Réglages enregistrés.'));
  }
);

app.post('/superadmin/users', requireSuperAdmin, (req, res) => {
  const { username, password, role } = req.body as {
    username?: string;
    password?: string;
    role?: string;
  };

  const renderWithError = (message: string) => {
    res
      .status(400)
      .send(renderSuperAdminUsers(listUsers(db), req.session!.userId, getAppSettings(db), message));
  };

  if (!username || !password) {
    renderWithError("Identifiant et mot de passe sont requis.");
    return;
  }
  if (role !== 'admin' && role !== 'superadmin') {
    renderWithError('Rôle invalide.');
    return;
  }

  try {
    createUser(db, username, password, role as UserRole);
    res.redirect('/superadmin/users');
  } catch (error) {
    // better-sqlite3 throws a raw SqliteError (code SQLITE_CONSTRAINT_UNIQUE)
    // on a duplicate username — the only realistic failure mode here, since
    // username/password/role are already validated above.
    if (error instanceof Error && 'code' in error && error.code === 'SQLITE_CONSTRAINT_UNIQUE') {
      renderWithError(`L'identifiant « ${username} » est déjà utilisé.`);
      return;
    }
    throw error;
  }
});

function setDisabledAndRedirect(disabled: boolean) {
  return (req: express.Request, res: express.Response) => {
    const targetId = Number(req.params.id);
    // Self-lockout only applies to disabling — a superadmin re-enabling
    // their own account can't happen anyway (a disabled account can't log
    // in to reach this route), but guarding disable is essential: the UI
    // already hides the button for one's own row, but that's not a
    // substitute for a server-side check against a direct POST.
    if (disabled && targetId === req.session!.userId) {
      res
        .status(400)
        .send(
          renderSuperAdminUsers(
            listUsers(db),
            req.session!.userId,
            getAppSettings(db),
            'Vous ne pouvez pas désactiver votre propre compte.'
          )
        );
      return;
    }
    setUserDisabled(db, targetId, disabled);
    res.redirect('/superadmin/users');
  };
}

app.post('/superadmin/users/:id/disable', requireSuperAdmin, setDisabledAndRedirect(true));
app.post('/superadmin/users/:id/enable', requireSuperAdmin, setDisabledAndRedirect(false));

// Default matches Coolify's default "Ports Exposes" (3000) so the app works
// out of the box even if the PORT environment variable doesn't reach the
// running container (e.g. a Coolify env var scoped to build-time only).
const port = Number(process.env.PORT ?? 3000);
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});

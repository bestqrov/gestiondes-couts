import { mkdirSync, readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import 'dotenv/config';
import express, { type Response } from 'express';
import multer from 'multer';
import type { Collection } from 'mongodb';
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
import {
  findUserByUsername,
  verifyPassword,
  seedSuperAdminIfEmpty,
  listUsers,
  createUser,
  setUserDisabled,
  updateUsername,
  updatePassword,
  deleteUser,
  ensureUsersIndexes,
  USERS_COLLECTION,
  type UserDocument,
  type UserRole,
} from '../db/usersRepository.js';
import {
  getAppSettings,
  updateAppSettings,
  APP_SETTINGS_COLLECTION,
  DEFAULT_APP_SETTINGS,
  type AppSettingsDocument,
} from '../db/appSettingsRepository.js';
import { getMongoDb } from '../db/mongoClient.js';
import {
  saveTransaction,
  getMostRecentTransaction,
  countTransactions,
  getCountryProductCounts,
  searchTransactionsByRedevable,
  TRANSACTIONS_COLLECTION,
  type TransactionDocument,
  type CountryProductCount,
} from '../db/transactionsRepository.js';
import { isValidHexColor } from '../domain/colorUtils.js';
import {
  FONT_OPTIONS,
  renderBrandOverrideStyle,
  renderLogoImg,
  renderFaviconLink,
  renderLoginBadge,
  renderLoginTitle,
  renderContactRows,
} from './brandingStyles.js';
import { generateDeclarationPdf } from '../pdf/declarationPdfGenerator.js';
import type { Declaration } from '../domain/types.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const UPLOAD_DIR = path.join(PROJECT_ROOT, '.tmp-uploads');
const OUTPUT_DIR = path.join(PROJECT_ROOT, '.tmp-output');
// These don't exist yet on a fresh checkout/container — created locally by
// hand during earlier development, but a fresh deployment has neither.
mkdirSync(UPLOAD_DIR, { recursive: true });
mkdirSync(OUTPUT_DIR, { recursive: true });

// Users and app settings live in MongoDB now (same database as saved
// declarations) instead of a local SQLite file — a local file requires a
// Coolify persistent volume to survive redeploys, and that was repeatedly
// left unconfigured/misconfigured, silently wiping accounts and branding
// on every deploy. MongoDB Atlas is external to the container, so it
// doesn't have that failure mode.
async function getUsersCollection() {
  const mongoDb = await getMongoDb();
  return mongoDb.collection<UserDocument>(USERS_COLLECTION);
}
async function getSettingsCollection() {
  const mongoDb = await getMongoDb();
  return mongoDb.collection<AppSettingsDocument>(APP_SETTINGS_COLLECTION);
}

async function bootstrap(): Promise<void> {
  const usersCollection = await getUsersCollection();
  await ensureUsersIndexes(usersCollection);
  const superAdminUsername = process.env.SUPERADMIN_USERNAME ?? 'redwan';
  const superAdminPassword = process.env.SUPERADMIN_PASSWORD ?? 'redwan2026';
  if (!process.env.SUPERADMIN_USERNAME || !process.env.SUPERADMIN_PASSWORD) {
    console.warn(
      'SUPERADMIN_USERNAME/SUPERADMIN_PASSWORD not set — falling back to default credentials for initial setup. Set these in production.'
    );
  }
  await seedSuperAdminIfEmpty(usersCollection, superAdminUsername, superAdminPassword);
}

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

async function renderLoginHtml(errorBlock: string): Promise<string> {
  const settings = await getAppSettings(await getSettingsCollection());
  return loginHtml
    .replace('{{ERROR_BLOCK}}', errorBlock)
    .replace('{{FAVICON_LINK}}', renderFaviconLink(settings))
    .replace('{{BRAND_OVERRIDE}}', renderBrandOverrideStyle(settings))
    .replace('{{BRAND_LOGO_LEFT}}', renderLoginBadge(settings))
    .replace('{{BRAND_LOGO_RIGHT}}', renderLoginBadge(settings))
    .replace('{{BRAND_TITLE}}', renderLoginTitle(settings))
    .replace('{{CONTACT_ROWS}}', renderContactRows(settings));
}

// The login page is the very first thing every visitor hits — if MongoDB
// happens to be unreachable, showing a raw stack trace there would look
// like the whole app is broken. A bare-bones but on-brand-enough fallback
// page is friendlier and makes the actual problem (Mongo down) obvious.
const SERVICE_UNAVAILABLE_HTML = `<!doctype html>
<html lang="fr"><head><meta charset="utf-8" /><title>Service indisponible</title>
<style>body{font-family:-apple-system,sans-serif;max-width:480px;margin:15vh auto;padding:0 24px;color:#334155;text-align:center;}
h1{font-size:20px;color:#0f172a;}p{font-size:14px;line-height:1.5;color:#64748b;}</style></head>
<body><h1>Service temporairement indisponible</h1><p>Impossible de se connecter à la base de données pour le moment. Réessayez dans quelques instants.</p></body></html>`;

app.get('/login', async (_req, res) => {
  try {
    res.send(await renderLoginHtml(''));
  } catch (error) {
    console.error('Failed to render /login (MongoDB unreachable?):', error);
    res.status(503).send(SERVICE_UNAVAILABLE_HTML);
  }
});

app.post('/login', async (req, res) => {
  const { username, password } = req.body as { username?: string; password?: string };
  const errorBlock =
    '<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>Identifiant ou mot de passe incorrect.</span></div>';

  try {
    if (!username || !password) {
      res.status(401).send(await renderLoginHtml(errorBlock));
      return;
    }

    const user = await findUserByUsername(await getUsersCollection(), username);
    // Always run a bcrypt compare, even for a nonexistent username, against a
    // fixed dummy hash — otherwise a missing user short-circuits before the
    // ~50-100ms bcrypt call a wrong-password attempt incurs, letting response
    // timing reveal which usernames exist.
    const passwordMatches = verifyPassword(user?.passwordHash ?? DUMMY_PASSWORD_HASH, password);
    if (!user || user.disabledAt || !passwordMatches) {
      res.status(401).send(await renderLoginHtml(errorBlock));
      return;
    }

    const sessionId = createSession({ userId: user.id, username: user.username, role: user.role });
    setSessionCookie(res, sessionId);
    res.redirect(user.role === 'superadmin' ? '/superadmin/dashboard' : '/');
  } catch (error) {
    console.error('Failed to process /login (MongoDB unreachable?):', error);
    res.status(503).send(SERVICE_UNAVAILABLE_HTML);
  }
});

// Tolerant of a missing/already-expired session cookie — placed before the
// requireAuth gate below so a logout attempt never itself 401s/redirects
// into a loop; there's simply nothing to clear in that case.
app.post('/logout', (req, res) => {
  destroySession(req, res);
  res.redirect('/login');
});

app.use(requireAuth);

app.get('/', async (req, res) => {
  const isSuperAdmin = req.session?.role === 'superadmin';
  const navLink = isSuperAdmin
    ? '<a href="/superadmin/dashboard" style="margin-left:auto;font-size:13px;color:var(--brand-600);text-decoration:none;font-weight:600;">Gestion des comptes &rarr;</a>'
    : '';
  const roleBadge = isSuperAdmin
    ? '<span class="role-pill role-pill-superadmin">Superadmin</span>'
    : '<span class="role-pill role-pill-admin">Admin</span>';
  const settings = await getAppSettings(await getSettingsCollection());
  res.send(
    uploadHtml
      .replace('{{ERROR_BLOCK}}', '')
      .replace('{{NAV_LINK}}', navLink)
      .replace('{{ROLE_BADGE}}', roleBadge)
      .replace('{{LOGO_IMG}}', renderLogoImg(settings))
      .replace('{{FAVICON_LINK}}', renderFaviconLink(settings))
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

      // Persisted to MongoDB (unlike the in-memory `last*` state above,
      // which only serves this admin's own immediate results/download/
      // cost-preview and is wiped on every restart) so the superadmin's
      // "Coût de produit" page survives redeploys instead of going blank
      // until someone generates again. Only totals/metadata are saved —
      // the generated .xlsx file itself is never persisted to the database
      // (it stays a short-lived file on disk, referenced only by the
      // in-memory `lastGeneratedFilePath` above, for this same request's
      // immediate re-download).
      const cost = calculateLandedCost(declaration, dum.shipmentCost ?? {});
      const totalTaxes = declaration.articles.reduce(
        (sum, article) => sum + article.taxes.reduce((s, tax) => s + tax.montant, 0),
        0
      );
      try {
        const mongoDb = await getMongoDb();
        const collection = mongoDb.collection<TransactionDocument>(TRANSACTIONS_COLLECTION);
        await saveTransaction(collection, {
          ownerUserId: req.session!.userId,
          code: declaration.code,
          redevable: declaration.redevable,
          valeurTotaleDeclaree: dum.shipmentCost?.valeurTotaleDeclaree ?? null,
          totalTaxes,
          totalLandedCost: cost.totalLandedCost,
          costEstimatePartial: cost.partial,
          articles: declaration.articles.map((article) => ({
            numero: article.numero,
            hsCode: article.hsCode,
            nomArticle: article.nomArticle,
            pays: article.pays,
            quantite: article.quantite,
            costPerUnit: cost.articleCosts.find((c) => c.numero === article.numero)!.costPerUnit,
          })),
        });
      } catch (mongoError) {
        // The Excel file was already generated successfully — a MongoDB
        // hiccup shouldn't block the admin from getting their file, only
        // the superadmin "Coût de produit" history entry for this run.
        console.error('Failed to save transaction to MongoDB:', mongoError);
      }

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

// Server-generated PDF (pdfkit, no headless browser) of the same tables as
// /last-declaration-results, colored like the Excel export, with a
// letterhead (logo + company name + generation date) and footer (company
// name) — replaces the earlier window.print()-based "Exporter PDF", which
// depended on the visiting browser's own print settings to render colors/
// backgrounds consistently.
app.get('/last-declaration-pdf', async (_req, res) => {
  if (!lastDeclaration) {
    res.status(404).json({ success: false, error: 'Aucune déclaration générée pour le moment.' });
    return;
  }
  const doc = generateDeclarationPdf(lastDeclaration, await getAppSettings(await getSettingsCollection()));
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', 'attachment; filename="Declaration.pdf"');
  doc.pipe(res);
  doc.end();
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

app.get('/superadmin/dashboard', requireSuperAdmin, async (_req, res) => {
  let transactionCount = 0;
  let countryCounts: CountryProductCount[] = [];
  try {
    const mongoDb = await getMongoDb();
    const collection = mongoDb.collection<TransactionDocument>(TRANSACTIONS_COLLECTION);
    transactionCount = await countTransactions(collection);
    countryCounts = await getCountryProductCounts(collection);
  } catch (mongoError) {
    console.error('Failed to reach MongoDB for the dashboard overview:', mongoError);
  }
  const [users, settings] = await Promise.all([
    listUsers(await getUsersCollection()),
    getAppSettings(await getSettingsCollection()),
  ]);
  res.send(renderSuperAdminOverview(users, transactionCount, settings, countryCounts));
});

// Lets a superadmin generate a declaration without leaving the sidebar
// dashboard — reuses the exact same /generate, /download,
// /last-declaration-cost-summary, and /last-declaration-results endpoints
// as the standalone admin tool at "/", just wrapped in the sidebar shell.
app.get('/superadmin/generate', requireSuperAdmin, async (_req, res) => {
  res.send(renderSuperAdminGenerate(await getAppSettings(await getSettingsCollection())));
});

app.get('/superadmin/users', requireSuperAdmin, async (req, res) => {
  const [users, settings] = await Promise.all([
    listUsers(await getUsersCollection()),
    getAppSettings(await getSettingsCollection()),
  ]);
  res.send(renderSuperAdminUsers(users, req.session!.userId, settings));
});

app.get('/superadmin/costs', requireSuperAdmin, async (req, res) => {
  // Reads from MongoDB (most recent transaction across all admins), not
  // the in-memory `lastDeclaration` — that state is per-process and wiped
  // on every restart/redeploy, which made this page go blank even though
  // declarations had already been generated before the restart.
  let collection: Collection<TransactionDocument>;
  try {
    const mongoDb = await getMongoDb();
    collection = mongoDb.collection<TransactionDocument>(TRANSACTIONS_COLLECTION);
  } catch (mongoError) {
    console.error('Failed to reach MongoDB for Coût de produit:', mongoError);
    res.status(503).send(
      renderSuperAdminPlaceholder(
        'Coût de produit',
        "Impossible de se connecter à la base de données pour le moment. Réessayez plus tard.",
        DEFAULT_APP_SETTINGS
      )
    );
    return;
  }

  const mostRecent = await getMostRecentTransaction(collection);
  const settings = await getAppSettings(await getSettingsCollection());
  if (!mostRecent) {
    res.send(
      renderSuperAdminPlaceholder(
        'Coût de produit',
        "Aucune déclaration n'a encore été générée sur l'application. Le coût par produit s'affichera ici après une génération.",
        settings
      )
    );
    return;
  }
  const searchQuery = typeof req.query.q === 'string' ? req.query.q.trim() : '';
  const searchResults = searchQuery
    ? await searchTransactionsByRedevable(collection, searchQuery)
    : undefined;
  res.send(renderSuperAdminCosts(mostRecent, settings, searchQuery, searchResults));
});

app.get('/superadmin/settings', requireSuperAdmin, async (req, res) => {
  const settings = await getAppSettings(await getSettingsCollection());
  res.send(renderSuperAdminSettings(settings, undefined, undefined, req.session!.username));
});

app.post(
  '/superadmin/settings',
  requireSuperAdmin,
  (req, res, next) => {
    logoUpload.single('logo')(req, res, async (err) => {
      if (err) {
        // multer throws for oversized files (LIMIT_FILE_SIZE) — surface it
        // the same way other form errors on this page are shown, instead
        // of an unhandled 500.
        const settings = await getAppSettings(await getSettingsCollection());
        res.status(400).send(
          renderSuperAdminSettings(
            settings,
            err.code === 'LIMIT_FILE_SIZE'
              ? 'Le logo dépasse la taille maximale autorisée (2 Mo).'
              : "Échec de l'envoi du fichier.",
            undefined,
            req.session!.username
          )
        );
        return;
      }
      next();
    });
  },
  async (req, res) => {
    const { companyName, brandColor, fontFamily, contactEmail, contactWhatsapp, removeLogo } =
      req.body as {
        companyName?: string;
        brandColor?: string;
        fontFamily?: string;
        contactEmail?: string;
        contactWhatsapp?: string;
        removeLogo?: string;
      };

    const settingsCollection = await getSettingsCollection();

    if (brandColor && !isValidHexColor(brandColor)) {
      res
        .status(400)
        .send(
          renderSuperAdminSettings(
            await getAppSettings(settingsCollection),
            'Couleur invalide.',
            undefined,
            req.session!.username
          )
        );
      return;
    }
    const allowedFonts = new Set(FONT_OPTIONS.map((opt) => opt.value));
    if (fontFamily && !allowedFonts.has(fontFamily)) {
      res
        .status(400)
        .send(
          renderSuperAdminSettings(
            await getAppSettings(settingsCollection),
            'Police invalide.',
            undefined,
            req.session!.username
          )
        );
      return;
    }

    const logoFile = req.file;
    if (logoFile && !LOGO_ALLOWED_MIME_TYPES.has(logoFile.mimetype)) {
      res
        .status(400)
        .send(
          renderSuperAdminSettings(
            await getAppSettings(settingsCollection),
            'Format de logo non supporté (PNG, JPEG, WEBP ou SVG uniquement).',
            undefined,
            req.session!.username
          )
        );
      return;
    }

    const updated = await updateAppSettings(settingsCollection, {
      companyName: companyName?.trim() ? companyName.trim() : null,
      brandColor: brandColor || null,
      fontFamily: fontFamily || null,
      contactEmail: contactEmail?.trim() ? contactEmail.trim() : null,
      contactWhatsapp: contactWhatsapp?.trim() ? contactWhatsapp.trim() : null,
      ...(logoFile
        ? { logoDataUri: `data:${logoFile.mimetype};base64,${logoFile.buffer.toString('base64')}` }
        : removeLogo === '1'
          ? { logoDataUri: null }
          : {}),
    });

    res.send(
      renderSuperAdminSettings(updated, undefined, 'Réglages enregistrés.', req.session!.username)
    );
  }
);

app.post('/superadmin/settings/credentials', requireSuperAdmin, async (req, res) => {
  const { username, newPassword } = req.body as { username?: string; newPassword?: string };
  const trimmedUsername = username?.trim();
  const settingsCollection = await getSettingsCollection();

  const renderWithError = async (message: string) => {
    res
      .status(400)
      .send(
        renderSuperAdminSettings(
          await getAppSettings(settingsCollection),
          undefined,
          undefined,
          req.session!.username,
          message
        )
      );
  };

  if (!trimmedUsername) {
    await renderWithError("Le nom d'utilisateur est requis.");
    return;
  }
  if (newPassword && newPassword.length < 6) {
    await renderWithError('Le mot de passe doit contenir au moins 6 caractères.');
    return;
  }

  try {
    const usersCollection = await getUsersCollection();
    if (trimmedUsername !== req.session!.username) {
      await updateUsername(usersCollection, req.session!.userId, trimmedUsername);
      req.session!.username = trimmedUsername;
    }
    if (newPassword) {
      await updatePassword(usersCollection, req.session!.userId, newPassword);
    }
  } catch (error) {
    console.error('Failed to update superadmin credentials:', error);
    await renderWithError('Ce nom d’utilisateur est déjà utilisé.');
    return;
  }

  res.send(
    renderSuperAdminSettings(
      await getAppSettings(settingsCollection),
      undefined,
      undefined,
      req.session!.username,
      undefined,
      'Identifiants mis à jour.'
    )
  );
});

app.post('/superadmin/users', requireSuperAdmin, async (req, res) => {
  const { username, password, role } = req.body as {
    username?: string;
    password?: string;
    role?: string;
  };

  const renderWithError = async (message: string) => {
    res
      .status(400)
      .send(
        renderSuperAdminUsers(
          await listUsers(await getUsersCollection()),
          req.session!.userId,
          await getAppSettings(await getSettingsCollection()),
          message
        )
      );
  };

  if (!username || !password) {
    await renderWithError('Identifiant et mot de passe sont requis.');
    return;
  }
  if (role !== 'admin' && role !== 'superadmin') {
    await renderWithError('Rôle invalide.');
    return;
  }

  try {
    await createUser(await getUsersCollection(), username, password, role as UserRole);
    res.redirect('/superadmin/users');
  } catch (error) {
    // MongoDB's unique index on username throws an E11000 error (code
    // 11000) on a duplicate — the only realistic failure mode here, since
    // username/password/role are already validated above.
    if (error instanceof Error && 'code' in error && error.code === 11000) {
      await renderWithError(`L'identifiant « ${username} » est déjà utilisé.`);
      return;
    }
    throw error;
  }
});

function setDisabledAndRedirect(disabled: boolean) {
  return async (req: express.Request, res: express.Response) => {
    const targetId = String(req.params.id);
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
            await listUsers(await getUsersCollection()),
            req.session!.userId,
            await getAppSettings(await getSettingsCollection()),
            'Vous ne pouvez pas désactiver votre propre compte.'
          )
        );
      return;
    }
    await setUserDisabled(await getUsersCollection(), targetId, disabled);
    res.redirect('/superadmin/users');
  };
}

app.post('/superadmin/users/:id/disable', requireSuperAdmin, setDisabledAndRedirect(true));
app.post('/superadmin/users/:id/enable', requireSuperAdmin, setDisabledAndRedirect(false));

app.post('/superadmin/users/:id/update', requireSuperAdmin, async (req, res) => {
  const targetId = String(req.params.id);
  const { username, newPassword } = req.body as { username?: string; newPassword?: string };
  const trimmedUsername = username?.trim();

  const renderWithError = async (message: string) => {
    res
      .status(400)
      .send(
        renderSuperAdminUsers(
          await listUsers(await getUsersCollection()),
          req.session!.userId,
          await getAppSettings(await getSettingsCollection()),
          message
        )
      );
  };

  if (!trimmedUsername) {
    await renderWithError("Le nom d'utilisateur est requis.");
    return;
  }
  if (newPassword && newPassword.length < 6) {
    await renderWithError('Le mot de passe doit contenir au moins 6 caractères.');
    return;
  }

  try {
    const usersCollection = await getUsersCollection();
    await updateUsername(usersCollection, targetId, trimmedUsername);
    if (newPassword) {
      await updatePassword(usersCollection, targetId, newPassword);
    }
    res.redirect('/superadmin/users');
  } catch (error) {
    if (error instanceof Error && 'code' in error && error.code === 11000) {
      await renderWithError(`L'identifiant « ${trimmedUsername} » est déjà utilisé.`);
      return;
    }
    throw error;
  }
});

app.post('/superadmin/users/:id/delete', requireSuperAdmin, async (req, res) => {
  const targetId = String(req.params.id);
  // Same self-lockout reasoning as disable: the UI already hides the button
  // for one's own row, but a direct POST must still be rejected server-side.
  if (targetId === req.session!.userId) {
    res
      .status(400)
      .send(
        renderSuperAdminUsers(
          await listUsers(await getUsersCollection()),
          req.session!.userId,
          await getAppSettings(await getSettingsCollection()),
          'Vous ne pouvez pas supprimer votre propre compte.'
        )
      );
    return;
  }
  await deleteUser(await getUsersCollection(), targetId);
  res.redirect('/superadmin/users');
});

// Default matches Coolify's default "Ports Exposes" (3000) so the app works
// out of the box even if the PORT environment variable doesn't reach the
// running container (e.g. a Coolify env var scoped to build-time only).
const port = Number(process.env.PORT ?? 3000);
// Attempt the superadmin seed before accepting traffic, but don't block
// startup on it — a transient Mongo outage at boot shouldn't stop the
// process from listening (Coolify's health check needs a live port);
// routes that need Mongo simply fail per-request until it's reachable,
// same as the rest of the app already handles it.
bootstrap().catch((error) => {
  console.error('Failed to seed initial superadmin (is MONGODB_URI reachable?):', error);
});
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});

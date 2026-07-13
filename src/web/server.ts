import { mkdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import 'dotenv/config';
import express from 'express';
import multer from 'multer';
import { extractDocumentText } from '../ocr/documentTextExtractor.js';
import { detectAndParsePair } from '../parser/detectAndParsePair.js';
import { mergeDeclaration } from '../merge/declarationMerger.js';
import { validateArticle } from '../domain/validators.js';
import { generateCombinedExcel } from '../excel/combinedExcelGenerator.js';
import { renderResultsPage } from './renderResultsPage.js';
import { checkCredentials, createSession, requireAuth, setSessionCookie } from './auth.js';
import type { Declaration } from '../domain/types.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const UPLOAD_DIR = path.join(PROJECT_ROOT, '.tmp-uploads');
const OUTPUT_DIR = path.join(PROJECT_ROOT, '.tmp-output');
// These don't exist yet on a fresh checkout/container — created locally by
// hand during earlier development, but a fresh deployment has neither.
mkdirSync(UPLOAD_DIR, { recursive: true });
mkdirSync(OUTPUT_DIR, { recursive: true });

const uploadHtml = readFileSync(path.join(__dirname, 'views/upload.html'), 'utf-8');
const loginHtml = readFileSync(path.join(__dirname, 'views/login.html'), 'utf-8');

// Single-user local tool (no auth, no concurrency handling) — holds the most
// recently generated declaration in memory so /download and the results
// preview can reference it without a database.
let lastDeclaration: Declaration | undefined;
const combinedFilePath = path.join(OUTPUT_DIR, 'Declaration.xlsx');

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
const app = express();
app.use(express.urlencoded({ extended: false }));

app.get('/login', (_req, res) => {
  res.send(loginHtml.replace('{{ERROR_BLOCK}}', ''));
});

app.post('/login', (req, res) => {
  const { username, password } = req.body as { username?: string; password?: string };

  if (!username || !password || !checkCredentials(username, password)) {
    const errorBlock = '<div class="error">Identifiant ou mot de passe incorrect.</div>';
    res.status(401).send(loginHtml.replace('{{ERROR_BLOCK}}', errorBlock));
    return;
  }

  const sessionId = createSession();
  setSessionCookie(res, sessionId);
  res.redirect('/');
});

app.use(requireAuth);

app.get('/', (_req, res) => {
  res.send(uploadHtml.replace('{{ERROR_BLOCK}}', ''));
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

      await generateCombinedExcel(declaration, combinedFilePath);

      lastDeclaration = declaration;
      res.download(combinedFilePath, 'Declaration.xlsx');
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

app.get('/download', (_req, res) => {
  res.download(combinedFilePath, 'Declaration.xlsx');
});

// Default matches Coolify's default "Ports Exposes" (3000) so the app works
// out of the box even if the PORT environment variable doesn't reach the
// running container (e.g. a Coolify env var scoped to build-time only).
const port = Number(process.env.PORT ?? 3000);
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});

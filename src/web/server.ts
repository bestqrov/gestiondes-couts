import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import multer from 'multer';
import { extractDocumentText } from '../ocr/documentTextExtractor.js';
import { parseLiquidation } from '../parser/liquidation/liquidationParser.js';
import { parseDum } from '../parser/dum/dumParser.js';
import { mergeDeclaration } from '../merge/declarationMerger.js';
import { validateArticle } from '../domain/validators.js';
import { generateArticleSummaryExcel } from '../excel/articleSummaryExcelGenerator.js';
import { generateUnitLevelExcel } from '../excel/unitLevelExcelGenerator.js';
import { renderResultsPage } from './renderResultsPage.js';
import type { Declaration } from '../domain/types.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const UPLOAD_DIR = path.join(PROJECT_ROOT, '.tmp-uploads');
const OUTPUT_DIR = path.join(PROJECT_ROOT, '.tmp-output');

const uploadHtml = readFileSync(path.join(__dirname, 'views/upload.html'), 'utf-8');

// Single-user local tool (no auth, no concurrency handling) — holds the most
// recently generated declaration in memory so /download and the results
// preview can reference it without a database.
let lastDeclaration: Declaration | undefined;
const file1Path = path.join(OUTPUT_DIR, 'File1-ArticleSummary.xlsx');
const file2Path = path.join(OUTPUT_DIR, 'File2-UnitLevelDetail.xlsx');

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

      const liquidation = parseLiquidation(liquidationOcr.text);
      const dum = parseDum(dumOcr.text);
      const declaration = mergeDeclaration(liquidation, dum);
      for (const article of declaration.articles) {
        validateArticle(article);
      }

      await generateArticleSummaryExcel(declaration, file1Path);
      await generateUnitLevelExcel(declaration, file2Path);

      lastDeclaration = declaration;
      res.json({ success: true });
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

app.get('/download/file1', (_req, res) => {
  res.download(file1Path, 'File1-ArticleSummary.xlsx');
});

app.get('/download/file2', (_req, res) => {
  res.download(file2Path, 'File2-UnitLevelDetail.xlsx');
});

const port = Number(process.env.PORT ?? 4310);
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});

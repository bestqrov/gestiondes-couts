import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { generateArticleSummaryExcel } from '../../src/excel/articleSummaryExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return mergeDeclaration(liquidation, dum);
}

describe('generateArticleSummaryExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes one row per article with the required columns, from the real merged declaration', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('article-summary');
    tempDir = dir;

    await generateArticleSummaryExcel(declaration, filePath, {
      companyName: 'ACME LOGISTICS SARL',
      brandColor: '#4f46e5',
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // Row 1: company name (merged, centered, bold). Row 2: document
    // reference. Row 3: the actual column header row.
    const titleRow = sheet.getRow(1);
    expect(titleRow.getCell(1).value).toBe('ACME LOGISTICS SARL');
    expect(sheet.getCell(1, 5).isMerged).toBe(true);
    expect(sheet.getCell(1, 5).master.address).toBe('A1'); // merged across all 5 columns

    const subtitleRow = sheet.getRow(2);
    expect(subtitleRow.getCell(1).value).toBe(`Déclaration ${declaration.code} — ${declaration.redevable}`);

    const headerRow = sheet.getRow(3);
    expect(headerRow.getCell(1).value).toBe('Nom Article');
    expect(headerRow.getCell(2).value).toBe('HSC');
    expect(headerRow.getCell(3).value).toBe('Pays');
    expect(headerRow.getCell(4).value).toBe('Valeur déclarée');
    expect(headerRow.getCell(5).value).toBe('Unité (Quantity)');

    // 2 title rows + 1 header row + 2 articles = 5 rows total
    expect(sheet.rowCount).toBe(5);

    const row1 = sheet.getRow(4);
    expect(row1.getCell(1).value).toBe('T-SHIRT');
    expect(row1.getCell(2).value).toBe('6109100010');
    expect(row1.getCell(3).value).toBe('ITALIE');
    expect(row1.getCell(4).value).toBeCloseTo(27147.0, 1);
    expect(row1.getCell(5).value).toBeCloseTo(354.0, 1);

    const row2 = sheet.getRow(5);
    expect(row2.getCell(3).value).toBe('BANGLADESH');
    expect(row2.getCell(5).value).toBeCloseTo(200.0, 1);
  });

  it('embeds the configured logo image on the left of the title row, right-aligning the company name next to it', async () => {
    // A minimal valid 10x10 PNG, base64-encoded as a data: URI — the same
    // shape a real uploaded logo takes (see LOGO_ALLOWED_MIME_TYPES).
    const logoDataUri =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAEklEQVR4nGP4z8CAB+GTG8HSALfKY52fTcuYAAAAAElFTkSuQmCC';
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('article-summary-logo');
    tempDir = dir;

    await generateArticleSummaryExcel(declaration, filePath, {
      companyName: 'ACME LOGISTICS SARL',
      brandColor: '#4f46e5',
      logoDataUri,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(workbook.model.media).toHaveLength(1);
    expect(workbook.model.media[0].type).toBe('image');

    const titleRow = sheet.getRow(1);
    expect(titleRow.getCell(1).value).toBe('ACME LOGISTICS SARL');
    expect(titleRow.getCell(1).alignment?.horizontal).toBe('right');
  });

  it('falls back to a generic company name when none is configured', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('article-summary-no-branding');
    tempDir = dir;

    await generateArticleSummaryExcel(declaration, filePath, { companyName: null, brandColor: null, logoDataUri: null });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(sheet.getRow(1).getCell(1).value).toBe('Gestion des Coûts');
  });

  it('writes only the title/header rows when the declaration has no articles', async () => {
    const emptyDeclaration = {
      code: '000000',
      redevable: 'EMPTY CO',
      benNumero: '1',
      articles: [],
    };
    const { filePath, dir } = createTempXlsxPath('article-summary-empty');
    tempDir = dir;

    await generateArticleSummaryExcel(emptyDeclaration, filePath, {
      companyName: null,
      brandColor: null,
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(sheet.getRow(3).getCell(1).value).toBe('Nom Article');
    expect(sheet.rowCount).toBe(3);
  });
});

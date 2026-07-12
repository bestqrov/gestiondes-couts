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

    await generateArticleSummaryExcel(declaration, filePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    const headerRow = sheet.getRow(1);
    expect(headerRow.getCell(1).value).toBe('Nom Article');
    expect(headerRow.getCell(2).value).toBe('HSC');
    expect(headerRow.getCell(3).value).toBe('Pays');
    expect(headerRow.getCell(4).value).toBe('Valeur déclarée');
    expect(headerRow.getCell(5).value).toBe('Unité (Quantity)');

    // sheet.rowCount includes the header row, so 2 articles -> 3 rows total
    expect(sheet.rowCount).toBe(3);

    const row1 = sheet.getRow(2);
    expect(row1.getCell(1).value).toBe('T-SHIRT');
    expect(row1.getCell(2).value).toBe('6109100010');
    expect(row1.getCell(3).value).toBe('ITALIE');
    expect(row1.getCell(4).value).toBeCloseTo(27147.0, 1);
    expect(row1.getCell(5).value).toBeCloseTo(354.0, 1);

    const row2 = sheet.getRow(3);
    expect(row2.getCell(3).value).toBe('BANGLADESH');
    expect(row2.getCell(5).value).toBeCloseTo(200.0, 1);
  });

  it('writes only the header row when the declaration has no articles', async () => {
    const emptyDeclaration = {
      code: '000000',
      redevable: 'EMPTY CO',
      benNumero: '1',
      articles: [],
    };
    const { filePath, dir } = createTempXlsxPath('article-summary-empty');
    tempDir = dir;

    await generateArticleSummaryExcel(emptyDeclaration, filePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(sheet.getRow(1).getCell(1).value).toBe('Nom Article');
    expect(sheet.rowCount).toBe(1);
  });
});

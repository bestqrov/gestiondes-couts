import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { generateUnitLevelExcel } from '../../src/excel/unitLevelExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { Declaration } from '../../src/domain/types.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration(): Declaration {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return mergeDeclaration(liquidation, dum);
}

describe('generateUnitLevelExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes one row per physical unit, with tax columns that reconcile exactly to the source montants', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('unit-level');
    tempDir = dir;

    await generateUnitLevelExcel(declaration, filePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    const headerRow = sheet.getRow(1);
    expect(headerRow.getCell(1).value).toBe('Nom Article');
    expect(headerRow.getCell(2).value).toBe('HSC');
    expect(headerRow.getCell(3).value).toBe('Serial Number');
    // union of tax codes across both articles, sorted: 000110, 002109, 007217
    expect(headerRow.getCell(4).value).toBe('000110');
    expect(headerRow.getCell(5).value).toBe('002109');
    expect(headerRow.getCell(6).value).toBe('007217');

    // article 1: 354 units, article 2: 200 units -> 554 data rows + 1 header = 555
    expect(sheet.rowCount).toBe(555);

    // first row of article 1
    const firstRow = sheet.getRow(2);
    expect(firstRow.getCell(1).value).toBe('T-SHIRT');
    expect(firstRow.getCell(3).value).toBe(1);

    // last row of article 1 (row 355 = header + 354 units), first row of article 2 resets serial number
    const lastRowArticle1 = sheet.getRow(355);
    expect(lastRowArticle1.getCell(3).value).toBe(354);
    const firstRowArticle2 = sheet.getRow(356);
    expect(firstRowArticle2.getCell(3).value).toBe(1);
    expect(firstRowArticle2.getCell(1).value).toBe('T-SHIRT');

    // Reconciliation: sum each tax column across article 1's 354 rows (rows 2-355)
    // against the known source montants from the Liquidation fixture:
    // 000110 = 0.00, 002109 = 5443.00, 007217 = 68.00
    let sum000110 = 0;
    let sum002109 = 0;
    let sum007217 = 0;
    for (let rowNum = 2; rowNum <= 355; rowNum++) {
      const row = sheet.getRow(rowNum);
      sum000110 += Number(row.getCell(4).value);
      sum002109 += Number(row.getCell(5).value);
      sum007217 += Number(row.getCell(6).value);
    }
    expect(sum000110).toBeCloseTo(0.0, 2);
    expect(sum002109).toBeCloseTo(5443.0, 2);
    expect(sum007217).toBeCloseTo(68.0, 2);
  });

  it('throws when an article quantite is not a whole number', async () => {
    const declaration = loadRealDeclaration();
    const brokenDeclaration: Declaration = {
      ...declaration,
      articles: declaration.articles.map((a, i) => (i === 0 ? { ...a, quantite: 354.5 } : a)),
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-broken');
    tempDir = dir;

    await expect(generateUnitLevelExcel(brokenDeclaration, filePath)).rejects.toThrow();
  });

  it('zero-fills tax columns for articles that lack a given code, across the actual generator (not just the union helper)', async () => {
    const declaration: Declaration = {
      code: '111111',
      redevable: 'DIVERGENT CODES CO',
      benNumero: '1',
      articles: [
        {
          numero: 1,
          hsCode: '1111111111',
          nomArticle: 'ARTICLE A',
          pays: 'ITALIE',
          paysCode: 'IT',
          valeurDeclaree: 100,
          quantite: 3,
          unite: 'U',
          taxes: [
            { code: '000110', assiette: 100, taux: 0, montant: 0 },
            { code: '007217', assiette: 100, taux: 0.25, montant: 3 },
          ],
          totalArticle: 3,
        },
        {
          numero: 2,
          hsCode: '2222222222',
          nomArticle: 'ARTICLE B',
          pays: 'BANGLADESH',
          paysCode: 'BD',
          valeurDeclaree: 50,
          quantite: 2,
          unite: 'U',
          taxes: [{ code: '002109', assiette: 50, taux: 20, montant: 10 }],
          totalArticle: 10,
        },
      ],
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-divergent-codes');
    tempDir = dir;

    await generateUnitLevelExcel(declaration, filePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // header: Nom Article | HSC | Serial Number | 000110 | 002109 | 007217 (sorted union)
    const headerRow = sheet.getRow(1);
    expect(headerRow.getCell(4).value).toBe('000110');
    expect(headerRow.getCell(5).value).toBe('002109');
    expect(headerRow.getCell(6).value).toBe('007217');

    // article A: 3 rows (rows 2-4), has 000110 and 007217 but NOT 002109 -> 002109 column must be 0
    for (let rowNum = 2; rowNum <= 4; rowNum++) {
      const row = sheet.getRow(rowNum);
      expect(Number(row.getCell(5).value)).toBe(0); // 002109 column, article A doesn't have this code
    }
    // article A's 007217 column (montant=3 across 3 units) should reconcile to 3
    let sumA007217 = 0;
    for (let rowNum = 2; rowNum <= 4; rowNum++) {
      sumA007217 += Number(sheet.getRow(rowNum).getCell(6).value);
    }
    expect(sumA007217).toBeCloseTo(3, 2);

    // article B: 2 rows (rows 5-6), has ONLY 002109 -> 000110 and 007217 columns must be 0
    for (let rowNum = 5; rowNum <= 6; rowNum++) {
      const row = sheet.getRow(rowNum);
      expect(Number(row.getCell(4).value)).toBe(0); // 000110 column, article B doesn't have this code
      expect(Number(row.getCell(6).value)).toBe(0); // 007217 column, article B doesn't have this code
    }
    // article B's 002109 column (montant=10 across 2 units) should reconcile to 10
    let sumB002109 = 0;
    for (let rowNum = 5; rowNum <= 6; rowNum++) {
      sumB002109 += Number(sheet.getRow(rowNum).getCell(5).value);
    }
    expect(sumB002109).toBeCloseTo(10, 2);
  });
});

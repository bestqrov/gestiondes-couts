import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { addPerArticleUnitSheets } from '../../src/excel/unitLevelExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { Declaration } from '../../src/domain/types.js';

function makeDeclaration(): Declaration {
  return {
    code: '500001',
    redevable: 'GLOBAL TRADE LOGISTICS SARL',
    benNumero: '501',
    articles: [
      {
        numero: 1,
        hsCode: '6109100010',
        nomArticle: 'T-SHIRT',
        pays: 'ITALIE',
        paysCode: 'IT',
        valeurDeclaree: 27147,
        quantite: 3,
        unite: 'NOMBRE',
        taxes: [
          { code: '000110', assiette: 27147, taux: 0, montant: 0 },
          { code: '007217', assiette: 27147, taux: 0.25, montant: 3 },
        ],
        totalArticle: 3,
      },
      {
        numero: 2,
        hsCode: '8471300000',
        nomArticle: 'ORDINATEUR',
        pays: 'ESPAGNE',
        paysCode: 'ES',
        valeurDeclaree: 9500,
        quantite: 2,
        unite: 'NOMBRE',
        taxes: [{ code: '002109', assiette: 9500, taux: 20, montant: 10 }],
        totalArticle: 10,
      },
    ],
  };
}

async function buildWorkbook(declaration: Declaration, filePath: string) {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: filePath, useStyles: false });
  addPerArticleUnitSheets(workbook, declaration);
  await workbook.commit();

  const reader = new ExcelJS.Workbook();
  await reader.xlsx.readFile(filePath);
  return reader;
}

describe('addPerArticleUnitSheets', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('creates one sheet per article, named by article number and product name', async () => {
    const declaration = makeDeclaration();
    const { filePath, dir } = createTempXlsxPath('per-article');
    tempDir = dir;

    const workbook = await buildWorkbook(declaration, filePath);

    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual([
      '1-T-SHIRT',
      '2-ORDINATEUR',
    ]);
  });

  it("only includes each article's own tax codes as columns, not the declaration-wide union", async () => {
    const declaration = makeDeclaration();
    const { filePath, dir } = createTempXlsxPath('per-article-columns');
    tempDir = dir;

    const workbook = await buildWorkbook(declaration, filePath);

    const sheet1 = workbook.getWorksheet('1-T-SHIRT')!;
    const header1 = sheet1.getRow(1);
    expect([1, 2, 3, 4, 5, 6].map((col) => header1.getCell(col).value)).toEqual([
      'Nom Article',
      'HSC',
      'Serial Number',
      'Valeur Déclarée',
      '000110',
      '007217',
    ]);

    const sheet2 = workbook.getWorksheet('2-ORDINATEUR')!;
    const header2 = sheet2.getRow(1);
    expect([1, 2, 3, 4, 5].map((col) => header2.getCell(col).value)).toEqual([
      'Nom Article',
      'HSC',
      'Serial Number',
      'Valeur Déclarée',
      '002109',
    ]);
  });

  it("gives every row of an article the same Valeur Déclarée / Unité value", async () => {
    const declaration = makeDeclaration();
    const { filePath, dir } = createTempXlsxPath('per-article-valeur-declaree');
    tempDir = dir;

    const workbook = await buildWorkbook(declaration, filePath);

    const sheet1 = workbook.getWorksheet('1-T-SHIRT')!;
    // article 1: valeurDeclaree 27147, quantite 3 -> 9049 per unit
    for (let rowNum = 2; rowNum <= 4; rowNum++) {
      expect(Number(sheet1.getRow(rowNum).getCell(4).value)).toBeCloseTo(27147 / 3, 4);
    }

    const sheet2 = workbook.getWorksheet('2-ORDINATEUR')!;
    // article 2: valeurDeclaree 9500, quantite 2 -> 4750 per unit
    for (let rowNum = 2; rowNum <= 3; rowNum++) {
      expect(Number(sheet2.getRow(rowNum).getCell(4).value)).toBeCloseTo(9500 / 2, 4);
    }
  });

  it('writes exactly quantite rows per article, with correct serial numbers', async () => {
    const declaration = makeDeclaration();
    const { filePath, dir } = createTempXlsxPath('per-article-rows');
    tempDir = dir;

    const workbook = await buildWorkbook(declaration, filePath);

    const sheet1 = workbook.getWorksheet('1-T-SHIRT')!;
    expect(sheet1.rowCount).toBe(4); // header + 3 units
    expect(sheet1.getRow(2).getCell(3).value).toBe(1);
    expect(sheet1.getRow(4).getCell(3).value).toBe(3);

    const sheet2 = workbook.getWorksheet('2-ORDINATEUR')!;
    expect(sheet2.rowCount).toBe(3); // header + 2 units
  });

  it('gives two articles with the same product name distinct sheet names', async () => {
    const declaration = makeDeclaration();
    declaration.articles[1].nomArticle = 'T-SHIRT'; // collide with article 1's name
    const { filePath, dir } = createTempXlsxPath('per-article-collision');
    tempDir = dir;

    const workbook = await buildWorkbook(declaration, filePath);

    const names = workbook.worksheets.map((sheet) => sheet.name);
    expect(new Set(names).size).toBe(names.length); // all unique
    expect(names).toContain('1-T-SHIRT');
    expect(names).toContain('2-T-SHIRT');
  });
});

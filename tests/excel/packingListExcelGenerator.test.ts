import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { addPackingListSheet } from '../../src/excel/packingListExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { PackingListRow } from '../../src/parser/packingList/packingListParser.js';
import type { Declaration } from '../../src/domain/types.js';

const SAMPLE_DECLARATION: Declaration = {
  code: '309536',
  redevable: 'MED AFRICA LOGISTICS',
  benNumero: '136',
  articles: [],
  ordonnancementTaxes: [],
  numeroEnregistrement: null,
  dateArrivee: null,
  donneesComptables: null,
  titreTransport: null,
};

const DECLARATION_WITH_TAXES: Declaration = {
  ...SAMPLE_DECLARATION,
  articles: [
    {
      numero: 1,
      hsCode: '61044300',
      nomArticle: 'VESTITO A FASCIA CORTO IN TULLE',
      pays: 'CHINE',
      paysCode: 'CN',
      valeurDeclaree: 172.98,
      quantite: 18,
      unite: 'U',
      totalArticle: 172.98,
      poidsNet: 12,
      taxes: [
        { code: '000110', assiette: 172.98, taux: 2.5, montant: 4.32 },
        { code: '002109', assiette: 172.98, taux: 20, montant: 34.6 },
      ],
    },
  ],
};

const SAMPLE_ROWS: PackingListRow[] = [
  {
    item: 'AB0141DOAY16',
    description: 'VESTITO A FASCIA CORTO IN TULLE',
    color: 'PK2 PINK MEDIUM',
    pieces: 18,
    unit: 9.61,
    total: 172.98,
    origin: 'CHINA',
    hsCode: '61044300',
  },
  {
    item: 'BD0015DOAY16',
    description: 'BODY A COSTE',
    color: 'WH2 WHITE',
    pieces: 18,
    unit: 4.8,
    total: 86.4,
    origin: 'BANGLADESH',
    hsCode: '61142000',
  },
];

describe('addPackingListSheet', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes one row per packing-list row, in the given order, with the source column headers', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      SAMPLE_DECLARATION,
      { companyName: 'ACME SARL', brandColor: '#4f46e5', logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    // Rows 1-3: letterhead. Row 4: header.
    expect(sheet.getRow(1).getCell(1).value).toBe('ACME SARL');
    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(1).value).toBe('item');
    expect(headerRow.getCell(2).value).toBe('DESCRIPTION');
    expect(headerRow.getCell(3).value).toBe('color');
    expect(headerRow.getCell(4).value).toBe('pieces');
    expect(headerRow.getCell(5).value).toBe('unit');
    expect(headerRow.getCell(6).value).toBe('total');
    expect(headerRow.getCell(7).value).toBe('origin');
    expect(headerRow.getCell(8).value).toBe('HS CODE');

    // 3 title rows + header + 2 rows = 6
    expect(sheet.rowCount).toBe(6);

    const row1 = sheet.getRow(5);
    expect(row1.getCell(1).value).toBe('AB0141DOAY16');
    expect(row1.getCell(7).value).toBe('CHINA');
    expect(Number(row1.getCell(6).value)).toBeCloseTo(172.98, 2);

    // Original order preserved — BD0015DOAY16 second, not resorted.
    const row2 = sheet.getRow(6);
    expect(row2.getCell(1).value).toBe('BD0015DOAY16');
  });

  it('colors the header by column group: identity, quantity, value', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-colors');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      SAMPLE_DECLARATION,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;
    const headerRow = sheet.getRow(4);

    const identityArgb = 'FF4F46E5';
    const quantityArgb = 'FF0891B2';
    const valueArgb = 'FF059669';

    expect((headerRow.getCell(1).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(identityArgb); // item
    expect((headerRow.getCell(4).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(quantityArgb); // pieces
    expect((headerRow.getCell(5).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(valueArgb); // unit
    expect((headerRow.getCell(6).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(valueArgb); // total
    expect((headerRow.getCell(7).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(identityArgb); // origin
    expect((headerRow.getCell(8).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(identityArgb); // HS CODE
  });

  it('adds tax columns from the matching article (Nom Article = description), colored orange', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-taxes');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      DECLARATION_WITH_TAXES,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;
    const headerRow = sheet.getRow(4);

    const taxArgb = 'FFD97706';
    expect((headerRow.getCell(9).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(taxArgb);
    expect((headerRow.getCell(10).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(taxArgb);
    expect(headerRow.getCell(9).value).toBe('DTS IMPORT NORMAL');
    expect(headerRow.getCell(10).value).toBe('TVA IMPORT AUTRE PDS');

    // Row 5 = "VESTITO A FASCIA CORTO IN TULLE" — matches the one article.
    const matchedRow = sheet.getRow(5);
    expect(Number(matchedRow.getCell(9).value)).toBeCloseTo(4.32, 2);
    expect(Number(matchedRow.getCell(10).value)).toBeCloseTo(34.6, 2);

    // Row 6 = "BODY A COSTE" — no matching article, taxes default to 0.
    const unmatchedRow = sheet.getRow(6);
    expect(Number(unmatchedRow.getCell(9).value)).toBe(0);
    expect(Number(unmatchedRow.getCell(10).value)).toBe(0);
  });

  it('writes only the letterhead/header when there are no rows', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-empty');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      [],
      SAMPLE_DECLARATION,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    expect(sheet.getRow(4).getCell(1).value).toBe('item');
    expect(sheet.rowCount).toBe(4);
  });
});

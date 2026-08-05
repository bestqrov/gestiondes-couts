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
  dateDeclaration: null,
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
      unitesComplementaires: 1,
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
    // The full HS code is shown (only the 6-digit prefix is used for matching).
    expect(row1.getCell(8).value).toBe('61044300');

    // Original order preserved — BD0015DOAY16 second, not resorted.
    const row2 = sheet.getRow(6);
    expect(row2.getCell(1).value).toBe('BD0015DOAY16');
  });

  it('shows the full HS code even when it has more than 8 digits', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-long-hscode');
    tempDir = dir;

    const rowsWithLongHsCode: PackingListRow[] = [
      { ...SAMPLE_ROWS[0], hsCode: '6104430011' },
    ];

    await addPackingListSheet(
      workbook,
      rowsWithLongHsCode,
      SAMPLE_DECLARATION,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    expect(sheet.getRow(5).getCell(8).value).toBe('6104430011');
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

  it('adds tax columns from the matching article (by HS code prefix), colored orange', async () => {
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

    // Row 5 — HS code 61044300 matches the one article's 610443 prefix.
    const matchedRow = sheet.getRow(5);
    expect(Number(matchedRow.getCell(9).value)).toBeCloseTo(4.32, 2);
    expect(Number(matchedRow.getCell(10).value)).toBeCloseTo(34.6, 2);
    // Every tax column is zero-padded to at least 4 digits before the
    // decimal separator, with no thousands separator.
    expect(matchedRow.getCell(9).numFmt).toBe('0000.00');
    expect(matchedRow.getCell(10).numFmt).toBe('0000.00');

    // Row 6 — HS code 61142000 (prefix 611420) has no matching article, taxes default to 0.
    const unmatchedRow = sheet.getRow(6);
    expect(Number(unmatchedRow.getCell(9).value)).toBe(0);
    expect(Number(unmatchedRow.getCell(10).value)).toBe(0);

    // Somme DD (col 11) = sum of the two tax montants; DD unitaire (col 12)
    // = Somme DD / pieces, formatted to 6 decimal digits.
    expect(headerRow.getCell(11).value).toBe('Somme DD');
    expect(headerRow.getCell(12).value).toBe('DD unitaire');
    expect((headerRow.getCell(11).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(taxArgb);
    expect((headerRow.getCell(12).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(taxArgb);

    expect(Number(matchedRow.getCell(11).value)).toBeCloseTo(38.92, 2);
    expect(Number(matchedRow.getCell(12).value)).toBeCloseTo(38.92 / 18, 6);
    expect(matchedRow.getCell(12).numFmt).toBe('0000.000000');
    // Somme DD (col 11) is zero-padded the same way as the tax columns.
    expect(matchedRow.getCell(11).numFmt).toBe('0000.00');
    // Pieces (col 4) is a plain whole count — no thousands separator, no decimals.
    expect(matchedRow.getCell(4).numFmt).toBe('0');

    expect(Number(unmatchedRow.getCell(11).value)).toBe(0);
    expect(Number(unmatchedRow.getCell(12).value)).toBe(0);
  });

  it('sums taxes across every article sharing the same HS code prefix, and applies the group total to every matching row', async () => {
    const declarationWithTwoVariants: Declaration = {
      ...SAMPLE_DECLARATION,
      articles: [
        {
          numero: 1,
          hsCode: '61044300',
          nomArticle: 'VESTITO PINK',
          pays: 'CHINE',
          paysCode: 'CN',
          valeurDeclaree: 100,
          quantite: 10,
          unite: 'U',
          totalArticle: 100,
          poidsNet: 5,
          unitesComplementaires: 1,
          taxes: [{ code: '000110', assiette: 100, taux: 2.5, montant: 2.5 }],
        },
        {
          numero: 2,
          // Same HS position (610443), different national suffix — treated
          // as the same HS group, same as declarationMerger's own comparison.
          hsCode: '61044399',
          nomArticle: 'VESTITO BLUE',
          pays: 'CHINE',
          paysCode: 'CN',
          valeurDeclaree: 50,
          quantite: 5,
          unite: 'U',
          totalArticle: 50,
          poidsNet: 2,
          unitesComplementaires: 1,
          taxes: [{ code: '000110', assiette: 50, taux: 2.5, montant: 1.25 }],
        },
      ],
    };
    const twoColorVariantRows: PackingListRow[] = [
      { ...SAMPLE_ROWS[0], hsCode: '61044300' },
      { ...SAMPLE_ROWS[0], item: 'AB0141DOAY17', color: 'BL2 BLUE', hsCode: '61044399' },
    ];

    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-hs-group');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      twoColorVariantRows,
      declarationWithTwoVariants,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    // Both rows share HS position 610443, so both get the 2.5 + 1.25 = 3.75 group total.
    expect(Number(sheet.getRow(5).getCell(9).value)).toBeCloseTo(3.75, 2);
    expect(Number(sheet.getRow(6).getCell(9).value)).toBeCloseTo(3.75, 2);
  });

  it('matches by HS code prefix AND country of origin — same HS position, different countries, must not mix tax montants', async () => {
    const declarationWithTwoOrigins: Declaration = {
      ...SAMPLE_DECLARATION,
      articles: [
        {
          numero: 1,
          hsCode: '61044300',
          nomArticle: 'VESTITO PINK',
          pays: 'CHINE',
          paysCode: 'CN',
          valeurDeclaree: 100,
          quantite: 10,
          unite: 'U',
          totalArticle: 100,
          poidsNet: 5,
          unitesComplementaires: 1,
          taxes: [{ code: '000110', assiette: 100, taux: 2.5, montant: 2.5 }],
        },
        {
          numero: 2,
          // Same HS position, different country — must be kept separate from
          // the CHINE article's tax montants above.
          hsCode: '61044300',
          nomArticle: 'VESTITO BLUE',
          pays: 'BANGLADESH',
          paysCode: 'BD',
          valeurDeclaree: 50,
          quantite: 5,
          unite: 'U',
          totalArticle: 50,
          poidsNet: 2,
          unitesComplementaires: 1,
          taxes: [{ code: '000110', assiette: 50, taux: 2.5, montant: 9.99 }],
        },
      ],
    };
    const rowsWithTwoOrigins: PackingListRow[] = [
      { ...SAMPLE_ROWS[0], hsCode: '61044300', origin: 'CHINA' },
      { ...SAMPLE_ROWS[0], item: 'AB0141DOAY18', hsCode: '61044300', origin: 'BANGLADESH' },
    ];

    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-hs-and-origin');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      rowsWithTwoOrigins,
      declarationWithTwoOrigins,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    // Row 5 (origin CHINA, normalized to CHINE) picks up only the CHINE article's 2.5.
    expect(Number(sheet.getRow(5).getCell(9).value)).toBeCloseTo(2.5, 2);
    // Row 6 (origin BANGLADESH) picks up only the BANGLADESH article's 9.99.
    expect(Number(sheet.getRow(6).getCell(9).value)).toBeCloseTo(9.99, 2);
  });

  it('still matches by HS code prefix alone when an HS position has only one country of origin, even if the origin spelling has no known alias (regression: origin-only matching zeroed out real files)', async () => {
    const declarationOneOrigin: Declaration = {
      ...SAMPLE_DECLARATION,
      articles: [
        {
          numero: 1,
          hsCode: '61044300',
          nomArticle: 'VESTITO A FASCIA CORTO IN TULLE',
          // A pays spelling with no entry in countryNames.ts, and no reason
          // to match the packing list's own free-text origin spelling
          // exactly — the point being that HS-only matching shouldn't care.
          pays: 'REPUBLIQUE TCHEQUE',
          paysCode: 'CZ',
          valeurDeclaree: 172.98,
          quantite: 18,
          unite: 'U',
          totalArticle: 172.98,
          poidsNet: 12,
          unitesComplementaires: 1,
          taxes: [{ code: '000110', assiette: 172.98, taux: 2.5, montant: 4.32 }],
        },
      ],
    };
    const rowWithUnmappedOrigin: PackingListRow[] = [
      { ...SAMPLE_ROWS[0], hsCode: '61044300', origin: 'CZECH REPUBLIC' },
    ];

    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-single-origin-unmapped');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      rowWithUnmappedOrigin,
      declarationOneOrigin,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    // Still matches — this HS prefix has exactly one origin in the
    // declaration, so the mismatched spelling never comes into play.
    expect(Number(sheet.getRow(5).getCell(9).value)).toBeCloseTo(4.32, 2);
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

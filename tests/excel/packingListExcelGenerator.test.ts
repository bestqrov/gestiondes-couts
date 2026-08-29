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

    // DESCRIPTION is left-aligned, not centered like the rest of the sheet.
    expect(row1.getCell(2).alignment?.horizontal).toBe('left');
  });

  it('shows the packing list\'s own full HS code, even one with more than 8 digits, when no declaration article matches', async () => {
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

  it('shows the matched declaration article\'s own full HS code, not the packing list\'s own (shorter) one', async () => {
    // The packing list only carries the supplier's 8-digit code ('61044300'),
    // while the declaration's (Global's) is the full, authoritative 10-digit
    // customs code — HS total should show the latter for a matched row, same
    // as Global does, not the packing list's shorter one.
    const declaration: Declaration = {
      ...DECLARATION_WITH_TAXES,
      articles: [{ ...DECLARATION_WITH_TAXES.articles[0], hsCode: '6104430099' }],
    };
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-matched-hscode');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      declaration,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    // Row 5 (SAMPLE_ROWS[0], hsCode '61044300') matches the declaration
    // article by its 6-digit prefix, so it shows the article's full
    // '6104430099', not the packing list's own '61044300'.
    expect(sheet.getRow(5).getCell(8).value).toBe('6104430099');

    // Row 6 (hsCode '61142000') has no matching article, so it still falls
    // back to its own packing-list HS code.
    expect(sheet.getRow(6).getCell(8).value).toBe('61142000');
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

  it('derives Somme DD HT/TTC from the matching article\'s first physical unit (by HS code prefix), without showing the individual tax columns', async () => {
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

    // Requested: the individual tax montant columns (DTS IMPORT NORMAL, TVA
    // IMPORT AUTRE PDS, ...) that used to sit between HS CODE and Somme DD
    // HT are no longer shown as their own columns — instead their "<tax>
    // Total" columns (11/12 below) sit right after HS CODE (col 8), and
    // Somme DD HT (col 11) sits right after those, immediately before Somme
    // DD TTC.
    const taxTotalArgb = 'FFDB2777';
    expect(headerRow.getCell(9).value).toBe('DTS IMPORT NORMAL Total');
    expect(headerRow.getCell(10).value).toBe('TVA IMPORT AUTRE PDS Total');
    expect((headerRow.getCell(9).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(taxTotalArgb);
    expect((headerRow.getCell(10).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(taxTotalArgb);

    const sumHtArgb = 'FF15803D';
    expect(headerRow.getCell(11).value).toBe('Somme DD HT');
    expect(headerRow.getCell(12).value).toBe('DD unitaire HT');
    expect((headerRow.getCell(11).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumHtArgb);
    expect((headerRow.getCell(12).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumHtArgb);

    // Row 5 — HS code 61044300 matches the one article's 610443 prefix. The
    // article's montant (4.32 / 34.6) is spread across its 18 units, and
    // only the FIRST unit's share is taken — not the article's full total.
    // allocateTaxAcrossUnits works in integer cents, spreading any leftover
    // cent onto the first few units — 4.32/18 divides evenly (0.24), but
    // 34.6/18 doesn't, so the first unit gets 1.93 (192 cents + 1 leftover),
    // not the plain division's 1.9222.
    const matchedRow = sheet.getRow(5);
    const firstUnit000110 = 0.24;
    const firstUnit002109 = 1.93;

    // Somme DD HT = Somme DD TTC - TVA IMPORT AUTRE PDS Total, i.e. every
    // tax's row total (montant x pieces) except VAT.
    const expectedSommeDd = firstUnit000110 * 18;
    expect(Number(matchedRow.getCell(11).value)).toBeCloseTo(expectedSommeDd, 2);
    expect(Number(matchedRow.getCell(12).value)).toBeCloseTo(expectedSommeDd / 18, 6);
    expect(matchedRow.getCell(12).numFmt).toBe('0000.000000');
    // Somme DD HT is zero-padded to at least 4 digits, no thousands separator.
    expect(matchedRow.getCell(11).numFmt).toBe('0000.00');
    // The light-green tint carries down the data rows too, not just the
    // header, so the columns read as visibly distinct all the way down.
    const sumHtRowFillArgb = 'FFDCFCE7';
    expect((matchedRow.getCell(11).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumHtRowFillArgb);
    expect((matchedRow.getCell(12).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumHtRowFillArgb);
    // Pieces (col 4) is a plain whole count — no thousands separator, no decimals.
    expect(matchedRow.getCell(4).numFmt).toBe('0');

    // Row 6 — HS code 61142000 (prefix 611420) has no matching article, so
    // Somme DD HT defaults to 0.
    const unmatchedRow = sheet.getRow(6);
    expect(Number(unmatchedRow.getCell(11).value)).toBe(0);
    expect(Number(unmatchedRow.getCell(12).value)).toBe(0);
    expect((unmatchedRow.getCell(11).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumHtRowFillArgb);

    // 13 Somme DD TTC / 14 DD unitaire TTC close out the taxes section —
    // Somme DD TTC is the sum of the "<tax> Total" columns for the row (VAT
    // included, unlike Somme DD HT), and DD unitaire TTC spreads that sum
    // per piece.
    const sumTtcArgb = 'FF4D7C0F';
    const sumTtcRowFillArgb = 'FFECFCCB';
    expect(headerRow.getCell(13).value).toBe('Somme DD TTC');
    expect(headerRow.getCell(14).value).toBe('DD unitaire TTC');
    expect((headerRow.getCell(13).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumTtcArgb);
    expect((headerRow.getCell(14).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumTtcArgb);

    const expectedTaxTotal000110 = firstUnit000110 * 18;
    const expectedTaxTotal002109 = firstUnit002109 * 18;
    const expectedSommeDdTtc = expectedTaxTotal000110 + expectedTaxTotal002109;
    expect(Number(matchedRow.getCell(9).value)).toBeCloseTo(expectedTaxTotal000110, 2);
    expect(Number(matchedRow.getCell(10).value)).toBeCloseTo(expectedTaxTotal002109, 2);
    expect(Number(matchedRow.getCell(13).value)).toBeCloseTo(expectedSommeDdTtc, 2);
    expect(Number(matchedRow.getCell(14).value)).toBeCloseTo(expectedSommeDdTtc / 18, 6);
    expect(matchedRow.getCell(13).numFmt).toBe('0000.00');
    expect(matchedRow.getCell(14).numFmt).toBe('0000.000000');
    expect((matchedRow.getCell(13).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumTtcRowFillArgb);
    expect((matchedRow.getCell(14).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumTtcRowFillArgb);

    expect(Number(unmatchedRow.getCell(13).value)).toBe(0);
    expect(Number(unmatchedRow.getCell(14).value)).toBe(0);
    expect((unmatchedRow.getCell(13).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(sumTtcRowFillArgb);
  });

  it('uses only the first matching article\'s first-unit tax value — not summed with other articles sharing the same HS code prefix', async () => {
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
          // A deliberately very different montant/quantite (per-unit 5.00,
          // vs article 1's per-unit 0.25) so any accidental blending with
          // article 1's value would be obvious.
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
          taxes: [{ code: '000110', assiette: 50, taux: 2.5, montant: 25 }],
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

    // Both rows share HS position 610443 and origin CHINE, so they merge
    // into a single HS-total group row (36 combined pieces). The group's
    // tax montant is only article 1's (the first encountered, matching
    // Global's row order) first unit: 2.5 / 10 = 0.25, not article 2's
    // 25 / 5 = 5.00, and not a blend of both. Read via Somme DD HT (col 10,
    // right after the one tax's "Total" column at col 9) since the
    // individual tax columns aren't shown; the only tax here is 000110 (not
    // VAT), so Somme DD HT equals Somme DD TTC, i.e. that per-unit value
    // spread across the merged row's 36 pieces (18 + 18).
    expect(Number(sheet.getRow(5).getCell(10).value)).toBeCloseTo(0.25 * 36, 2);
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

    // Row 5 (origin CHINA, normalized to CHINE) picks up only the CHINE
    // article's first-unit value: 2.5 / 10 = 0.25. Row 6 (origin
    // BANGLADESH) picks up only the BANGLADESH article's first-unit value:
    // 999 cents / 5 doesn't divide evenly, so allocateTaxAcrossUnits gives
    // the first unit 200 cents (2.00), not the plain division's 1.998. Read
    // via Somme DD HT (col 10, right after the one tax's "Total" column at
    // col 9) since the individual tax columns aren't shown; the only tax
    // here is 000110 (not VAT), so Somme DD HT equals Somme DD TTC, i.e.
    // that per-unit value spread across each row's 18 pieces.
    expect(Number(sheet.getRow(5).getCell(10).value)).toBeCloseTo((2.5 / 10) * 18, 2);
    expect(Number(sheet.getRow(6).getCell(10).value)).toBeCloseTo(2.0 * 18, 2);
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
    // declaration, so the mismatched spelling never comes into play. Value
    // is the article's first-unit share (4.32 / 18) spread across the row's
    // 18 pieces, i.e. 4.32 again. Read via Somme DD HT (col 10, right after
    // the one tax's "Total" column at col 9).
    expect(Number(sheet.getRow(5).getCell(10).value)).toBeCloseTo((4.32 / 18) * 18, 2);
  });

  it('adds a column for a RECAPITULATION rubrique that never appears on any article, filled as montant × Prorata for the matched article\'s first unit', async () => {
    // 002701 (REDV.INF.) only shows up in ordonnancementTaxes, never on the
    // article's own tax rows — same declaration-wide-only rubrique Global
    // shows via montant × Prorata; it must get its own column here too.
    const declarationWithOrdonnancement: Declaration = {
      ...DECLARATION_WITH_TAXES,
      ordonnancementTaxes: [{ code: '002701', designation: 'REDV.INF.(AVEC D et T)', montant: 100 }],
    };

    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-ordonnancement');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      declarationWithOrdonnancement,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;
    const headerRow = sheet.getRow(4);

    // Header: item..HS CODE (1-8), then the "<tax> Total" columns in
    // allTaxCodes order (000110, 002109, 002701): 9, 10, REDV.INF. Total
    // (11), then Somme DD HT (12), DD unitaire HT (13), right before Somme
    // DD TTC. REDV.INF. no longer gets its own montant column (the
    // individual tax columns aren't shown), but it still feeds Somme DD HT
    // and gets its own "Total" column.
    expect(headerRow.getCell(11).value).toBe('REDV.INF.(AVEC D et T) Total');
    expect(headerRow.getCell(12).value).toBe('Somme DD HT');
    expect(headerRow.getCell(13).value).toBe('DD unitaire HT');

    // The declaration's only article has valeurDeclaree 172.98 across 18
    // units, and it's also the declaration's only article, so its own
    // Valeur Déclarée is the whole declaration total: Prorata for its first
    // unit = (172.98 / 18) / 172.98 = 1 / 18. REDV.INF. share = 100 / 18.
    const matchedRow = sheet.getRow(5);
    const expectedRedvInf = 100 / 18;

    // 002109 (TVA IMPORT AUTRE PDS) is excluded from Somme DD HT — HT means
    // "Hors Taxe", i.e. without VAT. REDV.INF. is not VAT, so it counts.
    const firstUnit000110 = 0.24;
    const expectedSommeDd = (firstUnit000110 + expectedRedvInf) * 18;
    expect(Number(matchedRow.getCell(12).value)).toBeCloseTo(expectedSommeDd, 2);
    expect(Number(matchedRow.getCell(11).value)).toBeCloseTo(expectedRedvInf * 18, 2);

    // Row 6 has no matching article, so Somme DD HT and REDV.INF. Total
    // default to 0 too.
    const unmatchedRow = sheet.getRow(6);
    expect(Number(unmatchedRow.getCell(12).value)).toBe(0);
    expect(Number(unmatchedRow.getCell(11).value)).toBe(0);
  });

  it('adds a "<tax designation> Total" column per tax, right after DD unitaire HT, equal to montant x pieces', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-tax-totals');
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

    // Columns 1-8 base, then the two "Total" columns right after HS CODE:
    // 9, 10 — headers built from each tax's full designation + " Total"
    // (the individual tax montant columns themselves are no longer shown).
    // Somme DD HT / DD unitaire HT (11, 12) follow, right before Somme DD
    // TTC.
    expect(headerRow.getCell(9).value).toBe('DTS IMPORT NORMAL Total');
    expect(headerRow.getCell(10).value).toBe('TVA IMPORT AUTRE PDS Total');

    // Rose, distinct from the neighboring Somme DD HT/TTC greens.
    const taxTotalArgb = 'FFDB2777';
    expect((headerRow.getCell(9).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(taxTotalArgb);
    expect((headerRow.getCell(10).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(taxTotalArgb);

    // Wide enough that the full header text isn't truncated — not the flat
    // 18 used for the other columns.
    expect(sheet.getColumn(9).width).toBeGreaterThanOrEqual('DTS IMPORT NORMAL Total'.length);
    expect(sheet.getColumn(10).width).toBeGreaterThanOrEqual('TVA IMPORT AUTRE PDS Total'.length);

    // Row 5 (matched, 18 pieces): tax value x pieces.
    const matchedRow = sheet.getRow(5);
    const firstUnit000110 = 0.24;
    const firstUnit002109 = 1.93;
    expect(Number(matchedRow.getCell(9).value)).toBeCloseTo(firstUnit000110 * 18, 2);
    expect(Number(matchedRow.getCell(10).value)).toBeCloseTo(firstUnit002109 * 18, 2);
    expect(matchedRow.getCell(9).numFmt).toBe('0000.00');

    // Row 6 (unmatched article, taxes default to 0): total stays 0.
    const unmatchedRow = sheet.getRow(6);
    expect(Number(unmatchedRow.getCell(9).value)).toBe(0);
    expect(Number(unmatchedRow.getCell(10).value)).toBe(0);
  });

  it('builds RI SEGMA and REMISES CREDIT\'s "Total" column names from their full designation, not an abbreviation', async () => {
    const declarationWithBothRc: Declaration = {
      ...SAMPLE_DECLARATION,
      ordonnancementTaxes: [
        { code: '006901', designation: 'RI SEGMA', montant: 10 },
        { code: '006000', designation: 'REMISES CREDIT', montant: 5 },
      ],
    };

    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-rc-override');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      declarationWithBothRc,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;
    const headerRow = sheet.getRow(4);

    // Columns 1-8 base, 9 RI SEGMA Total, 10 REMISES CREDIT Total (RI
    // SEGMA/REMISES CREDIT no longer get their own montant columns), then
    // 11 Somme DD HT, 12 DD unitaire HT, right before Somme DD TTC.
    expect(headerRow.getCell(9).value).toBe('RI SEGMA Total');
    expect(headerRow.getCell(10).value).toBe('REMISES CREDIT Total');
    expect(headerRow.getCell(11).value).toBe('Somme DD HT');
    expect(headerRow.getCell(12).value).toBe('DD unitaire HT');
  });

  it('excludes TVA IMPORT AUTRE PDS (002109) from Somme DD HT but still counts it in Somme DD TTC', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-vat-excluded-from-ht');
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

    // Columns 1-8 base, 9-10 tax totals (000110, 002109), 11 Somme DD HT,
    // 12 DD unitaire HT, 13 Somme DD TTC.
    expect(headerRow.getCell(11).value).toBe('Somme DD HT');
    expect(headerRow.getCell(13).value).toBe('Somme DD TTC');

    const matchedRow = sheet.getRow(5);
    const firstUnit000110 = 0.24;
    const firstUnit002109 = 1.93;
    // Somme DD HT = 000110's total (montant x pieces), VAT left out.
    expect(Number(matchedRow.getCell(11).value)).toBeCloseTo(firstUnit000110 * 18, 2);
    // Somme DD TTC = both taxes' totals (montant x pieces), VAT included.
    const expectedSommeDdTtc = (firstUnit000110 + firstUnit002109) * 18;
    expect(Number(matchedRow.getCell(13).value)).toBeCloseTo(expectedSommeDdTtc, 2);
  });

  it('spreads an extra cost by the sum of Global Prorata across the row\'s HS+origin group, not derived from the row\'s own pieces — PRORATA itself is not a displayed column', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-prorata');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      DECLARATION_WITH_TAXES,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0),
      { fraisTransport: 500 }
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;
    const headerRow = sheet.getRow(4);

    // Columns 1-8 base, 9-10 tax totals, 11-12 Somme/DD unitaire HT, 13-14
    // Somme/DD unitaire TTC, 15 Montant Frais transport (PRORATA itself
    // isn't shown as its own column any more).
    expect(headerRow.getCell(15).value).toBe('Montant Frais transport');

    // Declaration's only article: valeurDeclaree 172.98, and it's also the
    // declaration's only article, so its HS+origin group's summed Prorata
    // share = 172.98 / 172.98 = 1 (rounded to 4 decimal places, matching
    // PRORATA's usual 0.00% display), so the whole 500 lands on this row.
    const matchedRow = sheet.getRow(5);
    expect(Number(matchedRow.getCell(15).value)).toBeCloseTo(500, 6);

    // Row 6 has no matching article, so its Prorata defaults to 0 and none
    // of the extra cost is spread onto it.
    const unmatchedRow = sheet.getRow(6);
    expect(Number(unmatchedRow.getCell(15).value)).toBe(0);
  });

  it('sums Prorata across every article sharing a row\'s HS+origin group, e.g. multiple color variants, and spreads extra costs by that sum', async () => {
    const declarationWithSharedHsAndOrigin: Declaration = {
      ...SAMPLE_DECLARATION,
      articles: [
        {
          numero: 1,
          hsCode: '61044300',
          nomArticle: 'VESTITO A FASCIA CORTO IN TULLE - PINK',
          pays: 'CHINE',
          paysCode: 'CN',
          valeurDeclaree: 100,
          quantite: 10,
          unite: 'U',
          totalArticle: 100,
          poidsNet: 10,
          unitesComplementaires: 1,
          taxes: [],
        },
        {
          numero: 2,
          hsCode: '61044300',
          nomArticle: 'VESTITO A FASCIA CORTO IN TULLE - WHITE',
          pays: 'CHINE',
          paysCode: 'CN',
          valeurDeclaree: 50,
          quantite: 5,
          unite: 'U',
          totalArticle: 50,
          poidsNet: 5,
          unitesComplementaires: 1,
          taxes: [],
        },
        {
          numero: 3,
          hsCode: '61142000',
          nomArticle: 'BODY A COSTE',
          pays: 'BANGLADESH',
          paysCode: 'BD',
          valeurDeclaree: 50,
          quantite: 5,
          unite: 'U',
          totalArticle: 50,
          poidsNet: 5,
          unitesComplementaires: 1,
          taxes: [],
        },
      ],
    };

    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-prorata-grouped');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      declarationWithSharedHsAndOrigin,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0),
      { fraisTransport: 1000 }
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    // No article has any tax, so there are no "<tax> Total" columns here:
    // Montant Frais transport sits at column 13 (8 base + 0 tax totals + 4
    // Somme/DD HT/TTC columns), not the 15 the other test uses with its 2
    // taxes.
    const fraisTransportColumn = 13;

    // Declaration total is 100 + 50 + 50 = 200. The first packing-list row
    // (61044300 / CHINA) matches both the PINK and WHITE articles (same HS
    // prefix + origin), so its Prorata is their combined share:
    // (100 + 50) / 200 = 0.75 — not just the first matched article's own
    // 100 / 200 = 0.5 — so it picks up 75% of the 1000 extra cost.
    const matchedRow = sheet.getRow(5);
    expect(Number(matchedRow.getCell(fraisTransportColumn).value)).toBeCloseTo(750, 6);

    // Second packing-list row (61142000 / BANGLADESH) matches only the BODY
    // article on its own: 50 / 200 = 0.25, i.e. 25% of the extra cost.
    const secondRow = sheet.getRow(6);
    expect(Number(secondRow.getCell(fraisTransportColumn).value)).toBeCloseTo(250, 6);
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

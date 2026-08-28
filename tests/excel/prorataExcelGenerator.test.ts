import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { addProrataSheet } from '../../src/excel/prorataExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
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

describe('addProrataSheet', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('groups articles by HS position + origin, one row each, summing PRORATA across the group', async () => {
    const declaration: Declaration = {
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
          taxes: [],
        },
        {
          numero: 2,
          // Same HS position (610443), different national suffix — same
          // group as article 1.
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
    const { filePath, dir } = createTempXlsxPath('prorata-sheet');
    tempDir = dir;

    await addProrataSheet(
      workbook,
      declaration,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('prorata')!;

    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(1).value).toBe('product');
    expect(headerRow.getCell(2).value).toBe('origin');
    expect(headerRow.getCell(3).value).toBe('HSC');
    expect(headerRow.getCell(4).value).toBe('PRORATA');

    // 3 title rows + header + 2 groups (610443/CHINE merged, 611420/BANGLADESH).
    expect(sheet.rowCount).toBe(6);

    // Declaration total is 100 + 50 + 50 = 200. Group 1 (610443/CHINE) sums
    // both articles' shares: (100 + 50) / 200 = 0.75 — product/HSC taken
    // from the first article encountered (VESTITO PINK / 61044300).
    const row1 = sheet.getRow(5);
    expect(row1.getCell(1).value).toBe('VESTITO PINK');
    expect(row1.getCell(2).value).toBe('CHINE');
    expect(row1.getCell(3).value).toBe('61044300');
    expect(Number(row1.getCell(4).value)).toBeCloseTo(0.75, 6);

    // Group 2 (611420/BANGLADESH) is just article 3 on its own: 50 / 200 = 0.25.
    const row2 = sheet.getRow(6);
    expect(row2.getCell(1).value).toBe('BODY A COSTE');
    expect(row2.getCell(2).value).toBe('BANGLADESH');
    expect(Number(row2.getCell(4).value)).toBeCloseTo(0.25, 6);
  });

  it('writes only the letterhead/header when there are no articles', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('prorata-sheet-empty');
    tempDir = dir;

    await addProrataSheet(
      workbook,
      SAMPLE_DECLARATION,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('prorata')!;

    expect(sheet.rowCount).toBe(4); // 3 title rows + header, no data rows
  });
});

import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { generateVerificationReportExcel } from '../../src/excel/verificationReportExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { Declaration } from '../../src/domain/types.js';
import type { PackingListRow } from '../../src/parser/packingList/packingListParser.js';

const DECLARATION: Declaration = {
  code: '309536',
  redevable: 'MED AFRICA LOGISTICS',
  benNumero: '199',
  articles: [
    {
      numero: 1,
      hsCode: '4202310000',
      nomArticle: 'PORTE FEUILLE ET AUTRE ARTICLE DE POCHE',
      pays: 'CHINE',
      paysCode: 'CN',
      valeurDeclaree: 8039,
      quantite: 15,
      unite: 'NOMBRE',
      totalArticle: 4528,
      poidsNet: 13.01,
      unitesComplementaires: 13.01,
      taxes: [{ code: '000110', assiette: 8039, taux: 30, montant: 2412 }],
    },
    // A second, unrelated article whose only purpose is making INDE a real,
    // known origin in this declaration — so a packing-list row honestly
    // marked INDE for the OTHER (CHINE-only) HS position is recognized as a
    // genuine mismatch, not an unmapped spelling variant to shrug off (see
    // packingListMatcher.ts's originSafeToIgnore).
    {
      numero: 2,
      hsCode: '4202110020',
      nomArticle: 'PORTE-DOCUMENT',
      pays: 'INDE',
      paysCode: 'IN',
      valeurDeclaree: 3507,
      quantite: 3,
      unite: 'NOMBRE',
      totalArticle: 1976,
      poidsNet: 2.6,
      unitesComplementaires: 3,
      taxes: [{ code: '000110', assiette: 3507, taux: 30, montant: 1053 }],
    },
  ],
  ordonnancementTaxes: [],
  numeroEnregistrement: null,
  dateArrivee: null,
  donneesComptables: null,
  titreTransport: null,
  dateDeclaration: null,
};

describe('generateVerificationReportExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('lists unmatched packing-list rows and under-covered declaration groups in two sections', async () => {
    // Row 1 matches (CHINE, 8 pieces of the article's 15). Row 2 is honestly
    // marked INDE — a real country, just not this article's — so it's
    // unmatched, leaving the article's other 7 units uncovered.
    const rows: PackingListRow[] = [
      { item: 'A1', description: 'CLUTCH', color: 'BEIGE', pieces: 8, unit: 73, total: 584, origin: 'CHINE', hsCode: '42023100' },
      { item: 'A2', description: 'WALLET', color: 'BLACK', pieces: 3, unit: 62, total: 186, origin: 'INDE', hsCode: '42023100' },
    ];
    const { filePath, dir } = createTempXlsxPath('verification-report');
    tempDir = dir;

    await generateVerificationReportExcel(DECLARATION, rows, filePath, {
      companyName: null,
      brandColor: null,
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('À vérifier')!;

    // Row 4: unmatched-rows section title. Row 5: its header. Row 6: the
    // one unmatched row (A2/INDE).
    expect(String(sheet.getRow(4).getCell(1).value)).toContain('non reconnues (1)');
    expect(sheet.getRow(5).getCell(1).value).toBe('item');
    expect(sheet.getRow(6).getCell(1).value).toBe('A2');
    expect(sheet.getRow(6).getCell(7).value).toBe('INDE');
    expect(sheet.getRow(6).getCell(8).value).toBe('42023100');

    // Row 7: blank spacer. Row 8: coverage-gap section title. Row 9: header.
    // Row 10: article 1's own gap (15 - 8 = 7 units uncovered). Row 11:
    // article 2 (INDE, 4202110020) — no packing-list row references it at
    // all here, so it's a full gap too, hence "(2)" not "(1)".
    expect(String(sheet.getRow(8).getCell(1).value)).toContain('partiellement ou pas du tout couverts par la packing list (2)');
    expect(sheet.getRow(9).getCell(1).value).toBe('HSC');
    const gapRow = sheet.getRow(10);
    expect(gapRow.getCell(1).value).toBe('4202310000');
    expect(gapRow.getCell(2).value).toBe('CHINE');
    expect(gapRow.getCell(4).value).toBe(15);
    expect(gapRow.getCell(5).value).toBe(8);
    expect(gapRow.getCell(6).value).toBe(7);
    // Missing value = (7/15) * 8039 (the group's total, single article here).
    expect(Number(gapRow.getCell(7).value)).toBeCloseTo((7 / 15) * 8039, 2);
  });

  it('shows an explicit "nothing to report" line in each section when everything matches and every group is fully covered', async () => {
    const rows: PackingListRow[] = Array.from({ length: 15 }, (_, i) => ({
      item: `A${i}`,
      description: 'CLUTCH',
      color: 'BEIGE',
      pieces: 1,
      unit: 73,
      total: 73,
      origin: 'CHINE',
      hsCode: '42023100',
    }));
    const { filePath, dir } = createTempXlsxPath('verification-report-clean');
    tempDir = dir;

    await generateVerificationReportExcel(DECLARATION, rows, filePath, {
      companyName: null,
      brandColor: null,
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('À vérifier')!;

    expect(String(sheet.getRow(4).getCell(1).value)).toContain('non reconnues (0)');
    expect(String(sheet.getRow(6).getCell(1).value)).toContain('Aucune');
  });
});

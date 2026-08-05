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
const NO_BRANDING = { companyName: null, brandColor: null, logoDataUri: null };

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

    await generateUnitLevelExcel(declaration, filePath, NO_BRANDING);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // Rows 1-3 are the company-name/document-reference/date-de-génération
    // title rows (see articleSummaryExcelGenerator.test.ts and
    // excelStyling.test.ts for dedicated coverage of those); row 4 is the
    // actual column header row.
    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(1).value).toBe('Nom Article');
    expect(headerRow.getCell(2).value).toBe('DONNEES COMPTABLES');
    expect(headerRow.getCell(3).value).toBe('Poids net (kg)');
    expect(headerRow.getCell(4).value).toBe("Date d'arrivée");
    expect(headerRow.getCell(5).value).toBe('Nature et numéro du titre de transport');
    expect(headerRow.getCell(6).value).toBe('N° Enregistrement');
    expect(headerRow.getCell(7).value).toBe('Date de déclaration');
    expect(headerRow.getCell(8).value).toBe('Pays');
    expect(headerRow.getCell(9).value).toBe('HSC');
    expect(headerRow.getCell(10).value).toBe('Serial Number');
    // union of tax codes across both articles, sorted: 000110, 002109, 007217
    // — headers show each code's designation, not the raw rubrique number.
    expect(headerRow.getCell(11).value).toBe('DTS IMPORT NORMAL');
    expect(headerRow.getCell(12).value).toBe('TVA IMPORT AUTRE PDS');
    expect(headerRow.getCell(13).value).toBe('TAXE F.P.E.I. EXP.');
    expect(headerRow.getCell(14).value).toBe('Valeur Déclarée');
    expect(headerRow.getCell(15).value).toBe('Prorata');

    // 3 title rows + header + article 1 (354 units) + article 2 (200 units) = 558
    expect(sheet.rowCount).toBe(558);

    // first row of article 1
    const firstRow = sheet.getRow(5);
    expect(firstRow.getCell(1).value).toBe('T-SHIRT');
    // DONNEES COMPTABLES — the DUM fixture's "MLV:" line, same on every row.
    expect(firstRow.getCell(2).value).toBe('MLV:29/06/2026 16:37');
    // Poids net (kg) — article 1's own DUM field 33 value (43.69), distinct
    // from article 2's (16.65), confirming this is per-article not constant.
    expect(Number(firstRow.getCell(3).value)).toBeCloseTo(43.69, 2);
    // Date d'arrivée — the DUM fixture's field 24 value ("24/06/2026"), same on every row.
    expect(firstRow.getCell(4).value).toBe('24/06/2026');
    // Nature et numéro du titre de transport — the DUM fixture's field 17
    // value, same on every row.
    expect(firstRow.getCell(5).value).toBe('01|30100020260009045|147-93618044|MXP|2030300463279');
    // N° Enregistrement — the DUM fixture's "A ENREGISTREMENT" registration
    // number ("0076481 X 25/06/2026"), same on every row.
    expect(firstRow.getCell(6).value).toBe('0076481 X 25/06/2026');
    // Date de déclaration — box 1 "010" + box 4 "301" + registration
    // sequence "0076481" + registration year "2026", same on every row.
    expect(firstRow.getCell(7).value).toBe('010 301 0076481 2026');
    // Pays — article 1's own DUM field 36 value, distinct from article 2's.
    expect(firstRow.getCell(8).value).toBe('ITALIE');
    expect(firstRow.getCell(10).value).toBe(1);
    // Valeur Déclarée (27147.0) / quantite (354) — same value on every row of article 1.
    expect(Number(firstRow.getCell(14).value)).toBeCloseTo(27147.0 / 354, 4);
    // Prorata — this unit's Valeur Déclarée over the whole declaration's total
    // Valeur Déclarée (article 1's 27147.0 + article 2's 12892.992 = 40039.992).
    expect(Number(firstRow.getCell(15).value)).toBeCloseTo(27147.0 / 354 / 40039.992, 6);

    // last row of article 1, first row of article 2 resets serial number
    const lastRowArticle1 = sheet.getRow(358);
    expect(lastRowArticle1.getCell(10).value).toBe(354);
    const firstRowArticle2 = sheet.getRow(359);
    expect(firstRowArticle2.getCell(10).value).toBe(1);
    expect(firstRowArticle2.getCell(1).value).toBe('T-SHIRT');
    // Article 2's own Poids net (kg) value (16.65), not article 1's.
    expect(Number(firstRowArticle2.getCell(3).value)).toBeCloseTo(16.65, 2);
    // Article 2's own Pays value, distinct from article 1's.
    expect(firstRowArticle2.getCell(8).value).toBe('BANGLADESH');
    // Valeur Déclarée (12892.992) / quantite (200) — article 2's own per-unit value.
    expect(Number(firstRowArticle2.getCell(14).value)).toBeCloseTo(12892.992 / 200, 4);
    // Prorata — still divided by the whole declaration's total, not article 2's own total.
    expect(Number(firstRowArticle2.getCell(15).value)).toBeCloseTo(12892.992 / 200 / 40039.992, 6);
    // A thicker top border marks where article 2's block starts.
    expect(firstRowArticle2.getCell(1).border?.top?.style).toBe('medium');

    // Reconciliation: sum each tax column across article 1's 354 rows
    // against the known source montants from the Liquidation fixture:
    // 000110 = 0.00, 002109 = 5443.00, 007217 = 68.00
    let sum000110 = 0;
    let sum002109 = 0;
    let sum007217 = 0;
    for (let rowNum = 5; rowNum <= 358; rowNum++) {
      const row = sheet.getRow(rowNum);
      sum000110 += Number(row.getCell(11).value);
      sum002109 += Number(row.getCell(12).value);
      sum007217 += Number(row.getCell(13).value);
    }
    expect(sum000110).toBeCloseTo(0.0, 2);
    expect(sum002109).toBeCloseTo(5443.0, 2);
    expect(sum007217).toBeCloseTo(68.0, 2);

    // Every unit row's Prorata, summed across both articles, must reconcile
    // to 100% of the declaration — confirms Prorata is divided by the whole
    // declaration's total Valeur Déclarée, not each article's own total.
    let prorataSum = 0;
    for (let rowNum = 5; rowNum <= 558; rowNum++) {
      prorataSum += Number(sheet.getRow(rowNum).getCell(15).value);
    }
    expect(prorataSum).toBeCloseTo(1, 6);
  });

  it('fills every row\'s N° Enregistrement, Date d\'arrivée, DONNEES COMPTABLES, and Nature et numéro du titre de transport with the DUM\'s declaration-wide values when present', async () => {
    const declaration = loadRealDeclaration();
    const withOverrides: Declaration = {
      ...declaration,
      numeroEnregistrement: '0066046 E 08/07/2026',
      dateArrivee: '04/07/2026',
      donneesComptables: 'MLV:14/07/2026 15:17',
      titreTransport: '08|30000020260005678|P3957263/3|ITGOA|2026500066156',
      dateDeclaration: '020 400 0099999 2027',
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-registration');
    tempDir = dir;

    await generateUnitLevelExcel(withOverrides, filePath, NO_BRANDING);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(sheet.getRow(5).getCell(2).value).toBe('MLV:14/07/2026 15:17');
    expect(sheet.getRow(5).getCell(4).value).toBe('04/07/2026');
    expect(sheet.getRow(5).getCell(5).value).toBe('08|30000020260005678|P3957263/3|ITGOA|2026500066156');
    expect(sheet.getRow(5).getCell(6).value).toBe('0066046 E 08/07/2026');
    expect(sheet.getRow(5).getCell(7).value).toBe('020 400 0099999 2027');
    // Every row (including article 2's) carries the same declaration-wide values.
    expect(sheet.getRow(359).getCell(2).value).toBe('MLV:14/07/2026 15:17');
    expect(sheet.getRow(359).getCell(4).value).toBe('04/07/2026');
    expect(sheet.getRow(359).getCell(5).value).toBe('08|30000020260005678|P3957263/3|ITGOA|2026500066156');
    expect(sheet.getRow(359).getCell(6).value).toBe('0066046 E 08/07/2026');
    expect(sheet.getRow(359).getCell(7).value).toBe('020 400 0099999 2027');
  });

  it('throws when an article quantite is not a whole number', async () => {
    const declaration = loadRealDeclaration();
    const brokenDeclaration: Declaration = {
      ...declaration,
      articles: declaration.articles.map((a, i) => (i === 0 ? { ...a, quantite: 354.5 } : a)),
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-broken');
    tempDir = dir;

    await expect(generateUnitLevelExcel(brokenDeclaration, filePath, NO_BRANDING)).rejects.toThrow();
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
          poidsNet: 1.5,
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
          poidsNet: 0.8,
        },
      ],
      ordonnancementTaxes: [],
      numeroEnregistrement: null,
      dateArrivee: null,
      donneesComptables: null,
      titreTransport: null,
      dateDeclaration: null,
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-divergent-codes');
    tempDir = dir;

    await generateUnitLevelExcel(declaration, filePath, NO_BRANDING);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // header: Nom Article | DONNEES COMPTABLES | Poids net (kg) | Date d'arrivée | Nature et numéro du titre de transport | N° Enregistrement | Date de déclaration | Pays | HSC | Serial Number | 000110 | 002109 | 007217 (sorted union) | Valeur Déclarée
    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(11).value).toBe('DTS IMPORT NORMAL');
    expect(headerRow.getCell(12).value).toBe('TVA IMPORT AUTRE PDS');
    expect(headerRow.getCell(13).value).toBe('TAXE F.P.E.I. EXP.');
    expect(headerRow.getCell(14).value).toBe('Valeur Déclarée');

    // article A: 3 rows (rows 5-7), has 000110 and 007217 but NOT 002109 -> 002109 column must be 0
    for (let rowNum = 5; rowNum <= 7; rowNum++) {
      const row = sheet.getRow(rowNum);
      expect(Number(row.getCell(12).value)).toBe(0); // 002109 column, article A doesn't have this code
    }
    // article A's 007217 column (montant=3 across 3 units) should reconcile to 3
    let sumA007217 = 0;
    for (let rowNum = 5; rowNum <= 7; rowNum++) {
      sumA007217 += Number(sheet.getRow(rowNum).getCell(13).value);
    }
    expect(sumA007217).toBeCloseTo(3, 2);

    // article B: 2 rows (rows 8-9), has ONLY 002109 -> 000110 and 007217 columns must be 0
    for (let rowNum = 8; rowNum <= 9; rowNum++) {
      const row = sheet.getRow(rowNum);
      expect(Number(row.getCell(11).value)).toBe(0); // 000110 column, article B doesn't have this code
      expect(Number(row.getCell(13).value)).toBe(0); // 007217 column, article B doesn't have this code
    }
    // article B's 002109 column (montant=10 across 2 units) should reconcile to 10
    let sumB002109 = 0;
    for (let rowNum = 8; rowNum <= 9; rowNum++) {
      sumB002109 += Number(sheet.getRow(rowNum).getCell(12).value);
    }
    expect(sumB002109).toBeCloseTo(10, 2);
  });

  it('adds a column for a RECAPITULATION rubrique that never appears in any article\'s own tax rows, filled as montant × Prorata', async () => {
    // 002701 (REDV.INF.) is a whole-declaration rubrique from the Liquidation's
    // RECAPITULATION table — it never shows up in article A's or B's own tax
    // rows, only in ordonnancementTaxes, so it must still get its own column
    // (built from montant × Prorata) rather than being silently dropped.
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
          taxes: [{ code: '000110', assiette: 100, taux: 0, montant: 0 }],
          totalArticle: 0,
          poidsNet: 1.5,
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
          taxes: [{ code: '000110', assiette: 50, taux: 0, montant: 0 }],
          totalArticle: 0,
          poidsNet: 0.8,
        },
      ],
      ordonnancementTaxes: [{ code: '002701', designation: 'REDV.INF.(AVEC D et T)', montant: 100 }],
      numeroEnregistrement: null,
      dateArrivee: null,
      donneesComptables: null,
      titreTransport: null,
      dateDeclaration: null,
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-ordonnancement');
    tempDir = dir;

    await generateUnitLevelExcel(declaration, filePath, NO_BRANDING);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // header: Nom Article | DONNEES COMPTABLES | Poids net (kg) | Date d'arrivée | Nature et numéro du titre de transport | N° Enregistrement | Date de déclaration | Pays | HSC | Serial Number | 000110 | REDV.INF.(AVEC D et T) | Valeur Déclarée | Prorata
    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(12).value).toBe('REDV.INF.(AVEC D et T)');
    expect(headerRow.getCell(13).value).toBe('Valeur Déclarée');
    expect(headerRow.getCell(14).value).toBe('Prorata');

    // 5 rows total (article A: 3 units, article B: 2 units); declaration total
    // Valeur Déclarée = 150, so each unit's Prorata = (its own per-unit value) / 150.
    let sumOrdonnancement = 0;
    for (let rowNum = 5; rowNum <= 9; rowNum++) {
      sumOrdonnancement += Number(sheet.getRow(rowNum).getCell(12).value);
    }
    // Reconciles back to the rubrique's full montant, since Prorata sums to 1
    // across every row of the declaration.
    expect(sumOrdonnancement).toBeCloseTo(100, 6);

    // Article A's per-unit value is 100/3; its share of the 002701 montant
    // is (100/3 / 150) × 100.
    const firstRow = sheet.getRow(5);
    expect(Number(firstRow.getCell(12).value)).toBeCloseTo(((100 / 3) / 150) * 100, 6);
  });
});

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { generateCombinedExcel } from '../../src/excel/combinedExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { PackingListRow } from '../../src/parser/packingList/packingListParser.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return mergeDeclaration(liquidation, dum);
}

const SAMPLE_PACKING_LIST_ROWS: PackingListRow[] = [
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
];

describe('generateCombinedExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes a single .xlsx file containing the Articles summary, a combined Global sheet, and an HS total sheet', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('combined');
    tempDir = dir;

    await generateCombinedExcel(declaration, SAMPLE_PACKING_LIST_ROWS, filePath, {
      companyName: 'ACME LOGISTICS SARL',
      brandColor: '#4f46e5',
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    // 1 summary sheet + 1 combined "Global" sheet + 1 "HS total" sheet + 1
    // "prorata" sheet.
    expect(workbook.worksheets).toHaveLength(4);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual([
      'Articles',
      'Global',
      'HS total',
      'prorata',
    ]);

    const articlesSheet = workbook.getWorksheet('Articles')!;
    expect(articlesSheet.getRow(1).getCell(1).value).toBe('ACME LOGISTICS SARL');
    expect(articlesSheet.getRow(2).getCell(1).value).toBe(
      `Déclaration ${declaration.code} — ${declaration.redevable}`
    );
    expect(articlesSheet.getRow(4).getCell(1).value).toBe('Nom Article');
    expect(articlesSheet.rowCount).toBe(6); // 3 title rows + header + 2 articles

    const globalSheet = workbook.getWorksheet('Global')!;
    expect(globalSheet.getRow(4).getCell(1).value).toBe('Nom Article');
    expect(globalSheet.rowCount).toBe(558); // 3 title rows + header + 354 + 200 unit rows

    const hsTotalSheet = workbook.getWorksheet('HS total')!;
    expect(hsTotalSheet.getRow(4).getCell(1).value).toBe('item');
    expect(hsTotalSheet.getRow(5).getCell(1).value).toBe('AB0141DOAY16');
    expect(hsTotalSheet.rowCount).toBe(5); // 3 title rows + header + 1 packing-list row

    const prorataSheet = workbook.getWorksheet('prorata')!;
    expect(prorataSheet.getRow(4).getCell(1).value).toBe('product');
    expect(prorataSheet.getRow(4).getCell(4).value).toBe('PRORATA');
    // The 2 real articles share the same HS position (6109100010) but have
    // different origins (ITALIE, BANGLADESH), so they're 2 separate groups —
    // one row each, PRORATA summing to 1 across both (27147 / 40039.992 and
    // 12892.992 / 40039.992).
    expect(prorataSheet.rowCount).toBe(6); // 3 title rows + header + 2 groups
    expect(prorataSheet.getRow(5).getCell(2).value).toBe('ITALIE');
    expect(Number(prorataSheet.getRow(5).getCell(4).value)).toBeCloseTo(27147 / 40039.992, 3);
    expect(prorataSheet.getRow(6).getCell(2).value).toBe('BANGLADESH');
    expect(Number(prorataSheet.getRow(6).getCell(4).value)).toBeCloseTo(12892.992 / 40039.992, 3);

    // All four sheets share the identical "Date de génération" timestamp —
    // proves generateCombinedExcel computed it once, not once per sheet.
    expect(articlesSheet.getRow(3).getCell(1).value).toBe(globalSheet.getRow(3).getCell(1).value);
    expect(articlesSheet.getRow(3).getCell(1).value).toBe(hsTotalSheet.getRow(3).getCell(1).value);
    expect(articlesSheet.getRow(3).getCell(1).value).toBe(prorataSheet.getRow(3).getCell(1).value);
  });
});

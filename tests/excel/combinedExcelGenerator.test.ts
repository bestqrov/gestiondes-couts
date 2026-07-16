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

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return mergeDeclaration(liquidation, dum);
}

describe('generateCombinedExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes a single .xlsx file containing the Articles summary, a combined Global sheet, plus one sheet per product', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('combined');
    tempDir = dir;

    await generateCombinedExcel(declaration, filePath, {
      companyName: 'ACME LOGISTICS SARL',
      brandColor: '#4f46e5',
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    // 1 summary sheet + 1 combined "Global" sheet + 1 sheet per article (2 articles in this fixture)
    expect(workbook.worksheets).toHaveLength(4);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual([
      'Articles',
      'Global',
      '1-T-SHIRT',
      '2-T-SHIRT',
    ]);

    // Every sheet gets the same 2-row letterhead: company name, then the
    // document reference — checked once here (the per-sheet content itself
    // is exercised in each generator's own test file).
    const articlesSheet = workbook.getWorksheet('Articles')!;
    expect(articlesSheet.getRow(1).getCell(1).value).toBe('ACME LOGISTICS SARL');
    expect(articlesSheet.getRow(2).getCell(1).value).toBe(
      `Déclaration ${declaration.code} — ${declaration.redevable}`
    );
    expect(articlesSheet.getRow(3).getCell(1).value).toBe('Nom Article');
    expect(articlesSheet.rowCount).toBe(5); // 2 title rows + header + 2 articles

    const globalSheet = workbook.getWorksheet('Global')!;
    expect(globalSheet.getRow(3).getCell(1).value).toBe('Nom Article');
    expect(globalSheet.rowCount).toBe(557); // 2 title rows + header + 354 + 200 unit rows, both articles combined
    // First article's rows come before the second's, each stacked one under the other.
    expect(globalSheet.getRow(4).getCell(1).value).toBe('T-SHIRT');
    expect(globalSheet.getRow(4).getCell(3).value).toBe(1); // article 1, serial 1
    expect(globalSheet.getRow(357).getCell(3).value).toBe(354); // article 1, serial 354 (last row)
    expect(globalSheet.getRow(358).getCell(3).value).toBe(1); // article 2, serial 1 (first row after article 1)
    expect(globalSheet.getRow(557).getCell(3).value).toBe(200); // article 2, serial 200 (last row)

    // A thicker top border marks where article 2's block starts, visually
    // separating it from article 1's block right above it.
    const separatorBorder = globalSheet.getRow(358).getCell(1).border;
    expect(separatorBorder?.top?.style).toBe('medium');
    // No separator on the very first product's block (row 4) — nothing to separate it from.
    expect(globalSheet.getRow(4).getCell(1).border?.top?.style).not.toBe('medium');

    const article1Sheet = workbook.getWorksheet('1-T-SHIRT')!;
    expect(article1Sheet.getRow(3).getCell(1).value).toBe('Nom Article');
    expect(article1Sheet.rowCount).toBe(357); // 2 title rows + header + 354 unit rows

    const article2Sheet = workbook.getWorksheet('2-T-SHIRT')!;
    expect(article2Sheet.rowCount).toBe(203); // 2 title rows + header + 200 unit rows
  });
});

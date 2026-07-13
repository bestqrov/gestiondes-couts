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

    await generateCombinedExcel(declaration, filePath);

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

    const articlesSheet = workbook.getWorksheet('Articles')!;
    expect(articlesSheet.getRow(1).getCell(1).value).toBe('Nom Article');
    expect(articlesSheet.rowCount).toBe(3); // header + 2 articles

    const globalSheet = workbook.getWorksheet('Global')!;
    expect(globalSheet.getRow(1).getCell(1).value).toBe('Nom Article');
    expect(globalSheet.rowCount).toBe(555); // header + 354 + 200 unit rows, both articles combined
    // First article's rows come before the second's, each stacked one under the other.
    expect(globalSheet.getRow(2).getCell(1).value).toBe('T-SHIRT');
    expect(globalSheet.getRow(2).getCell(3).value).toBe(1); // article 1, serial 1
    expect(globalSheet.getRow(355).getCell(3).value).toBe(354); // article 1, serial 354 (last row)
    expect(globalSheet.getRow(356).getCell(3).value).toBe(1); // article 2, serial 1 (first row after article 1)
    expect(globalSheet.getRow(555).getCell(3).value).toBe(200); // article 2, serial 200 (last row)

    const article1Sheet = workbook.getWorksheet('1-T-SHIRT')!;
    expect(article1Sheet.getRow(1).getCell(1).value).toBe('Nom Article');
    expect(article1Sheet.rowCount).toBe(355); // header + 354 unit rows

    const article2Sheet = workbook.getWorksheet('2-T-SHIRT')!;
    expect(article2Sheet.rowCount).toBe(201); // header + 200 unit rows
  });
});

import ExcelJS from 'exceljs';
import { describe, it, expect } from 'vitest';
import { addSheetTitleRows } from '../../src/excel/excelStyling.js';

describe('addSheetTitleRows', () => {
  it('writes a third row with the formatted "Date de génération" label, using the given Date', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Test');
    sheet.columns = [{ width: 20 }, { width: 20 }, { width: 20 }];

    const generatedAt = new Date(2026, 6, 25, 9, 5); // 25/07/2026 09:05 (month is 0-indexed)
    await addSheetTitleRows(
      workbook,
      sheet,
      3,
      'ACME SARL',
      'Déclaration 123 — ACME',
      'FF4F46E5',
      'FF3730A3',
      null,
      generatedAt
    );

    expect(sheet.getRow(1).getCell(1).value).toBe('ACME SARL');
    expect(sheet.getRow(2).getCell(1).value).toBe('Déclaration 123 — ACME');
    expect(sheet.getRow(3).getCell(1).value).toBe('Date de génération : 25/07/2026 09:05');
    // Merged across all 3 columns, same as rows 1-2.
    expect(sheet.getCell(3, 3).isMerged).toBe(true);
    expect(sheet.getCell(3, 3).master.address).toBe('A3');
  });

  it('zero-pads single-digit day/month/hour/minute', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Test');
    sheet.columns = [{ width: 20 }];

    const generatedAt = new Date(2026, 0, 5, 4, 9); // 05/01/2026 04:09
    await addSheetTitleRows(
      workbook,
      sheet,
      1,
      'ACME SARL',
      'Déclaration 123 — ACME',
      'FF4F46E5',
      'FF3730A3',
      null,
      generatedAt
    );

    expect(sheet.getRow(3).getCell(1).value).toBe('Date de génération : 05/01/2026 04:09');
  });

  it('gives the date row a distinct, smaller, non-bold style from the two rows above it', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Test');
    sheet.columns = [{ width: 20 }];

    await addSheetTitleRows(
      workbook,
      sheet,
      1,
      'ACME SARL',
      'Déclaration 123 — ACME',
      'FF4F46E5',
      'FF3730A3',
      null,
      new Date(2026, 6, 25, 9, 5)
    );

    const dateCell = sheet.getRow(3).getCell(1);
    expect(dateCell.font?.size).toBe(10);
    expect(dateCell.font?.bold).toBeFalsy();
    expect(dateCell.alignment?.horizontal).toBe('center');
  });
});

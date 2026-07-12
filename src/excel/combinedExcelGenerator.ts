import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { addArticleSummarySheet } from './articleSummaryExcelGenerator.js';
import { addUnitLevelSheet } from './unitLevelExcelGenerator.js';

// A single .xlsx with both sheets (Articles, Unit Detail), for the "one file
// to download" workflow — same sheet-building logic as the two standalone
// generators, just added to one shared workbook instead of two.
export async function generateCombinedExcel(declaration: Declaration, outputPath: string): Promise<void> {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: false,
  });

  addArticleSummarySheet(workbook, declaration);
  addUnitLevelSheet(workbook, declaration);

  await workbook.commit();
}

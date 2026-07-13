import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { addArticleSummarySheet } from './articleSummaryExcelGenerator.js';
import { addPerArticleUnitSheets } from './unitLevelExcelGenerator.js';

// A single .xlsx: one "Articles" summary sheet plus one unit-detail sheet
// per product, for the "one file to download" workflow.
export async function generateCombinedExcel(declaration: Declaration, outputPath: string): Promise<void> {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: false,
  });

  addArticleSummarySheet(workbook, declaration);
  addPerArticleUnitSheets(workbook, declaration);

  await workbook.commit();
}

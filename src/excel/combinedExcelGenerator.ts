import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { addArticleSummarySheet } from './articleSummaryExcelGenerator.js';
import { addPerArticleUnitSheets, addUnitLevelSheet } from './unitLevelExcelGenerator.js';

// A single .xlsx: one "Articles" summary sheet, one "Global" sheet with
// every article's unit rows combined (one article's rows after another),
// plus one unit-detail sheet per product — for the "one file to download"
// workflow.
export async function generateCombinedExcel(declaration: Declaration, outputPath: string): Promise<void> {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: false,
  });

  addArticleSummarySheet(workbook, declaration);
  addUnitLevelSheet(workbook, declaration, 'Global');
  addPerArticleUnitSheets(workbook, declaration);

  await workbook.commit();
}

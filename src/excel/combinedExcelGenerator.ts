import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { addArticleSummarySheet } from './articleSummaryExcelGenerator.js';
import { addUnitLevelSheet } from './unitLevelExcelGenerator.js';
import type { BrandingInfo } from './excelStyling.js';

// A single .xlsx: one "Articles" summary sheet, one "Global" sheet with
// every article's unit rows combined (one article's rows after another) —
// for the "one file to download" workflow.
export async function generateCombinedExcel(
  declaration: Declaration,
  outputPath: string,
  branding: BrandingInfo
): Promise<void> {
  const workbook = new ExcelJS.Workbook();

  await addArticleSummarySheet(workbook, declaration, branding);
  await addUnitLevelSheet(workbook, declaration, branding, 'Global');

  await workbook.xlsx.writeFile(outputPath);
}

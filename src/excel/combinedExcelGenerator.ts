import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { addArticleSummarySheet } from './articleSummaryExcelGenerator.js';
import { addUnitLevelSheet } from './unitLevelExcelGenerator.js';
import {
  addPackingListSheet,
  BASE_EXTRA_COST_FIELDS,
  type ExtraCostField,
  type ExtraCosts,
} from './packingListExcelGenerator.js';
import type { BrandingInfo } from './excelStyling.js';
import type { PackingListRow } from '../parser/packingList/packingListParser.js';

// A single .xlsx: one "Articles" summary sheet, one "Global" sheet with
// every article's unit rows combined, and one "HS total" sheet with the
// uploaded packing list's rows as-is — for the "one file to download"
// workflow.
export async function generateCombinedExcel(
  declaration: Declaration,
  packingListRows: PackingListRow[],
  outputPath: string,
  branding: BrandingInfo,
  extraCosts?: ExtraCosts,
  // Custom fields a superadmin has added on top of the 6 built-in ones (see
  // AppSettings.extraCostFields) — defaults to just the built-ins.
  extraCostFields: readonly ExtraCostField[] = BASE_EXTRA_COST_FIELDS
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  // Computed once and passed to every sheet so their letterheads show the
  // identical timestamp, rather than each sheet capturing its own `Date()`
  // a few milliseconds apart.
  const generatedAt = new Date();

  await addArticleSummarySheet(workbook, declaration, branding, generatedAt);
  await addUnitLevelSheet(workbook, declaration, branding, generatedAt, 'Global');
  await addPackingListSheet(
    workbook,
    packingListRows,
    declaration,
    branding,
    generatedAt,
    extraCosts,
    extraCostFields
  );

  await workbook.xlsx.writeFile(outputPath);
}

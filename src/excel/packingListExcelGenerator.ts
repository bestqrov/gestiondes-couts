import ExcelJS from 'exceljs';
import type { PackingListRow } from '../parser/packingList/packingListParser.js';
import type { Declaration } from '../domain/types.js';
import {
  styleDataRow,
  styleHeaderRowGrouped,
  addSheetTitleRows,
  resolveBrandArgb,
  resolveBrandDarkArgb,
  resolveCompanyName,
  resolveDocumentTitle,
  type BrandingInfo,
  type ColumnGroup,
} from './excelStyling.js';

const COLUMN_COUNT = 8;
const UNIT_COLUMN = 5;
const TOTAL_COLUMN = 6;
const COLUMN_GROUPS: ColumnGroup[] = [
  { kind: 'identity', from: 1, to: 3 }, // item, DESCRIPTION, color
  { kind: 'quantity', from: 4, to: 4 }, // pieces
  { kind: 'value', from: 5, to: 6 }, // unit, total
  { kind: 'identity', from: 7, to: 8 }, // origin, HS CODE
];

// Adds the "HS total" sheet — a direct, unaggregated copy of the uploaded
// packing list's rows, in their original order, independent of the
// Liquidation/DUM merge/validation pipeline (`declaration` is only used for
// the letterhead's document-title line, same as every other sheet).
export async function addPackingListSheet(
  workbook: ExcelJS.Workbook,
  rows: PackingListRow[],
  declaration: Declaration,
  branding: BrandingInfo,
  generatedAt: Date,
  sheetName = 'HS total'
): Promise<void> {
  const sheet = workbook.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 4 }] });

  sheet.columns = [
    { key: 'item', width: 24 },
    { key: 'description', width: 40 },
    { key: 'color', width: 24 },
    { key: 'pieces', width: 16 },
    { key: 'unit', width: 18 },
    { key: 'total', width: 18 },
    { key: 'origin', width: 22 },
    { key: 'hsCode', width: 18 },
  ];

  await addSheetTitleRows(
    workbook,
    sheet,
    COLUMN_COUNT,
    resolveCompanyName(branding.companyName),
    resolveDocumentTitle(declaration),
    resolveBrandArgb(branding.brandColor),
    resolveBrandDarkArgb(branding.brandColor),
    branding.logoDataUri,
    generatedAt
  );

  const headerRow = sheet.addRow(['item', 'DESCRIPTION', 'color', 'pieces', 'unit', 'total', 'origin', 'HS CODE']);
  styleHeaderRowGrouped(headerRow, COLUMN_COUNT, COLUMN_GROUPS);

  rows.forEach((row, index) => {
    const excelRow = sheet.addRow({
      item: row.item,
      description: row.description,
      color: row.color,
      pieces: row.pieces,
      unit: row.unit,
      total: row.total,
      origin: row.origin,
      hsCode: row.hsCode,
    });
    styleDataRow(excelRow, COLUMN_COUNT, index, new Set([UNIT_COLUMN, TOTAL_COLUMN]));
  });
}

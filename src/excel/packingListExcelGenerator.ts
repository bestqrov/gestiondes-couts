import ExcelJS from 'exceljs';
import type { PackingListRow } from '../parser/packingList/packingListParser.js';
import type { Declaration } from '../domain/types.js';
import { unionTaxCodes, taxCodeDesignation } from './unitLevelTaxHelpers.js';
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

const BASE_COLUMN_COUNT = 8;
const UNIT_COLUMN = 5;
const TOTAL_COLUMN = 6;
const BASE_COLUMN_GROUPS: ColumnGroup[] = [
  { kind: 'identity', from: 1, to: 3 }, // item, DESCRIPTION, color
  { kind: 'quantity', from: 4, to: 4 }, // pieces
  { kind: 'value', from: 5, to: 6 }, // unit, total
  { kind: 'identity', from: 7, to: 8 }, // origin, HS CODE
];

// Matches a packing-list row to its declaration article by exact name
// (trimmed, case-insensitive) — "Nom Article" (Global) vs "description"
// (packing list) — since both name the same product, just from different
// source documents.
function normalizeArticleName(name: string): string {
  return name.trim().toLowerCase();
}

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
  const taxCodes = unionTaxCodes(declaration.articles);
  const columnCount = BASE_COLUMN_COUNT + taxCodes.length;
  const taxColumns = new Set<number>(taxCodes.map((_, i) => BASE_COLUMN_COUNT + 1 + i));
  const columnGroups: ColumnGroup[] = [
    ...BASE_COLUMN_GROUPS,
    ...(taxCodes.length > 0
      ? [{ kind: 'tax' as const, from: BASE_COLUMN_COUNT + 1, to: columnCount }]
      : []),
  ];

  const articleByName = new Map(
    declaration.articles.map((article) => [normalizeArticleName(article.nomArticle), article])
  );

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
    ...taxCodes.map((code) => ({ key: code, width: 24 })),
  ];

  await addSheetTitleRows(
    workbook,
    sheet,
    columnCount,
    resolveCompanyName(branding.companyName),
    resolveDocumentTitle(declaration),
    resolveBrandArgb(branding.brandColor),
    resolveBrandDarkArgb(branding.brandColor),
    branding.logoDataUri,
    generatedAt
  );

  const headerRow = sheet.addRow([
    'item',
    'DESCRIPTION',
    'color',
    'pieces',
    'unit',
    'total',
    'origin',
    'HS CODE',
    ...taxCodes.map(taxCodeDesignation),
  ]);
  styleHeaderRowGrouped(headerRow, columnCount, columnGroups);

  rows.forEach((row, index) => {
    const matchedArticle = articleByName.get(normalizeArticleName(row.description));
    const rowValues: Record<string, string | number> = {
      item: row.item,
      description: row.description,
      color: row.color,
      pieces: row.pieces,
      unit: row.unit,
      total: row.total,
      origin: row.origin,
      hsCode: row.hsCode,
    };
    for (const code of taxCodes) {
      const tax = matchedArticle?.taxes.find((t) => t.code === code);
      rowValues[code] = tax?.montant ?? 0;
    }
    const excelRow = sheet.addRow(rowValues);
    styleDataRow(excelRow, columnCount, index, new Set([UNIT_COLUMN, TOTAL_COLUMN, ...taxColumns]));
  });
}

import ExcelJS from 'exceljs';
import { isValidHexColor, darken } from '../domain/colorUtils.js';
import type { Declaration } from '../domain/types.js';

// Matches the fallback used on the login page (renderLoginTitle) when no
// company name has been configured in Réglages, so an unbranded deployment
// doesn't ship blank-looking spreadsheets.
export function resolveCompanyName(companyName: string | null | undefined): string {
  return companyName?.trim() ? companyName.trim() : 'Gestion des Coûts';
}

export function resolveDocumentTitle(declaration: Declaration): string {
  return `Déclaration ${declaration.code} — ${declaration.redevable}`;
}

// Deliberately just the two fields the excel generators actually need,
// rather than importing the full AppSettings type from the db layer — keeps
// this module decoupled from persistence concerns.
export interface BrandingInfo {
  companyName: string | null;
  brandColor: string | null;
}

// Shared visual style for every generated sheet — a colored header row,
// thin borders, and alternating row banding for readability. Style objects
// are built once and reused by reference (rather than re-assigning
// cell.fill/cell.font/cell.border separately per cell) so the streaming
// writer's style deduplication (useStyles: true, required for any of this
// to actually render) has to intern only a handful of distinct styles even
// across a sheet with thousands of rows.
const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
};

const MONEY_FORMAT = '#,##0.00';

export const HEADER_STYLE: Partial<ExcelJS.Style> = {
  font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } },
  alignment: { vertical: 'middle', horizontal: 'left' },
  border: THIN_BORDER,
};

const ROW_STYLE_PLAIN: Partial<ExcelJS.Style> = { border: THIN_BORDER };
const ROW_STYLE_PLAIN_BANDED: Partial<ExcelJS.Style> = {
  border: THIN_BORDER,
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } },
};
const ROW_STYLE_MONEY: Partial<ExcelJS.Style> = { border: THIN_BORDER, numFmt: MONEY_FORMAT };
const ROW_STYLE_MONEY_BANDED: Partial<ExcelJS.Style> = {
  border: THIN_BORDER,
  numFmt: MONEY_FORMAT,
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } },
};

/** Applies the shared header look to every cell in a header row. */
export function styleHeaderRow(row: ExcelJS.Row, columnCount: number): void {
  for (let col = 1; col <= columnCount; col++) {
    row.getCell(col).style = HEADER_STYLE;
  }
  row.height = 20;
}

// Fixed, distinguishable header colors per "kind" of column (identity /
// tax / value / quantity), rather than every column sharing one flat
// header color — makes it visually obvious at a glance which fields are
// product identity, which are tax codes, and which are money totals, the
// way a hand-built administrative spreadsheet would be color-coded.
export const COLUMN_GROUP_ARGB = {
  identity: 'FF4F46E5', // indigo — Nom Article / HSC / Pays / Serial Number
  tax: 'FFD97706', // amber — tax code columns
  value: 'FF059669', // emerald — Valeur Déclarée / money totals
  quantity: 'FF0891B2', // teal — Quantité
} as const;

export type ColumnGroupKind = keyof typeof COLUMN_GROUP_ARGB;

export interface ColumnGroup {
  kind: ColumnGroupKind;
  /** 1-indexed, inclusive column range. */
  from: number;
  to: number;
}

/**
 * Same as styleHeaderRow, but colors each cell according to which
 * ColumnGroup its column falls in — the multi-color header look. Columns
 * not covered by any group fall back to the default indigo HEADER_STYLE.
 */
export function styleHeaderRowGrouped(
  row: ExcelJS.Row,
  columnCount: number,
  groups: ColumnGroup[]
): void {
  for (let col = 1; col <= columnCount; col++) {
    const group = groups.find((g) => col >= g.from && col <= g.to);
    row.getCell(col).style = group
      ? {
          font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLUMN_GROUP_ARGB[group.kind] } },
          alignment: { vertical: 'middle', horizontal: 'left' },
          border: THIN_BORDER,
        }
      : HEADER_STYLE;
  }
  row.height = 20;
}

function hexToArgb(hex: string): string {
  return `FF${hex.slice(1).toUpperCase()}`;
}

const DEFAULT_BRAND_HEX = '#4F46E5';

/** The configured brand color as an Excel ARGB string, falling back to the app's default indigo when unset/invalid. */
export function resolveBrandArgb(brandColor: string | null | undefined): string {
  const hex = brandColor && isValidHexColor(brandColor) ? brandColor : DEFAULT_BRAND_HEX;
  return hexToArgb(hex);
}

/** A darker shade of the brand color, for the subtitle row under the main title row. */
export function resolveBrandDarkArgb(brandColor: string | null | undefined): string {
  const hex = brandColor && isValidHexColor(brandColor) ? brandColor : DEFAULT_BRAND_HEX;
  return hexToArgb(darken(hex, 0.25));
}

const TITLE_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'medium', color: { argb: 'FF1E293B' } },
  left: { style: 'medium', color: { argb: 'FF1E293B' } },
  bottom: { style: 'medium', color: { argb: 'FF1E293B' } },
  right: { style: 'medium', color: { argb: 'FF1E293B' } },
};

/**
 * Adds two merged, centered, bold, brand-colored, framed rows at the top of
 * a sheet: the company name (large) and the document reference (e.g.
 * "Déclaration 309536 — MED AFRICA LOGISTICS", smaller, on a darker shade
 * of the same brand color) — the letterhead look for an administrative
 * spreadsheet. Must be called before any other row is added to the sheet,
 * and — because this is a streaming writer — the merge must happen before
 * each row is committed, since a committed row's XML can no longer change.
 */
export function addSheetTitleRows(
  sheet: ExcelJS.Worksheet,
  columnCount: number,
  companyName: string,
  documentTitle: string,
  brandArgb: string,
  brandDarkArgb: string
): void {
  const titleRow = sheet.addRow([companyName]);
  for (let col = 1; col <= columnCount; col++) {
    titleRow.getCell(col).style = {
      font: { bold: true, size: 16, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: brandArgb } },
      alignment: { vertical: 'middle', horizontal: 'center' },
      border: TITLE_BORDER,
    };
  }
  titleRow.height = 30;
  sheet.mergeCells(titleRow.number, 1, titleRow.number, columnCount);
  titleRow.commit();

  const subtitleRow = sheet.addRow([documentTitle]);
  for (let col = 1; col <= columnCount; col++) {
    subtitleRow.getCell(col).style = {
      font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: brandDarkArgb } },
      alignment: { vertical: 'middle', horizontal: 'center' },
      border: TITLE_BORDER,
    };
  }
  subtitleRow.height = 22;
  sheet.mergeCells(subtitleRow.number, 1, subtitleRow.number, columnCount);
  subtitleRow.commit();
}

const GROUP_SEPARATOR_BORDER_COLOR = 'FF1E293B';

/**
 * Marks the start of a new product's block of unit rows in a combined
 * sheet (e.g. "Global", where every article's rows are stacked one after
 * another) with a thicker top border across the row — a clear visual
 * divider between products without inserting an extra spacer row (which
 * would break serial-number/row-count assumptions elsewhere).
 */
export function addGroupSeparatorBorder(row: ExcelJS.Row, columnCount: number): void {
  for (let col = 1; col <= columnCount; col++) {
    const cell = row.getCell(col);
    const existing = (cell.style as Partial<ExcelJS.Style>).border ?? THIN_BORDER;
    cell.style = {
      ...(cell.style as Partial<ExcelJS.Style>),
      border: { ...existing, top: { style: 'medium', color: { argb: GROUP_SEPARATOR_BORDER_COLOR } } },
    };
  }
}

/**
 * Applies thin borders + alternating banding to a data row. `moneyColumns`
 * (1-indexed) get the shared 2-decimal number format on top of that.
 * `rowIndex` is the zero-based index of this row within the sheet's data
 * (not its absolute row number), so banding stays consistent regardless of
 * how many header rows precede it.
 */
export function styleDataRow(row: ExcelJS.Row, columnCount: number, rowIndex: number, moneyColumns: Set<number> = new Set()): void {
  const banded = rowIndex % 2 === 1;
  for (let col = 1; col <= columnCount; col++) {
    const isMoney = moneyColumns.has(col);
    row.getCell(col).style = isMoney
      ? banded
        ? ROW_STYLE_MONEY_BANDED
        : ROW_STYLE_MONEY
      : banded
        ? ROW_STYLE_PLAIN_BANDED
        : ROW_STYLE_PLAIN;
  }
}

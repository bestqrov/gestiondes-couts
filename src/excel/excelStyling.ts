import ExcelJS from 'exceljs';

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

import ExcelJS from 'exceljs';

export interface PackingListRow {
  item: string;
  description: string;
  color: string;
  pieces: number;
  unit: number;
  total: number;
  origin: string;
  hsCode: string;
}

// The source file's header text casing is inconsistent (confirmed against
// real exports: "item" lowercase, "DESCRIPTION"/"HS CODE" uppercase) — every
// header is matched case-insensitively, trimmed, against this lowercase
// reference rather than assuming any particular casing or column order.
const EXPECTED_HEADERS: Record<keyof PackingListRow, string> = {
  item: 'item',
  description: 'description',
  color: 'color',
  pieces: 'pieces',
  unit: 'unit',
  total: 'total',
  origin: 'origin',
  hsCode: 'hs code',
};

type ColumnIndexes = Record<keyof PackingListRow, number>;

function findColumnIndexes(headerRow: ExcelJS.Row, columnCount: number): ColumnIndexes {
  const found: Partial<ColumnIndexes> = {};
  for (let col = 1; col <= columnCount; col++) {
    const cellText = String(headerRow.getCell(col).value ?? '').trim().toLowerCase();
    for (const [field, expected] of Object.entries(EXPECTED_HEADERS) as [keyof PackingListRow, string][]) {
      if (cellText === expected) {
        found[field] = col;
      }
    }
  }

  const missing = (Object.keys(EXPECTED_HEADERS) as (keyof PackingListRow)[]).filter(
    (field) => found[field] === undefined
  );
  if (missing.length > 0) {
    throw new Error(`Packing list: missing expected column(s): ${missing.join(', ')}`);
  }

  return found as ColumnIndexes;
}

// Handles both shapes ExcelJS can hand back for a currency-formatted cell:
// a plain number (numeric cell with a currency display format) or a string
// like "€ 9,61" (text cell) — strips "€" and whitespace, converts the
// decimal comma to a dot.
function parseMoneyCell(value: ExcelJS.CellValue): number {
  if (typeof value === 'number') return value;
  const text = String(value ?? '')
    .replace(/[€\s]/g, '')
    .replace(',', '.');
  const parsed = Number.parseFloat(text);
  if (Number.isNaN(parsed)) {
    throw new Error(`Packing list: cannot parse money value from "${String(value)}"`);
  }
  return parsed;
}

function isRowEmpty(row: ExcelJS.Row, columnIndexes: ColumnIndexes): boolean {
  return Object.values(columnIndexes).every((col) => {
    const value = row.getCell(col).value;
    return value === null || value === undefined || String(value).trim() === '';
  });
}

function cellText(row: ExcelJS.Row, col: number): string {
  return String(row.getCell(col).value ?? '').trim();
}

// Reads a real .xlsx buffer directly (no OCR — this is structured data, not
// a scanned document). Row 1 is the header row; column position for each
// field is found by name, not assumed to be in a fixed order. Rows are
// returned in the exact order they appear in the sheet — no sorting or
// grouping/aggregation by HS code, since this data is placed as-is into its
// own "HS total" sheet, independent of the Liquidation/DUM merge pipeline.
export async function parsePackingList(buffer: Buffer): Promise<PackingListRow[]> {
  const workbook = new ExcelJS.Workbook();
  // Same duplicate-@types/node Buffer nominal-type mismatch worked around
  // elsewhere in this codebase (see excelStyling.ts's addImage call) — both
  // are the same real Buffer instance at runtime.
  await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error('Packing list: no worksheet found in the uploaded file');
  }

  const columnCount = sheet.columnCount;
  const columnIndexes = findColumnIndexes(sheet.getRow(1), columnCount);

  const rows: PackingListRow[] = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    if (isRowEmpty(row, columnIndexes)) continue;

    rows.push({
      item: cellText(row, columnIndexes.item),
      description: cellText(row, columnIndexes.description),
      color: cellText(row, columnIndexes.color),
      pieces: Math.round(Number(row.getCell(columnIndexes.pieces).value)),
      unit: parseMoneyCell(row.getCell(columnIndexes.unit).value),
      total: parseMoneyCell(row.getCell(columnIndexes.total).value),
      origin: cellText(row, columnIndexes.origin),
      hsCode: cellText(row, columnIndexes.hsCode),
    });
  }

  return rows;
}

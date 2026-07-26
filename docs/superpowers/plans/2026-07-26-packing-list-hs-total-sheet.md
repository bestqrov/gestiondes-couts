# Packing List Excel Upload + "HS total" Sheet Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a third, mandatory file upload (a supplier packing-list `.xlsx`) to the "Générer une déclaration" flow, parse it directly with ExcelJS (no OCR), and copy its rows as-is into a new "HS total" sheet appended after "Global" in the generated workbook — independent of the Liquidation/DUM merge/validation pipeline.

**Architecture:** `src/parser/packingList/packingListParser.ts` reads an uploaded `.xlsx` buffer with ExcelJS, matches its 8 expected column headers case-insensitively (order-independent), and returns a plain `PackingListRow[]`. `src/excel/packingListExcelGenerator.ts` mirrors the existing per-sheet generator shape (`addArticleSummarySheet`, `addUnitLevelSheet`) to add the "HS total" sheet with the same letterhead/column-group-coloring conventions. `combinedExcelGenerator.ts`, `scripts/cli.ts`, and `server.ts`'s `/generate` route thread the new required file/rows through. `upload.html` gains a third drop-zone.

**Tech Stack:** TypeScript, ExcelJS, Vitest, Express + multer, vanilla JS in a server-rendered HTML view.

**Design doc:** `docs/superpowers/specs/2026-07-26-packing-list-hs-total-sheet-design.md`

---

### Task 1: `parsePackingList`

**Files:**
- Create: `src/parser/packingList/packingListParser.ts`
- Test: `tests/parser/packingList/packingListParser.test.ts`

- [ ] **Step 1: Write the failing tests**

Create `tests/parser/packingList/packingListParser.test.ts`:

```ts
import ExcelJS from 'exceljs';
import { describe, it, expect } from 'vitest';
import { parsePackingList } from '../../../src/parser/packingList/packingListParser.js';

async function buildFixtureBuffer(
  rows: (string | number)[][],
  headers: string[] = ['item', 'DESCRIPTION', 'color', 'pieces', 'unit', 'total', 'origin', 'HS CODE']
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.addRow(headers);
  for (const row of rows) {
    sheet.addRow(row);
  }
  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

describe('parsePackingList', () => {
  it('parses every field from real-shaped rows (matching the screenshot sample)', async () => {
    const buffer = await buildFixtureBuffer([
      ['AB0141DOAY16', 'VESTITO A FASCIA CORTO IN TULLE', 'PK2 PINK MEDIUM', 18, '€ 9,61', '€ 172,98', 'CHINA', '61044300'],
      ['BD0015DOAY16', 'BODY A COSTE', 'WH2 WHITE', 18, '€ 4,80', '€ 86,40', 'BANGLADESH', '61142000'],
    ]);

    const rows = await parsePackingList(buffer);

    expect(rows).toEqual([
      {
        item: 'AB0141DOAY16',
        description: 'VESTITO A FASCIA CORTO IN TULLE',
        color: 'PK2 PINK MEDIUM',
        pieces: 18,
        unit: 9.61,
        total: 172.98,
        origin: 'CHINA',
        hsCode: '61044300',
      },
      {
        item: 'BD0015DOAY16',
        description: 'BODY A COSTE',
        color: 'WH2 WHITE',
        pieces: 18,
        unit: 4.8,
        total: 86.4,
        origin: 'BANGLADESH',
        hsCode: '61142000',
      },
    ]);
  });

  it('matches headers case-insensitively, tolerating the real mixed-case source (item lowercase, DESCRIPTION uppercase, HS CODE uppercase)', async () => {
    const buffer = await buildFixtureBuffer(
      [['CF0093DOAY16', 'CAMICIA CROP IN SATIN', 'WH1 OFF WHITE', 2, 6.65, 13.3, 'CAMBOGIA', '62064000']],
      ['ITEM', 'Description', 'COLOR', 'Pieces', 'UNIT', 'Total', 'ORIGIN', 'hs code']
    );

    const rows = await parsePackingList(buffer);

    expect(rows).toHaveLength(1);
    expect(rows[0].item).toBe('CF0093DOAY16');
    expect(rows[0].hsCode).toBe('62064000');
  });

  it('parses unit/total from plain numeric cells, not just "€ X,XX"-formatted strings', async () => {
    const buffer = await buildFixtureBuffer([
      ['CF0093DOAY16', 'CAMICIA CROP IN SATIN', 'WH1 OFF WHITE', 2, 6.65, 13.3, 'CAMBOGIA', '62064000'],
    ]);

    const rows = await parsePackingList(buffer);

    expect(rows[0].unit).toBeCloseTo(6.65, 2);
    expect(rows[0].total).toBeCloseTo(13.3, 2);
  });

  it('skips a trailing all-blank row', async () => {
    const buffer = await buildFixtureBuffer([
      ['CF0093DOAY16', 'CAMICIA CROP IN SATIN', 'WH1 OFF WHITE', 2, '€ 6,65', '€ 13,30', 'CAMBOGIA', '62064000'],
      ['', '', '', '', '', '', '', ''],
    ]);

    const rows = await parsePackingList(buffer);

    expect(rows).toHaveLength(1);
  });

  it('preserves the source row order (no sorting/grouping by HS code)', async () => {
    const buffer = await buildFixtureBuffer([
      ['ZZ_LAST', 'ITEM Z', 'COLOR Z', 1, 1, 1, 'CHINA', '90000000'],
      ['AA_FIRST', 'ITEM A', 'COLOR A', 1, 1, 1, 'CHINA', '10000000'],
    ]);

    const rows = await parsePackingList(buffer);

    expect(rows.map((r) => r.item)).toEqual(['ZZ_LAST', 'AA_FIRST']);
  });

  it('throws when a required header is missing', async () => {
    const buffer = await buildFixtureBuffer(
      [['CF0093DOAY16', 'CAMICIA CROP IN SATIN', 'WH1 OFF WHITE', 2, '€ 6,65', '€ 13,30', 'CAMBOGIA']],
      ['item', 'DESCRIPTION', 'color', 'pieces', 'unit', 'total', 'origin']
    );

    await expect(parsePackingList(buffer)).rejects.toThrow(/hsCode/i);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/parser/packingList/packingListParser.test.ts`
Expected: FAIL — module `src/parser/packingList/packingListParser.js` doesn't exist yet.

- [ ] **Step 3: Implement**

Create `src/parser/packingList/packingListParser.ts`:

```ts
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
  await workbook.xlsx.load(buffer);
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/parser/packingList/packingListParser.test.ts`
Expected: PASS (6 tests)

- [ ] **Step 5: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors (this file has no callers yet).

- [ ] **Step 6: Commit**

```bash
git add src/parser/packingList/packingListParser.ts tests/parser/packingList/packingListParser.test.ts
git commit -m "Add parsePackingList: reads a supplier packing-list .xlsx directly"
```

---

### Task 2: `addPackingListSheet`

**Files:**
- Create: `src/excel/packingListExcelGenerator.ts`
- Test: `tests/excel/packingListExcelGenerator.test.ts`

- [ ] **Step 1: Write the failing tests**

Create `tests/excel/packingListExcelGenerator.test.ts`:

```ts
import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { addPackingListSheet } from '../../src/excel/packingListExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { PackingListRow } from '../../src/parser/packingList/packingListParser.js';
import type { Declaration } from '../../src/domain/types.js';

const SAMPLE_DECLARATION: Declaration = {
  code: '309536',
  redevable: 'MED AFRICA LOGISTICS',
  benNumero: '136',
  articles: [],
  ordonnancementTaxes: [],
  numeroEnregistrement: null,
  dateArrivee: null,
  donneesComptables: null,
  titreTransport: null,
};

const SAMPLE_ROWS: PackingListRow[] = [
  {
    item: 'AB0141DOAY16',
    description: 'VESTITO A FASCIA CORTO IN TULLE',
    color: 'PK2 PINK MEDIUM',
    pieces: 18,
    unit: 9.61,
    total: 172.98,
    origin: 'CHINA',
    hsCode: '61044300',
  },
  {
    item: 'BD0015DOAY16',
    description: 'BODY A COSTE',
    color: 'WH2 WHITE',
    pieces: 18,
    unit: 4.8,
    total: 86.4,
    origin: 'BANGLADESH',
    hsCode: '61142000',
  },
];

describe('addPackingListSheet', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes one row per packing-list row, in the given order, with the source column headers', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      SAMPLE_DECLARATION,
      { companyName: 'ACME SARL', brandColor: '#4f46e5', logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    // Rows 1-3: letterhead. Row 4: header.
    expect(sheet.getRow(1).getCell(1).value).toBe('ACME SARL');
    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(1).value).toBe('item');
    expect(headerRow.getCell(2).value).toBe('DESCRIPTION');
    expect(headerRow.getCell(3).value).toBe('color');
    expect(headerRow.getCell(4).value).toBe('pieces');
    expect(headerRow.getCell(5).value).toBe('unit');
    expect(headerRow.getCell(6).value).toBe('total');
    expect(headerRow.getCell(7).value).toBe('origin');
    expect(headerRow.getCell(8).value).toBe('HS CODE');

    // 3 title rows + header + 2 rows = 6
    expect(sheet.rowCount).toBe(6);

    const row1 = sheet.getRow(5);
    expect(row1.getCell(1).value).toBe('AB0141DOAY16');
    expect(row1.getCell(7).value).toBe('CHINA');
    expect(Number(row1.getCell(6).value)).toBeCloseTo(172.98, 2);

    // Original order preserved — BD0015DOAY16 second, not resorted.
    const row2 = sheet.getRow(6);
    expect(row2.getCell(1).value).toBe('BD0015DOAY16');
  });

  it('colors the header by column group: identity, quantity, value', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-colors');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      SAMPLE_ROWS,
      SAMPLE_DECLARATION,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;
    const headerRow = sheet.getRow(4);

    const identityArgb = 'FF4F46E5';
    const quantityArgb = 'FF0891B2';
    const valueArgb = 'FF059669';

    expect((headerRow.getCell(1).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(identityArgb); // item
    expect((headerRow.getCell(4).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(quantityArgb); // pieces
    expect((headerRow.getCell(5).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(valueArgb); // unit
    expect((headerRow.getCell(6).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(valueArgb); // total
    expect((headerRow.getCell(7).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(identityArgb); // origin
    expect((headerRow.getCell(8).fill as ExcelJS.FillPattern).fgColor?.argb).toBe(identityArgb); // HS CODE
  });

  it('writes only the letterhead/header when there are no rows', async () => {
    const workbook = new ExcelJS.Workbook();
    const { filePath, dir } = createTempXlsxPath('packing-list-empty');
    tempDir = dir;

    await addPackingListSheet(
      workbook,
      [],
      SAMPLE_DECLARATION,
      { companyName: null, brandColor: null, logoDataUri: null },
      new Date(2026, 6, 26, 10, 0)
    );
    await workbook.xlsx.writeFile(filePath);

    const readBack = new ExcelJS.Workbook();
    await readBack.xlsx.readFile(filePath);
    const sheet = readBack.getWorksheet('HS total')!;

    expect(sheet.getRow(4).getCell(1).value).toBe('item');
    expect(sheet.rowCount).toBe(4);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/excel/packingListExcelGenerator.test.ts`
Expected: FAIL — module doesn't exist yet.

- [ ] **Step 3: Implement**

Create `src/excel/packingListExcelGenerator.ts`:

```ts
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/excel/packingListExcelGenerator.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 5: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors.

- [ ] **Step 6: Commit**

```bash
git add src/excel/packingListExcelGenerator.ts tests/excel/packingListExcelGenerator.test.ts
git commit -m "Add addPackingListSheet: writes the HS total sheet"
```

---

### Task 3: Wire the packing list into `generateCombinedExcel`

**Files:**
- Modify: `src/excel/combinedExcelGenerator.ts`
- Modify: `tests/excel/combinedExcelGenerator.test.ts`

- [ ] **Step 1: Update `generateCombinedExcel`**

Replace the full contents of `src/excel/combinedExcelGenerator.ts` with:

```ts
import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { addArticleSummarySheet } from './articleSummaryExcelGenerator.js';
import { addUnitLevelSheet } from './unitLevelExcelGenerator.js';
import { addPackingListSheet } from './packingListExcelGenerator.js';
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
  branding: BrandingInfo
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  // Computed once and passed to every sheet so their letterheads show the
  // identical timestamp, rather than each sheet capturing its own `Date()`
  // a few milliseconds apart.
  const generatedAt = new Date();

  await addArticleSummarySheet(workbook, declaration, branding, generatedAt);
  await addUnitLevelSheet(workbook, declaration, branding, generatedAt, 'Global');
  await addPackingListSheet(workbook, packingListRows, declaration, branding, generatedAt);

  await workbook.xlsx.writeFile(outputPath);
}
```

- [ ] **Step 2: Update the test file**

Replace the full contents of `tests/excel/combinedExcelGenerator.test.ts` with:

```ts
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { generateCombinedExcel } from '../../src/excel/combinedExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { PackingListRow } from '../../src/parser/packingList/packingListParser.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return mergeDeclaration(liquidation, dum);
}

const SAMPLE_PACKING_LIST_ROWS: PackingListRow[] = [
  {
    item: 'AB0141DOAY16',
    description: 'VESTITO A FASCIA CORTO IN TULLE',
    color: 'PK2 PINK MEDIUM',
    pieces: 18,
    unit: 9.61,
    total: 172.98,
    origin: 'CHINA',
    hsCode: '61044300',
  },
];

describe('generateCombinedExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes a single .xlsx file containing the Articles summary, a combined Global sheet, and an HS total sheet', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('combined');
    tempDir = dir;

    await generateCombinedExcel(declaration, SAMPLE_PACKING_LIST_ROWS, filePath, {
      companyName: 'ACME LOGISTICS SARL',
      brandColor: '#4f46e5',
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    // 1 summary sheet + 1 combined "Global" sheet + 1 "HS total" sheet.
    expect(workbook.worksheets).toHaveLength(3);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(['Articles', 'Global', 'HS total']);

    const articlesSheet = workbook.getWorksheet('Articles')!;
    expect(articlesSheet.getRow(1).getCell(1).value).toBe('ACME LOGISTICS SARL');
    expect(articlesSheet.getRow(2).getCell(1).value).toBe(
      `Déclaration ${declaration.code} — ${declaration.redevable}`
    );
    expect(articlesSheet.getRow(4).getCell(1).value).toBe('Nom Article');
    expect(articlesSheet.rowCount).toBe(6); // 3 title rows + header + 2 articles

    const globalSheet = workbook.getWorksheet('Global')!;
    expect(globalSheet.getRow(4).getCell(1).value).toBe('Nom Article');
    expect(globalSheet.rowCount).toBe(558); // 3 title rows + header + 354 + 200 unit rows

    const hsTotalSheet = workbook.getWorksheet('HS total')!;
    expect(hsTotalSheet.getRow(4).getCell(1).value).toBe('item');
    expect(hsTotalSheet.getRow(5).getCell(1).value).toBe('AB0141DOAY16');
    expect(hsTotalSheet.rowCount).toBe(5); // 3 title rows + header + 1 packing-list row

    // All three sheets share the identical "Date de génération" timestamp —
    // proves generateCombinedExcel computed it once, not once per sheet.
    expect(articlesSheet.getRow(3).getCell(1).value).toBe(globalSheet.getRow(3).getCell(1).value);
    expect(articlesSheet.getRow(3).getCell(1).value).toBe(hsTotalSheet.getRow(3).getCell(1).value);
  });
});
```

- [ ] **Step 3: Run tests**

Run: `npx vitest run tests/excel/combinedExcelGenerator.test.ts`
Expected: PASS (1 test)

- [ ] **Step 4: Typecheck**

Run: `npx tsc --noEmit`
Expected: errors in `scripts/cli.ts` (still calling `generateCombinedExcel` with the old 3-argument signature) — expected at this point, fixed in Task 4.

- [ ] **Step 5: Commit**

```bash
git add src/excel/combinedExcelGenerator.ts tests/excel/combinedExcelGenerator.test.ts
git commit -m "Add the HS total sheet to generateCombinedExcel"
```

---

### Task 4: Update `scripts/cli.ts`

**Files:**
- Modify: `scripts/cli.ts`

- [ ] **Step 1: Add the packing-list path as a new required CLI argument**

Replace the full contents of `scripts/cli.ts` with:

```ts
import 'dotenv/config';
import { readFile } from 'node:fs/promises';
import { extractDocumentText } from '../src/ocr/documentTextExtractor.js';
import { detectAndParsePair } from '../src/parser/detectAndParsePair.js';
import { mergeDeclaration } from '../src/merge/declarationMerger.js';
import { validateArticle } from '../src/domain/validators.js';
import { generateCombinedExcel } from '../src/excel/combinedExcelGenerator.js';
import { parsePackingList } from '../src/parser/packingList/packingListParser.js';

async function main() {
  const [, , liquidationPath, dumPath, packingListPath, outDir = '.'] = process.argv;

  if (!liquidationPath || !dumPath || !packingListPath) {
    console.error(
      'Usage: npm run generate -- <liquidation-file> <dum-file> <packing-list-file> [output-dir]'
    );
    process.exit(1);
  }

  console.log(`Reading Liquidation: ${liquidationPath}`);
  const liquidationOcr = await extractDocumentText(liquidationPath);
  console.log(`Reading DUM: ${dumPath}`);
  const dumOcr = await extractDocumentText(dumPath);
  console.log(`Reading Packing List: ${packingListPath}`);
  const packingListBuffer = await readFile(packingListPath);
  const packingListRows = await parsePackingList(packingListBuffer);
  console.log(`Packing list: ${packingListRows.length} row(s)`);

  console.log('\n--- Liquidation extracted text (confidence %s) ---', liquidationOcr.confidence);
  console.log(liquidationOcr.text);
  console.log('\n--- DUM extracted text (confidence %s) ---', dumOcr.confidence);
  console.log(dumOcr.text);

  console.log('\n--- Parsing ---');
  const { liquidation, dum, swapped } = detectAndParsePair(liquidationOcr.text, dumOcr.text);
  if (swapped) {
    console.log('(Note: files were auto-detected in reversed order from the arguments given)');
  }

  console.log('--- Merging ---');
  const declaration = mergeDeclaration(liquidation, dum);
  for (const article of declaration.articles) {
    validateArticle(article);
  }
  console.log(`Merged declaration: code=${declaration.code}, ${declaration.articles.length} article(s)`);

  const outputPath = `${outDir}/Declaration.xlsx`;
  // No app settings context from a CLI run — falls back to the default
  // branding (generic company name, indigo accent).
  await generateCombinedExcel(declaration, packingListRows, outputPath, {
    companyName: null,
    brandColor: null,
    logoDataUri: null,
  });

  console.log(`\nGenerated: ${outputPath} (3 sheets: Articles, Global, HS total)`);
}

main().catch((error) => {
  console.error('\nFAILED:', error instanceof Error ? error.message : error);
  process.exit(1);
});
```

- [ ] **Step 2: Typecheck**

Run: `npx tsc --noEmit`
Expected: errors remaining only in `src/web/server.ts` (still calling the old signature) — fixed in Task 5.

- [ ] **Step 3: Commit**

```bash
git add scripts/cli.ts
git commit -m "Add packing-list path as a required CLI argument"
```

---

### Task 5: Wire the third upload into `server.ts`'s `/generate` route

**Files:**
- Modify: `src/web/server.ts`

- [ ] **Step 1: Add the import**

Find the imports block (near the top of `src/web/server.ts`, alongside the other parser imports):

```ts
import { detectAndParsePair } from '../parser/detectAndParsePair.js';
```

Add immediately after it:

```ts
import { parsePackingList } from '../parser/packingList/packingListParser.js';
```

- [ ] **Step 2: Add the third multer field and require it**

Find the `/generate` route (currently):

```ts
app.post(
  '/generate',
  upload.fields([
    { name: 'liquidation', maxCount: 1 },
    { name: 'dum', maxCount: 1 },
  ]),
  async (req, res) => {
    const files = req.files as Record<string, Express.Multer.File[]>;
    const liquidationFile = files.liquidation?.[0];
    const dumFile = files.dum?.[0];

    try {
      if (!liquidationFile || !dumFile) {
        throw new Error('Les deux fichiers (Liquidation et DUM) sont requis.');
      }

      const liquidationOcr = await extractDocumentText(liquidationFile.path);
      const dumOcr = await extractDocumentText(dumFile.path);

      const { liquidation, dum } = detectAndParsePair(liquidationOcr.text, dumOcr.text);
      const declaration = mergeDeclaration(liquidation, dum);
      for (const article of declaration.articles) {
        validateArticle(article);
      }

      // Each request writes to its own uniquely-named file rather than a
      // fixed shared path — two overlapping /generate requests (e.g. a
      // double-submit, or two users at once) were racing on the same fixed
      // "Declaration.xlsx" path, so a request could try to res.download() a
      // file that a concurrent request had just truncated/replaced,
      // producing a spurious "Not Found" (confirmed in production logs).
      const generatedFilePath = path.join(OUTPUT_DIR, `declaration-${randomUUID()}.xlsx`);
      const branding = await getAppSettings(await getSettingsCollection());
      await generateCombinedExcel(declaration, generatedFilePath, branding);
```

Replace it with:

```ts
app.post(
  '/generate',
  upload.fields([
    { name: 'liquidation', maxCount: 1 },
    { name: 'dum', maxCount: 1 },
    { name: 'packingList', maxCount: 1 },
  ]),
  async (req, res) => {
    const files = req.files as Record<string, Express.Multer.File[]>;
    const liquidationFile = files.liquidation?.[0];
    const dumFile = files.dum?.[0];
    const packingListFile = files.packingList?.[0];

    try {
      if (!liquidationFile || !dumFile || !packingListFile) {
        throw new Error('Les trois fichiers (Liquidation, DUM et Excel des articles) sont requis.');
      }

      const liquidationOcr = await extractDocumentText(liquidationFile.path);
      const dumOcr = await extractDocumentText(dumFile.path);
      const packingListBuffer = await readFile(packingListFile.path);
      const packingListRows = await parsePackingList(packingListBuffer);

      const { liquidation, dum } = detectAndParsePair(liquidationOcr.text, dumOcr.text);
      const declaration = mergeDeclaration(liquidation, dum);
      for (const article of declaration.articles) {
        validateArticle(article);
      }

      // Each request writes to its own uniquely-named file rather than a
      // fixed shared path — two overlapping /generate requests (e.g. a
      // double-submit, or two users at once) were racing on the same fixed
      // "Declaration.xlsx" path, so a request could try to res.download() a
      // file that a concurrent request had just truncated/replaced,
      // producing a spurious "Not Found" (confirmed in production logs).
      const generatedFilePath = path.join(OUTPUT_DIR, `declaration-${randomUUID()}.xlsx`);
      const branding = await getAppSettings(await getSettingsCollection());
      await generateCombinedExcel(declaration, packingListRows, generatedFilePath, branding);
```

(`readFile` is already imported at the top of this file from `node:fs/promises` — used elsewhere in `sendXlsxFile`.)

- [ ] **Step 3: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors.

- [ ] **Step 4: Commit**

```bash
git add src/web/server.ts
git commit -m "Require and parse the packing-list upload in /generate"
```

---

### Task 6: Add the third drop-zone to `upload.html`

**Files:**
- Modify: `src/web/views/upload.html`

- [ ] **Step 1: Update the lede paragraph**

Find (near the top of the page body):

```html
    <p class="lede" id="pageLede">Déposez les deux documents (Liquidation et DUM) — l'ordre n'a pas d'importance, ils sont identifiés automatiquement. Un fichier Excel avec une feuille par produit sera généré et téléchargé.</p>
```

Replace with:

```html
    <p class="lede" id="pageLede">Déposez les trois fichiers (Liquidation, DUM, et l'Excel des articles) — l'ordre de Liquidation/DUM n'a pas d'importance, ils sont identifiés automatiquement. Un fichier Excel avec une feuille par produit sera généré et téléchargé.</p>
```

- [ ] **Step 2: Add the third drop-zone HTML**

Find the DUM drop-zone (currently):

```html
        <div class="drop-zone" id="zone-dum">
          <div class="dz-icon">
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M6 2.75h8l4 4V19.5A1.75 1.75 0 0 1 16.25 21h-9A1.75 1.75 0 0 1 5.5 19.25V4.5A1.75 1.75 0 0 1 6 2.75Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
              <path d="M13.5 2.75V7a1 1 0 0 0 1 1h4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
            </svg>
          </div>
          <div class="dz-body">
            <div class="label">DUM</div>
            <div class="hint">PDF ou image — cliquez ou glissez-déposez</div>
            <div class="filename" id="name-dum"></div>
          </div>
          <input type="file" name="dum" id="input-dum" accept=".pdf,.png,.jpg,.jpeg,.tif,.tiff,.bmp" required />
        </div>

        <button type="submit" id="submitBtn">
```

Replace with:

```html
        <div class="drop-zone" id="zone-dum">
          <div class="dz-icon">
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M6 2.75h8l4 4V19.5A1.75 1.75 0 0 1 16.25 21h-9A1.75 1.75 0 0 1 5.5 19.25V4.5A1.75 1.75 0 0 1 6 2.75Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
              <path d="M13.5 2.75V7a1 1 0 0 0 1 1h4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
            </svg>
          </div>
          <div class="dz-body">
            <div class="label">DUM</div>
            <div class="hint">PDF ou image — cliquez ou glissez-déposez</div>
            <div class="filename" id="name-dum"></div>
          </div>
          <input type="file" name="dum" id="input-dum" accept=".pdf,.png,.jpg,.jpeg,.tif,.tiff,.bmp" required />
        </div>

        <div class="drop-zone" id="zone-packingList">
          <div class="dz-icon">
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M6 2.75h8l4 4V19.5A1.75 1.75 0 0 1 16.25 21h-9A1.75 1.75 0 0 1 5.5 19.25V4.5A1.75 1.75 0 0 1 6 2.75Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
              <path d="M13.5 2.75V7a1 1 0 0 0 1 1h4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
            </svg>
          </div>
          <div class="dz-body">
            <div class="label">Excel des articles</div>
            <div class="hint">.xlsx — cliquez ou glissez-déposez</div>
            <div class="filename" id="name-packingList"></div>
          </div>
          <input type="file" name="packingList" id="input-packingList" accept=".xlsx" required />
        </div>

        <button type="submit" id="submitBtn">
```

- [ ] **Step 3: Wire the new drop-zone's JS**

Find:

```js
    wireDropZone('zone-liquidation', 'input-liquidation', 'name-liquidation');
    wireDropZone('zone-dum', 'input-dum', 'name-dum');
```

Replace with:

```js
    wireDropZone('zone-liquidation', 'input-liquidation', 'name-liquidation');
    wireDropZone('zone-dum', 'input-dum', 'name-dum');
    wireDropZone('zone-packingList', 'input-packingList', 'name-packingList');
```

- [ ] **Step 4: Extend the submit handler's missing-file validation**

Find:

```js
      const liquidationFile = document.getElementById('input-liquidation').files[0];
      const dumFile = document.getElementById('input-dum').files[0];
      if (!liquidationFile || !dumFile) {
        const missing = [];
        if (!liquidationFile) missing.push('Liquidation Douanière');
        if (!dumFile) missing.push('DUM');
        showValidationModal('Fichier(s) manquant(s) : ' + missing.join(', ') + '. Sélectionnez-les avant de générer.');
        return;
      }
```

Replace with:

```js
      const liquidationFile = document.getElementById('input-liquidation').files[0];
      const dumFile = document.getElementById('input-dum').files[0];
      const packingListFile = document.getElementById('input-packingList').files[0];
      if (!liquidationFile || !dumFile || !packingListFile) {
        const missing = [];
        if (!liquidationFile) missing.push('Liquidation Douanière');
        if (!dumFile) missing.push('DUM');
        if (!packingListFile) missing.push('Excel des articles');
        showValidationModal('Fichier(s) manquant(s) : ' + missing.join(', ') + '. Sélectionnez-les avant de générer.');
        return;
      }
```

- [ ] **Step 5: Append the third file to the FormData sent to `/generate`**

Find:

```js
      const formData = new FormData();
      formData.append('liquidation', liquidationFile);
      formData.append('dum', dumFile);
```

Replace with:

```js
      const formData = new FormData();
      formData.append('liquidation', liquidationFile);
      formData.append('dum', dumFile);
      formData.append('packingList', packingListFile);
```

- [ ] **Step 6: Verify the file parses as valid HTML/JS**

Run:

```bash
node -e "
const fs = require('fs');
const html = fs.readFileSync('src/web/views/upload.html', 'utf-8');
const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
scripts.forEach((m, i) => {
  try { new Function(m[1]); console.log('block', i, 'OK'); }
  catch (e) { console.log('block', i, 'FAILS:', e.message); }
});
"
```

Expected: both script blocks print `OK`.

- [ ] **Step 7: Commit**

```bash
git add src/web/views/upload.html
git commit -m "Add the third packing-list drop-zone to the upload page"
```

---

### Task 7: Full verification

- [ ] **Step 1: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors.

- [ ] **Step 2: Full test suite**

Run: `npx vitest run`
Expected: all tests pass.

- [ ] **Step 3: Manual smoke test via the CLI**

Build a small real `.xlsx` matching the screenshot's shape (8 columns, header row 1), then:

```bash
npx tsx scripts/cli.ts <path-to-a-real-liquidation-pdf> <path-to-a-real-dum-pdf> <path-to-the-packing-list.xlsx> /tmp
```

Open the resulting `/tmp/Declaration.xlsx` and confirm:
- 3 sheets exist: "Articles", "Global", "HS total".
- "HS total" has the same letterhead as the other two sheets, the 8 source columns in the source's row order, and identity/quantity/value column-group header coloring.

- [ ] **Step 4: Manual smoke test of the upload page**

Run the app (`npm run web` or the project's existing dev-run command), open `/`, and confirm:
- Three drop-zones appear: Liquidation Douanière, DUM, Excel des articles.
- Submitting with the packing-list file missing shows the validation modal naming it.
- Submitting all three files succeeds and the downloaded file has the 3 sheets described above.

- [ ] **Step 5: Push**

```bash
git push
```

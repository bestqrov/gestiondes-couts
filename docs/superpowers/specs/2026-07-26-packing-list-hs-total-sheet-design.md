# Packing List Excel Upload + "HS total" Sheet

## Context

A third, mandatory file upload: a supplier packing-list `.xlsx` with one row
per item/color variant (columns: `item`, `DESCRIPTION`, `color`, `pieces`,
`unit`, `total`, `origin`, `HS CODE` — see the user-provided screenshot).
Its rows are copied as-is into a new sheet, "HS total", appended after
"Global" in the same generated workbook. This file's data is **not**
cross-validated against Liquidation/DUM — it's read, parsed, and placed in
its own sheet, independent of the merge/validation pipeline that produces
`Declaration`.

The file format is stable (confirmed by the user): same 8 columns, same
header names, header row always row 1, though header text casing is
inconsistent in the source (`item` lowercase, `DESCRIPTION` uppercase,
`HS CODE` uppercase) — so header matching must be case-insensitive.

## 1. Parsing: `src/parser/packingList/packingListParser.ts`

Unlike Liquidation/DUM (PDF/image, OCR'd into raw text, then regex-parsed),
this is a real `.xlsx` — read directly with `ExcelJS.Workbook().xlsx.load(buffer)`,
no OCR involved.

```ts
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

export function parsePackingList(buffer: Buffer): PackingListRow[]
```

- Reads the first worksheet only.
- Row 1 is the header row. Column position for each of the 8 fields is
  found by matching header cell text case-insensitively (trimmed) against
  `item`, `description`, `color`, `pieces`, `unit`, `total`, `origin`,
  `hs code` — **not** assumed to be in a fixed column order, even though the
  user says it currently always is; this costs almost nothing and removes a
  fragile assumption.
- Throws (`Error`) if any of the 8 expected headers isn't found in row 1 —
  a genuinely malformed upload should fail loudly here, not silently produce
  an empty/wrong sheet.
- For each subsequent non-empty row: reads `pieces` as an integer (via
  `Math.round(Number(cell.value))`), `unit`/`total` via a new
  `parseMoneyCell(cellValue)` helper that handles both cases ExcelJS can
  hand back for a currency-formatted cell — a plain `number` (when the
  source cell is numeric with a currency display format) or a `string` like
  `"€ 9,61"` (when the source cell is text) — stripping `€` and whitespace
  and converting the decimal comma to a dot before parsing. `item`,
  `description`, `color`, `origin`, `hsCode` are read as trimmed strings.
- A row where every one of the 8 cells is empty is skipped (blank trailing
  rows some spreadsheet exports leave behind). A row with only some cells
  empty is not specially handled — the real exports this parses don't
  produce that shape, so this is intentionally not hardened beyond the
  all-empty-row skip.
- Returns the array in the same order the rows appear in the sheet — no
  sorting, no grouping/aggregation by HS code.

## 2. Excel generation: `src/excel/packingListExcelGenerator.ts`

Mirrors the existing per-sheet generator shape (`addArticleSummarySheet`,
`addUnitLevelSheet`):

```ts
export async function addPackingListSheet(
  workbook: ExcelJS.Workbook,
  rows: PackingListRow[],
  declaration: Declaration,
  branding: BrandingInfo,
  generatedAt: Date,
  sheetName = 'HS total'
): Promise<void>
```

- `declaration` is only used for the letterhead's document-title line
  (`resolveDocumentTitle(declaration)`, same as the other two sheets) — no
  other field of `declaration` is read.
- Columns, in this exact order, with the source's own header text (not
  translated to French, since this data already arrives in English from the
  source file): `item`, `DESCRIPTION`, `color`, `pieces`, `unit`, `total`,
  `origin`, `HS CODE`.
- Column-group header coloring (reusing `styleHeaderRowGrouped` /
  `COLUMN_GROUP_ARGB`, the same convention as Articles/Global): identity
  (`item`, `DESCRIPTION`, `color`, `origin`, `HS CODE`) = indigo, quantity
  (`pieces`) = teal, value (`unit`, `total`) = emerald.
- `unit`/`total` columns get the shared money number format (2 decimals);
  `pieces` is a plain integer.
- Same letterhead as every other sheet (`addSheetTitleRows`, 3 rows: company
  name, document title, "Date de génération"), same wide+centered column
  convention established earlier in this project.
- Rows are written in the exact order `rows` is given in (parser's output
  order) — no sorting, no separators between HS-code groups, no
  aggregation, per explicit instruction.
- No group-separator borders (unlike Global's per-product thicker top
  border) — not requested, and there's no grouping concept here to mark
  boundaries of.

## 3. Wiring: `combinedExcelGenerator.ts`, `server.ts`, `upload.html`

**`combinedExcelGenerator.ts`**: `generateCombinedExcel` gains a
`packingListRows: PackingListRow[]` parameter (after `declaration`, before
`outputPath`), and calls `addPackingListSheet` after `addUnitLevelSheet`, so
sheet tab order is Articles → Global → HS total. `scripts/cli.ts`'s call
site updates too (parses a packing list path the same way it parses
liquidation/DUM paths, as a new required CLI argument).

**`server.ts` `/generate`**: `upload.fields([...])` gains a third field,
`{ name: 'packingList', maxCount: 1 }`. The handler requires all three files
(`liquidationFile`, `dumFile`, `packingListFile`) — missing any of them
throws the same `"...sont requis."`-style error as today, just naming all
three. The packing-list file's buffer is read from disk
(`readFile(packingListFile.path)`) and passed to `parsePackingList`
directly — no OCR/`extractDocumentText` call for this one, since it's
already structured data.

**`upload.html`**: a third `.drop-zone` (`zone-packingList` /
`input-packingList`, `accept=".xlsx"`) is added after the DUM drop-zone,
labeled "Fichier Excel (articles)" with hint text "Excel (.xlsx) —
cliquez ou glissez-déposez". `wireDropZone('zone-packingList',
'input-packingList', 'name-packingList')` is called alongside the existing
two. The submit handler's missing-file check and validation-modal message
extend to cover it (three possible missing-file names instead of two). The
page's lede paragraph text updates to mention three documents instead of
two.

## Out of scope

- No cross-validation between the packing list and Liquidation/DUM (no HS
  code or total reconciliation) — confirmed explicitly not wanted.
- No aggregation/grouping by HS code in the "HS total" sheet — it's a
  literal per-row copy despite the sheet's name.
- No sorting or visual separators between HS-code groups.
- The page's broader visual redesign (colors/typography/layout polish
  beyond adding the third drop-zone) is a separate, later piece of work —
  not covered by this spec.

## Testing

- **Parser**: unit tests against a real `.xlsx` fixture built from the
  screenshot's rows (checked into `tests/parser/fixtures/`), covering:
  successful parse of all 8 fields for multiple rows; case-insensitive
  header matching (mixed-case headers, matching the real inconsistency);
  money parsing from both a numeric cell and a `"€ X,XX"`-formatted string
  cell; a trailing all-blank row being skipped; and a missing required
  header throwing.
- **Sheet generator**: `addPackingListSheet`, tested the same way as
  `addArticleSummarySheet`/`addUnitLevelSheet` — row/column values, header
  text and column-group coloring, letterhead presence — using a small
  in-memory `PackingListRow[]` fixture.
- **Combined generator**: extend `combinedExcelGenerator.test.ts` to assert
  a 3rd sheet named "HS total" exists with the right row/column shape.
- **Server**: no new automated test infra exists for `server.ts`'s HTTP
  layer in this repo (confirmed — there's no supertest-style route test
  today); the three-required-files validation is verified manually via the
  existing browser smoke-test pattern used throughout this project.

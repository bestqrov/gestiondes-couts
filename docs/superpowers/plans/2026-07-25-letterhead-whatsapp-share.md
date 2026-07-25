# Excel Letterhead "Date de génération" Row + WhatsApp Share Button Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a third "Date de génération" row to every generated Excel sheet's letterhead, and add a "Partager sur WhatsApp" button next to the existing download button in the web UI that shares the generated `.xlsx` file via the Web Share API.

**Architecture:** `addSheetTitleRows` (in `src/excel/excelStyling.ts`) gains a `generatedAt: Date` parameter and writes a third merged row below the existing two. `combinedExcelGenerator.ts` computes one `Date` and passes it to both sheet builders so both sheets show the identical timestamp. Every downstream "row 3 is the header" assumption (frozen pane `ySplit`, tests) shifts to row 4. The WhatsApp button is pure client-side `<script>` in `upload.html`: feature-detects `navigator.canShare({ files: [...] })` to decide whether to show itself, and on click re-fetches `/download` (the same endpoint the existing "Télécharger" button uses) and calls `navigator.share({ files: [file] })`.

**Tech Stack:** TypeScript, ExcelJS (`excelStyling.ts`), Vitest, vanilla JS in a server-rendered HTML view (`upload.html`) — no new dependencies.

**Design doc:** `docs/superpowers/specs/2026-07-25-letterhead-whatsapp-share-design.md`

---

### Task 1: Add the "Date de génération" row to `addSheetTitleRows`

**Files:**
- Modify: `src/excel/excelStyling.ts:198-254` (the `addSheetTitleRows` function and its JSDoc above it)
- Test: `tests/excel/excelStyling.test.ts` (new file)

- [ ] **Step 1: Write the failing test**

Create `tests/excel/excelStyling.test.ts`:

```ts
import ExcelJS from 'exceljs';
import { describe, it, expect } from 'vitest';
import { addSheetTitleRows } from '../../src/excel/excelStyling.js';

describe('addSheetTitleRows', () => {
  it('writes a third row with the formatted "Date de génération" label, using the given Date', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Test');
    sheet.columns = [{ width: 20 }, { width: 20 }, { width: 20 }];

    const generatedAt = new Date(2026, 6, 25, 9, 5); // 25/07/2026 09:05 (month is 0-indexed)
    await addSheetTitleRows(
      workbook,
      sheet,
      3,
      'ACME SARL',
      'Déclaration 123 — ACME',
      'FF4F46E5',
      'FF3730A3',
      null,
      generatedAt
    );

    expect(sheet.getRow(1).getCell(1).value).toBe('ACME SARL');
    expect(sheet.getRow(2).getCell(1).value).toBe('Déclaration 123 — ACME');
    expect(sheet.getRow(3).getCell(1).value).toBe('Date de génération : 25/07/2026 09:05');
    // Merged across all 3 columns, same as rows 1-2.
    expect(sheet.getCell(3, 3).isMerged).toBe(true);
    expect(sheet.getCell(3, 3).master.address).toBe('A3');
  });

  it('zero-pads single-digit day/month/hour/minute', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Test');
    sheet.columns = [{ width: 20 }];

    const generatedAt = new Date(2026, 0, 5, 4, 9); // 05/01/2026 04:09
    await addSheetTitleRows(
      workbook,
      sheet,
      1,
      'ACME SARL',
      'Déclaration 123 — ACME',
      'FF4F46E5',
      'FF3730A3',
      null,
      generatedAt
    );

    expect(sheet.getRow(3).getCell(1).value).toBe('Date de génération : 05/01/2026 04:09');
  });

  it('gives the date row a distinct, smaller, non-bold style from the two rows above it', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Test');
    sheet.columns = [{ width: 20 }];

    await addSheetTitleRows(
      workbook,
      sheet,
      1,
      'ACME SARL',
      'Déclaration 123 — ACME',
      'FF4F46E5',
      'FF3730A3',
      null,
      new Date(2026, 6, 25, 9, 5)
    );

    const dateCell = sheet.getRow(3).getCell(1);
    expect(dateCell.font?.size).toBe(10);
    expect(dateCell.font?.bold).toBeFalsy();
    expect(dateCell.alignment?.horizontal).toBe('center');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/excel/excelStyling.test.ts`
Expected: FAIL — `addSheetTitleRows` doesn't accept an 8th argument yet, and row 3 doesn't exist (TypeScript will actually fail to compile at this point since the function signature doesn't have a `generatedAt` parameter; that's expected — proceed to Step 3 before re-running).

- [ ] **Step 3: Implement — add `generatedAt` parameter and the third row**

In `src/excel/excelStyling.ts`, add a date-formatting helper right before `addSheetTitleRows` (after the JSDoc comment block that currently ends at line 197, before `export async function addSheetTitleRows`):

```ts
// "Date de génération : DD/MM/YYYY HH:mm" — manually formatted (no
// Intl/locale dependency, matching the rest of this codebase's date
// handling) with a 24h clock and zero-padded fields.
function formatGeneratedAtLabel(generatedAt: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  const day = pad(generatedAt.getDate());
  const month = pad(generatedAt.getMonth() + 1);
  const year = generatedAt.getFullYear();
  const hours = pad(generatedAt.getHours());
  const minutes = pad(generatedAt.getMinutes());
  return `Date de génération : ${day}/${month}/${year} ${hours}:${minutes}`;
}
```

Change the `addSheetTitleRows` signature (currently ending `logoDataUri: string | null\n): Promise<void> {`) to add the new parameter:

```ts
export async function addSheetTitleRows(
  workbook: ExcelJS.Workbook,
  sheet: ExcelJS.Worksheet,
  columnCount: number,
  companyName: string,
  documentTitle: string,
  brandArgb: string,
  brandDarkArgb: string,
  logoDataUri: string | null,
  generatedAt: Date
): Promise<void> {
```

At the very end of the function body, right after the existing subtitle-row block (after `subtitleRow.commit();` and before the function's closing `}`), add the third row:

```ts
  const dateRow = sheet.addRow([formatGeneratedAtLabel(generatedAt)]);
  for (let col = 1; col <= columnCount; col++) {
    dateRow.getCell(col).style = {
      font: { bold: false, size: 10, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: brandDarkArgb } },
      alignment: { vertical: 'middle', horizontal: 'center' },
      border: TITLE_BORDER,
    };
  }
  dateRow.height = 16;
  sheet.mergeCells(dateRow.number, 1, dateRow.number, columnCount);
  dateRow.commit();
```

Also update the function's JSDoc (the comment block starting `/**` above it, currently describing "Adds two merged, bold, brand-colored, framed rows") to say "three" instead of "two", and mention the new date row, e.g. change:

```
 * Adds two merged, bold, brand-colored, framed rows at the top of a sheet:
 * the company name (large, with the logo anchored to its left when one is
 * configured) and the document reference (e.g. "Déclaration 309536 — MED
 * AFRICA LOGISTICS", smaller, on a darker shade of the same brand color) —
 * the letterhead look for an administrative spreadsheet. Must be called
 * before any other row is added to the sheet.
```

to:

```
 * Adds three merged, framed rows at the top of a sheet: the company name
 * (large, bold, with the logo anchored to its left when one is configured),
 * the document reference (e.g. "Déclaration 309536 — MED AFRICA LOGISTICS",
 * smaller, bold, on a darker shade of the same brand color), and a
 * "Date de génération" row (smaller still, not bold, same darker shade) —
 * the letterhead look for an administrative spreadsheet. Must be called
 * before any other row is added to the sheet.
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/excel/excelStyling.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 5: Typecheck**

Run: `npx tsc --noEmit`
Expected: Errors in every caller of `addSheetTitleRows` (`articleSummaryExcelGenerator.ts`, `unitLevelExcelGenerator.ts`) about a missing argument — expected at this point, fixed in Tasks 2-3.

- [ ] **Step 6: Commit**

```bash
git add src/excel/excelStyling.ts tests/excel/excelStyling.test.ts
git commit -m "Add a Date de génération row to addSheetTitleRows"
```

---

### Task 2: Thread `generatedAt` through the Articles sheet generator

**Files:**
- Modify: `src/excel/articleSummaryExcelGenerator.ts`
- Modify: `tests/excel/articleSummaryExcelGenerator.test.ts`

- [ ] **Step 1: Update `addArticleSummarySheet` and `generateArticleSummaryExcel`**

In `src/excel/articleSummaryExcelGenerator.ts`:

Change the `views: [{ state: 'frozen', ySplit: 3 }]` on line 31 to `ySplit: 4` (one more title row before the header):

```ts
  const sheet = workbook.addWorksheet('Articles', { views: [{ state: 'frozen', ySplit: 4 }] });
```

Add a `generatedAt: Date` parameter to `addArticleSummarySheet`'s signature (currently `(workbook, declaration, branding)`):

```ts
export async function addArticleSummarySheet(
  workbook: ExcelJS.Workbook,
  declaration: Declaration,
  branding: BrandingInfo,
  generatedAt: Date
): Promise<void> {
```

Pass it through to `addSheetTitleRows` (add `generatedAt` as the new final argument to the existing call):

```ts
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
```

Update `generateArticleSummaryExcel` (the standalone generator at the bottom of the file) to create its own `generatedAt` and pass it through:

```ts
export async function generateArticleSummaryExcel(
  declaration: Declaration,
  outputPath: string,
  branding: BrandingInfo
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  await addArticleSummarySheet(workbook, declaration, branding, new Date());
  await workbook.xlsx.writeFile(outputPath);
}
```

- [ ] **Step 2: Update the test file for the row shift**

Replace the full contents of `tests/excel/articleSummaryExcelGenerator.test.ts` with:

```ts
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { generateArticleSummaryExcel } from '../../src/excel/articleSummaryExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return mergeDeclaration(liquidation, dum);
}

describe('generateArticleSummaryExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes one row per article with the required columns, from the real merged declaration', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('article-summary');
    tempDir = dir;

    await generateArticleSummaryExcel(declaration, filePath, {
      companyName: 'ACME LOGISTICS SARL',
      brandColor: '#4f46e5',
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // Row 1: company name (merged, centered, bold). Row 2: document
    // reference. Row 3: "Date de génération". Row 4: the actual column
    // header row.
    const titleRow = sheet.getRow(1);
    expect(titleRow.getCell(1).value).toBe('ACME LOGISTICS SARL');
    expect(sheet.getCell(1, 5).isMerged).toBe(true);
    expect(sheet.getCell(1, 5).master.address).toBe('A1'); // merged across all 5 columns

    const subtitleRow = sheet.getRow(2);
    expect(subtitleRow.getCell(1).value).toBe(`Déclaration ${declaration.code} — ${declaration.redevable}`);

    const dateRow = sheet.getRow(3);
    expect(String(dateRow.getCell(1).value)).toMatch(/^Date de génération : \d{2}\/\d{2}\/\d{4} \d{2}:\d{2}$/);

    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(1).value).toBe('Nom Article');
    expect(headerRow.getCell(2).value).toBe('HSC');
    expect(headerRow.getCell(3).value).toBe('Pays');
    expect(headerRow.getCell(4).value).toBe('Valeur déclarée');
    expect(headerRow.getCell(5).value).toBe('Unité (Quantity)');

    // 3 title rows + 1 header row + 2 articles = 6 rows total
    expect(sheet.rowCount).toBe(6);

    const row1 = sheet.getRow(5);
    expect(row1.getCell(1).value).toBe('T-SHIRT');
    expect(row1.getCell(2).value).toBe('6109100010');
    expect(row1.getCell(3).value).toBe('ITALIE');
    expect(row1.getCell(4).value).toBeCloseTo(27147.0, 1);
    expect(row1.getCell(5).value).toBeCloseTo(354.0, 1);

    const row2 = sheet.getRow(6);
    expect(row2.getCell(3).value).toBe('BANGLADESH');
    expect(row2.getCell(5).value).toBeCloseTo(200.0, 1);
  });

  it('embeds the configured logo image on the left of the title row, right-aligning the company name next to it', async () => {
    // A minimal valid 10x10 PNG, base64-encoded as a data: URI — the same
    // shape a real uploaded logo takes (see LOGO_ALLOWED_MIME_TYPES).
    const logoDataUri =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAEklEQVR4nGP4z8CAB+GTG8HSALfKY52fTcuYAAAAAElFTkSuQmCC';
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('article-summary-logo');
    tempDir = dir;

    await generateArticleSummaryExcel(declaration, filePath, {
      companyName: 'ACME LOGISTICS SARL',
      brandColor: '#4f46e5',
      logoDataUri,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(workbook.model.media).toHaveLength(1);
    expect(workbook.model.media[0].type).toBe('image');

    const titleRow = sheet.getRow(1);
    expect(titleRow.getCell(1).value).toBe('ACME LOGISTICS SARL');
    expect(titleRow.getCell(1).alignment?.horizontal).toBe('right');
  });

  it('falls back to a generic company name when none is configured', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('article-summary-no-branding');
    tempDir = dir;

    await generateArticleSummaryExcel(declaration, filePath, { companyName: null, brandColor: null, logoDataUri: null });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(sheet.getRow(1).getCell(1).value).toBe('Gestion des Coûts');
  });

  it('writes only the title/header rows when the declaration has no articles', async () => {
    const emptyDeclaration = {
      code: '000000',
      redevable: 'EMPTY CO',
      benNumero: '1',
      articles: [],
      ordonnancementTaxes: [],
      numeroEnregistrement: null,
      dateArrivee: null,
      donneesComptables: null,
      titreTransport: null,
    };
    const { filePath, dir } = createTempXlsxPath('article-summary-empty');
    tempDir = dir;

    await generateArticleSummaryExcel(emptyDeclaration, filePath, {
      companyName: null,
      brandColor: null,
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(sheet.getRow(4).getCell(1).value).toBe('Nom Article');
    expect(sheet.rowCount).toBe(4);
  });
});
```

- [ ] **Step 3: Run tests**

Run: `npx vitest run tests/excel/articleSummaryExcelGenerator.test.ts`
Expected: PASS (4 tests). (`generateArticleSummaryExcel` is not yet called from `combinedExcelGenerator.ts` with a shared timestamp — that's Task 4 — but this file's own tests only exercise the standalone `generateArticleSummaryExcel`, which now creates its own `Date()` internally, so they pass already.)

- [ ] **Step 4: Commit**

```bash
git add src/excel/articleSummaryExcelGenerator.ts tests/excel/articleSummaryExcelGenerator.test.ts
git commit -m "Thread generatedAt through the Articles sheet generator"
```

---

### Task 3: Thread `generatedAt` through the unit-level (Global) sheet generator

**Files:**
- Modify: `src/excel/unitLevelExcelGenerator.ts`
- Modify: `tests/excel/unitLevelExcelGenerator.test.ts`
- Modify: `tests/excel/unitLevelExcelGenerator.performance.test.ts`

- [ ] **Step 1: Update `addUnitLevelSheet` and `generateUnitLevelExcel`**

In `src/excel/unitLevelExcelGenerator.ts`:

Change line 71's `views: [{ state: 'frozen', ySplit: 3 }]` to `ySplit: 4`:

```ts
  const sheet = workbook.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 4 }] });
```

Add a `generatedAt: Date` parameter to `addUnitLevelSheet`'s signature (currently `(workbook, declaration, branding, sheetName = 'Unit Detail')`) — insert it before `sheetName` since `sheetName` has a default value and TypeScript requires parameters without defaults to come first:

```ts
export async function addUnitLevelSheet(
  workbook: ExcelJS.Workbook,
  declaration: Declaration,
  branding: BrandingInfo,
  generatedAt: Date,
  sheetName = 'Unit Detail'
): Promise<void> {
```

Pass it through to `addSheetTitleRows` (add `generatedAt` as the new final argument):

```ts
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
```

Update `generateUnitLevelExcel` (the standalone generator at the bottom of the file) to create its own `generatedAt` and pass it as the 4th argument (before the now-implicit default `sheetName`):

```ts
export async function generateUnitLevelExcel(
  declaration: Declaration,
  outputPath: string,
  branding: BrandingInfo
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  await addUnitLevelSheet(workbook, declaration, branding, new Date());
  await workbook.xlsx.writeFile(outputPath);
}
```

- [ ] **Step 2: Update `tests/excel/unitLevelExcelGenerator.test.ts` for the row shift**

Replace the full contents of `tests/excel/unitLevelExcelGenerator.test.ts` with:

```ts
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { generateUnitLevelExcel } from '../../src/excel/unitLevelExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { Declaration } from '../../src/domain/types.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');
const NO_BRANDING = { companyName: null, brandColor: null, logoDataUri: null };

function loadRealDeclaration(): Declaration {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return mergeDeclaration(liquidation, dum);
}

describe('generateUnitLevelExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes one row per physical unit, with tax columns that reconcile exactly to the source montants', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('unit-level');
    tempDir = dir;

    await generateUnitLevelExcel(declaration, filePath, NO_BRANDING);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // Rows 1-3 are the company-name/document-reference/date-de-génération
    // title rows (see articleSummaryExcelGenerator.test.ts and
    // excelStyling.test.ts for dedicated coverage of those); row 4 is the
    // actual column header row.
    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(1).value).toBe('Nom Article');
    expect(headerRow.getCell(2).value).toBe('DONNEES COMPTABLES');
    expect(headerRow.getCell(3).value).toBe('Poids net (kg)');
    expect(headerRow.getCell(4).value).toBe("Date d'arrivée");
    expect(headerRow.getCell(5).value).toBe('Nature et numéro du titre de transport');
    expect(headerRow.getCell(6).value).toBe('N° Enregistrement');
    expect(headerRow.getCell(7).value).toBe('HSC');
    expect(headerRow.getCell(8).value).toBe('Serial Number');
    // union of tax codes across both articles, sorted: 000110, 002109, 007217
    // — headers show each code's designation, not the raw rubrique number.
    expect(headerRow.getCell(9).value).toBe('DTS IMPORT NORMAL');
    expect(headerRow.getCell(10).value).toBe('TVA IMPORT AUTRE PDS');
    expect(headerRow.getCell(11).value).toBe('TAXE F.P.E.I. EXP.');
    expect(headerRow.getCell(12).value).toBe('Valeur Déclarée');
    expect(headerRow.getCell(13).value).toBe('Prorata');

    // 3 title rows + header + article 1 (354 units) + article 2 (200 units) = 558
    expect(sheet.rowCount).toBe(558);

    // first row of article 1
    const firstRow = sheet.getRow(5);
    expect(firstRow.getCell(1).value).toBe('T-SHIRT');
    // DONNEES COMPTABLES — the DUM fixture's "MLV:" line, same on every row.
    expect(firstRow.getCell(2).value).toBe('MLV:29/06/2026 16:37');
    // Poids net (kg) — article 1's own DUM field 33 value (43.69), distinct
    // from article 2's (16.65), confirming this is per-article not constant.
    expect(Number(firstRow.getCell(3).value)).toBeCloseTo(43.69, 2);
    // Date d'arrivée — the DUM fixture's field 24 value ("24/06/2026"), same on every row.
    expect(firstRow.getCell(4).value).toBe('24/06/2026');
    // Nature et numéro du titre de transport — the DUM fixture's field 17
    // value, same on every row.
    expect(firstRow.getCell(5).value).toBe('01|30100020260009045|147-93618044|MXP|2030300463279');
    // N° Enregistrement — the DUM fixture's "A ENREGISTREMENT" registration
    // number ("0076481 X 25/06/2026"), same on every row.
    expect(firstRow.getCell(6).value).toBe('0076481 X 25/06/2026');
    expect(firstRow.getCell(8).value).toBe(1);
    // Valeur Déclarée (27147.0) / quantite (354) — same value on every row of article 1.
    expect(Number(firstRow.getCell(12).value)).toBeCloseTo(27147.0 / 354, 4);
    // Prorata — this unit's Valeur Déclarée over the whole declaration's total
    // Valeur Déclarée (article 1's 27147.0 + article 2's 12892.992 = 40039.992).
    expect(Number(firstRow.getCell(13).value)).toBeCloseTo(27147.0 / 354 / 40039.992, 6);

    // last row of article 1, first row of article 2 resets serial number
    const lastRowArticle1 = sheet.getRow(358);
    expect(lastRowArticle1.getCell(8).value).toBe(354);
    const firstRowArticle2 = sheet.getRow(359);
    expect(firstRowArticle2.getCell(8).value).toBe(1);
    expect(firstRowArticle2.getCell(1).value).toBe('T-SHIRT');
    // Article 2's own Poids net (kg) value (16.65), not article 1's.
    expect(Number(firstRowArticle2.getCell(3).value)).toBeCloseTo(16.65, 2);
    // Valeur Déclarée (12892.992) / quantite (200) — article 2's own per-unit value.
    expect(Number(firstRowArticle2.getCell(12).value)).toBeCloseTo(12892.992 / 200, 4);
    // Prorata — still divided by the whole declaration's total, not article 2's own total.
    expect(Number(firstRowArticle2.getCell(13).value)).toBeCloseTo(12892.992 / 200 / 40039.992, 6);
    // A thicker top border marks where article 2's block starts.
    expect(firstRowArticle2.getCell(1).border?.top?.style).toBe('medium');

    // Reconciliation: sum each tax column across article 1's 354 rows
    // against the known source montants from the Liquidation fixture:
    // 000110 = 0.00, 002109 = 5443.00, 007217 = 68.00
    let sum000110 = 0;
    let sum002109 = 0;
    let sum007217 = 0;
    for (let rowNum = 5; rowNum <= 358; rowNum++) {
      const row = sheet.getRow(rowNum);
      sum000110 += Number(row.getCell(9).value);
      sum002109 += Number(row.getCell(10).value);
      sum007217 += Number(row.getCell(11).value);
    }
    expect(sum000110).toBeCloseTo(0.0, 2);
    expect(sum002109).toBeCloseTo(5443.0, 2);
    expect(sum007217).toBeCloseTo(68.0, 2);

    // Every unit row's Prorata, summed across both articles, must reconcile
    // to 100% of the declaration — confirms Prorata is divided by the whole
    // declaration's total Valeur Déclarée, not each article's own total.
    let prorataSum = 0;
    for (let rowNum = 5; rowNum <= 558; rowNum++) {
      prorataSum += Number(sheet.getRow(rowNum).getCell(13).value);
    }
    expect(prorataSum).toBeCloseTo(1, 6);
  });

  it('fills every row\'s N° Enregistrement, Date d\'arrivée, DONNEES COMPTABLES, and Nature et numéro du titre de transport with the DUM\'s declaration-wide values when present', async () => {
    const declaration = loadRealDeclaration();
    const withOverrides: Declaration = {
      ...declaration,
      numeroEnregistrement: '0066046 E 08/07/2026',
      dateArrivee: '04/07/2026',
      donneesComptables: 'MLV:14/07/2026 15:17',
      titreTransport: '08|30000020260005678|P3957263/3|ITGOA|2026500066156',
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-registration');
    tempDir = dir;

    await generateUnitLevelExcel(withOverrides, filePath, NO_BRANDING);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    expect(sheet.getRow(5).getCell(2).value).toBe('MLV:14/07/2026 15:17');
    expect(sheet.getRow(5).getCell(4).value).toBe('04/07/2026');
    expect(sheet.getRow(5).getCell(5).value).toBe('08|30000020260005678|P3957263/3|ITGOA|2026500066156');
    expect(sheet.getRow(5).getCell(6).value).toBe('0066046 E 08/07/2026');
    // Every row (including article 2's) carries the same declaration-wide values.
    expect(sheet.getRow(359).getCell(2).value).toBe('MLV:14/07/2026 15:17');
    expect(sheet.getRow(359).getCell(4).value).toBe('04/07/2026');
    expect(sheet.getRow(359).getCell(5).value).toBe('08|30000020260005678|P3957263/3|ITGOA|2026500066156');
    expect(sheet.getRow(359).getCell(6).value).toBe('0066046 E 08/07/2026');
  });

  it('throws when an article quantite is not a whole number', async () => {
    const declaration = loadRealDeclaration();
    const brokenDeclaration: Declaration = {
      ...declaration,
      articles: declaration.articles.map((a, i) => (i === 0 ? { ...a, quantite: 354.5 } : a)),
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-broken');
    tempDir = dir;

    await expect(generateUnitLevelExcel(brokenDeclaration, filePath, NO_BRANDING)).rejects.toThrow();
  });

  it('zero-fills tax columns for articles that lack a given code, across the actual generator (not just the union helper)', async () => {
    const declaration: Declaration = {
      code: '111111',
      redevable: 'DIVERGENT CODES CO',
      benNumero: '1',
      articles: [
        {
          numero: 1,
          hsCode: '1111111111',
          nomArticle: 'ARTICLE A',
          pays: 'ITALIE',
          paysCode: 'IT',
          valeurDeclaree: 100,
          quantite: 3,
          unite: 'U',
          taxes: [
            { code: '000110', assiette: 100, taux: 0, montant: 0 },
            { code: '007217', assiette: 100, taux: 0.25, montant: 3 },
          ],
          totalArticle: 3,
          poidsNet: 1.5,
        },
        {
          numero: 2,
          hsCode: '2222222222',
          nomArticle: 'ARTICLE B',
          pays: 'BANGLADESH',
          paysCode: 'BD',
          valeurDeclaree: 50,
          quantite: 2,
          unite: 'U',
          taxes: [{ code: '002109', assiette: 50, taux: 20, montant: 10 }],
          totalArticle: 10,
          poidsNet: 0.8,
        },
      ],
      ordonnancementTaxes: [],
      numeroEnregistrement: null,
      dateArrivee: null,
      donneesComptables: null,
      titreTransport: null,
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-divergent-codes');
    tempDir = dir;

    await generateUnitLevelExcel(declaration, filePath, NO_BRANDING);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // header: Nom Article | DONNEES COMPTABLES | Poids net (kg) | Date d'arrivée | Nature et numéro du titre de transport | N° Enregistrement | HSC | Serial Number | 000110 | 002109 | 007217 (sorted union) | Valeur Déclarée
    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(9).value).toBe('DTS IMPORT NORMAL');
    expect(headerRow.getCell(10).value).toBe('TVA IMPORT AUTRE PDS');
    expect(headerRow.getCell(11).value).toBe('TAXE F.P.E.I. EXP.');
    expect(headerRow.getCell(12).value).toBe('Valeur Déclarée');

    // article A: 3 rows (rows 5-7), has 000110 and 007217 but NOT 002109 -> 002109 column must be 0
    for (let rowNum = 5; rowNum <= 7; rowNum++) {
      const row = sheet.getRow(rowNum);
      expect(Number(row.getCell(10).value)).toBe(0); // 002109 column, article A doesn't have this code
    }
    // article A's 007217 column (montant=3 across 3 units) should reconcile to 3
    let sumA007217 = 0;
    for (let rowNum = 5; rowNum <= 7; rowNum++) {
      sumA007217 += Number(sheet.getRow(rowNum).getCell(11).value);
    }
    expect(sumA007217).toBeCloseTo(3, 2);

    // article B: 2 rows (rows 8-9), has ONLY 002109 -> 000110 and 007217 columns must be 0
    for (let rowNum = 8; rowNum <= 9; rowNum++) {
      const row = sheet.getRow(rowNum);
      expect(Number(row.getCell(9).value)).toBe(0); // 000110 column, article B doesn't have this code
      expect(Number(row.getCell(11).value)).toBe(0); // 007217 column, article B doesn't have this code
    }
    // article B's 002109 column (montant=10 across 2 units) should reconcile to 10
    let sumB002109 = 0;
    for (let rowNum = 8; rowNum <= 9; rowNum++) {
      sumB002109 += Number(sheet.getRow(rowNum).getCell(10).value);
    }
    expect(sumB002109).toBeCloseTo(10, 2);
  });

  it('adds a column for a RECAPITULATION rubrique that never appears in any article\'s own tax rows, filled as montant × Prorata', async () => {
    // 002701 (REDV.INF.) is a whole-declaration rubrique from the Liquidation's
    // RECAPITULATION table — it never shows up in article A's or B's own tax
    // rows, only in ordonnancementTaxes, so it must still get its own column
    // (built from montant × Prorata) rather than being silently dropped.
    const declaration: Declaration = {
      code: '111111',
      redevable: 'DIVERGENT CODES CO',
      benNumero: '1',
      articles: [
        {
          numero: 1,
          hsCode: '1111111111',
          nomArticle: 'ARTICLE A',
          pays: 'ITALIE',
          paysCode: 'IT',
          valeurDeclaree: 100,
          quantite: 3,
          unite: 'U',
          taxes: [{ code: '000110', assiette: 100, taux: 0, montant: 0 }],
          totalArticle: 0,
          poidsNet: 1.5,
        },
        {
          numero: 2,
          hsCode: '2222222222',
          nomArticle: 'ARTICLE B',
          pays: 'BANGLADESH',
          paysCode: 'BD',
          valeurDeclaree: 50,
          quantite: 2,
          unite: 'U',
          taxes: [{ code: '000110', assiette: 50, taux: 0, montant: 0 }],
          totalArticle: 0,
          poidsNet: 0.8,
        },
      ],
      ordonnancementTaxes: [{ code: '002701', designation: 'REDV.INF.(AVEC D et T)', montant: 100 }],
      numeroEnregistrement: null,
      dateArrivee: null,
      donneesComptables: null,
      titreTransport: null,
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-ordonnancement');
    tempDir = dir;

    await generateUnitLevelExcel(declaration, filePath, NO_BRANDING);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    // header: Nom Article | DONNEES COMPTABLES | Poids net (kg) | Date d'arrivée | Nature et numéro du titre de transport | N° Enregistrement | HSC | Serial Number | 000110 | REDV.INF.(AVEC D et T) | Valeur Déclarée | Prorata
    const headerRow = sheet.getRow(4);
    expect(headerRow.getCell(10).value).toBe('REDV.INF.(AVEC D et T)');
    expect(headerRow.getCell(11).value).toBe('Valeur Déclarée');
    expect(headerRow.getCell(12).value).toBe('Prorata');

    // 5 rows total (article A: 3 units, article B: 2 units); declaration total
    // Valeur Déclarée = 150, so each unit's Prorata = (its own per-unit value) / 150.
    let sumOrdonnancement = 0;
    for (let rowNum = 5; rowNum <= 9; rowNum++) {
      sumOrdonnancement += Number(sheet.getRow(rowNum).getCell(10).value);
    }
    // Reconciles back to the rubrique's full montant, since Prorata sums to 1
    // across every row of the declaration.
    expect(sumOrdonnancement).toBeCloseTo(100, 6);

    // Article A's per-unit value is 100/3; its share of the 002701 montant
    // is (100/3 / 150) × 100.
    const firstRow = sheet.getRow(5);
    expect(Number(firstRow.getCell(10).value)).toBeCloseTo(((100 / 3) / 150) * 100, 6);
  });
});
```

- [ ] **Step 3: Update `tests/excel/unitLevelExcelGenerator.performance.test.ts` for the row-count shift**

Replace the full contents of `tests/excel/unitLevelExcelGenerator.performance.test.ts` with:

```ts
import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { generateUnitLevelExcel } from '../../src/excel/unitLevelExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { Declaration } from '../../src/domain/types.js';

function makeLargeDeclaration(quantite: number): Declaration {
  return {
    code: '999999',
    redevable: 'PERFORMANCE TEST CO',
    benNumero: '1',
    articles: [
      {
        numero: 1,
        hsCode: '6109100010',
        nomArticle: 'T-SHIRT',
        pays: 'ITALIE',
        paysCode: 'IT',
        valeurDeclaree: 100000,
        quantite,
        unite: 'U',
        taxes: [
          { code: '000110', assiette: 100000, taux: 0, montant: 0 },
          { code: '007217', assiette: 100000, taux: 0.25, montant: 250 },
          { code: '002109', assiette: 100000, taux: 20, montant: 20000 },
        ],
        totalArticle: 20250,
        poidsNet: 43.69,
      },
    ],
    ordonnancementTaxes: [],
    numeroEnregistrement: null,
    dateArrivee: null,
    donneesComptables: null,
    titreTransport: null,
  };
}

describe('generateUnitLevelExcel performance', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  // Generation itself uses ExcelJS's regular in-memory Workbook (not the
  // streaming WorkbookWriter) — switched off streaming so a real logo image
  // could be embedded in the letterhead, which the streaming writer doesn't
  // support (only a whole-sheet background image). Benchmarked separately
  // at this same 10,000-row size beforehand: under 1s, ~55MB heap, so this
  // test exists to guard that budget going forward, not to prove streaming
  // specifically.
  it('generates 10,000+ rows within a reasonable time and with the correct row count', async () => {
    const quantite = 10000;
    const declaration = makeLargeDeclaration(quantite);
    const { filePath, dir } = createTempXlsxPath('unit-level-performance');
    tempDir = dir;

    const start = Date.now();
    await generateUnitLevelExcel(declaration, filePath, { companyName: null, brandColor: null, logoDataUri: null });
    const durationMs = Date.now() - start;

    expect(durationMs).toBeLessThan(15000);

    // The streaming reader here is just a low-memory way to count rows
    // back out for verification — unrelated to how the file was written.
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {});
    let rowCount = 0;
    for await (const worksheetReader of reader) {
      for await (const _row of worksheetReader) {
        rowCount++;
      }
    }

    // quantite data rows + 3 title rows + 1 header row
    expect(rowCount).toBe(quantite + 4);
  }, 30000);
});
```

- [ ] **Step 4: Run tests**

Run: `npx vitest run tests/excel/unitLevelExcelGenerator.test.ts tests/excel/unitLevelExcelGenerator.performance.test.ts`
Expected: PASS (5 + 1 tests)

- [ ] **Step 5: Commit**

```bash
git add src/excel/unitLevelExcelGenerator.ts tests/excel/unitLevelExcelGenerator.test.ts tests/excel/unitLevelExcelGenerator.performance.test.ts
git commit -m "Thread generatedAt through the Global sheet generator"
```

---

### Task 4: Share one `generatedAt` between both sheets in the combined workbook

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
  // Computed once and passed to both sheets so their letterheads show the
  // identical timestamp, rather than each sheet capturing its own `Date()`
  // a few milliseconds apart.
  const generatedAt = new Date();

  await addArticleSummarySheet(workbook, declaration, branding, generatedAt);
  await addUnitLevelSheet(workbook, declaration, branding, generatedAt, 'Global');

  await workbook.xlsx.writeFile(outputPath);
}
```

- [ ] **Step 2: Update the test file for the row shift**

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

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return mergeDeclaration(liquidation, dum);
}

describe('generateCombinedExcel', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('writes a single .xlsx file containing only the Articles summary and a combined Global sheet', async () => {
    const declaration = loadRealDeclaration();
    const { filePath, dir } = createTempXlsxPath('combined');
    tempDir = dir;

    await generateCombinedExcel(declaration, filePath, {
      companyName: 'ACME LOGISTICS SARL',
      brandColor: '#4f46e5',
      logoDataUri: null,
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    // 1 summary sheet + 1 combined "Global" sheet — no per-article sheets.
    expect(workbook.worksheets).toHaveLength(2);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(['Articles', 'Global']);

    // Every sheet gets the same 3-row letterhead: company name, document
    // reference, date de génération — checked once here (the per-sheet
    // content itself is exercised in each generator's own test file).
    const articlesSheet = workbook.getWorksheet('Articles')!;
    expect(articlesSheet.getRow(1).getCell(1).value).toBe('ACME LOGISTICS SARL');
    expect(articlesSheet.getRow(2).getCell(1).value).toBe(
      `Déclaration ${declaration.code} — ${declaration.redevable}`
    );
    expect(articlesSheet.getRow(4).getCell(1).value).toBe('Nom Article');
    expect(articlesSheet.rowCount).toBe(6); // 3 title rows + header + 2 articles

    const globalSheet = workbook.getWorksheet('Global')!;
    expect(globalSheet.getRow(4).getCell(1).value).toBe('Nom Article');
    expect(globalSheet.getRow(4).getCell(2).value).toBe('DONNEES COMPTABLES');
    expect(globalSheet.getRow(4).getCell(3).value).toBe('Poids net (kg)');
    expect(globalSheet.getRow(4).getCell(4).value).toBe("Date d'arrivée");
    expect(globalSheet.getRow(4).getCell(5).value).toBe('Nature et numéro du titre de transport');
    expect(globalSheet.getRow(4).getCell(6).value).toBe('N° Enregistrement');
    expect(globalSheet.getRow(4).getCell(7).value).toBe('HSC');

    // Both sheets' "Date de génération" rows (row 3) show the identical
    // timestamp — proves generateCombinedExcel computed it once, not once
    // per sheet.
    expect(articlesSheet.getRow(3).getCell(1).value).toBe(globalSheet.getRow(3).getCell(1).value);

    // Same DUM-sourced values on every row of the sheet.
    expect(globalSheet.getRow(5).getCell(2).value).toBe('MLV:29/06/2026 16:37');
    expect(globalSheet.getRow(5).getCell(4).value).toBe('24/06/2026');
    expect(globalSheet.getRow(5).getCell(5).value).toBe(
      '01|30100020260009045|147-93618044|MXP|2030300463279'
    );
    expect(globalSheet.getRow(5).getCell(6).value).toBe('0076481 X 25/06/2026');
    expect(globalSheet.getRow(558).getCell(2).value).toBe('MLV:29/06/2026 16:37');
    expect(globalSheet.getRow(558).getCell(4).value).toBe('24/06/2026');
    expect(globalSheet.getRow(558).getCell(6).value).toBe('0076481 X 25/06/2026');
    expect(globalSheet.rowCount).toBe(558); // 3 title rows + header + 354 + 200 unit rows, both articles combined
    // First article's rows come before the second's, each stacked one under the other.
    expect(globalSheet.getRow(5).getCell(1).value).toBe('T-SHIRT');
    expect(globalSheet.getRow(5).getCell(8).value).toBe(1); // article 1, serial 1
    expect(globalSheet.getRow(358).getCell(8).value).toBe(354); // article 1, serial 354 (last row)
    expect(globalSheet.getRow(359).getCell(8).value).toBe(1); // article 2, serial 1 (first row after article 1)
    expect(globalSheet.getRow(558).getCell(8).value).toBe(200); // article 2, serial 200 (last row)

    // A thicker top border marks where article 2's block starts, visually
    // separating it from article 1's block right above it.
    const separatorBorder = globalSheet.getRow(359).getCell(1).border;
    expect(separatorBorder?.top?.style).toBe('medium');
    // No separator on the very first product's block (row 5) — nothing to separate it from.
    expect(globalSheet.getRow(5).getCell(1).border?.top?.style).not.toBe('medium');
  });
});
```

- [ ] **Step 3: Run tests**

Run: `npx vitest run tests/excel/combinedExcelGenerator.test.ts`
Expected: PASS (1 test)

- [ ] **Step 4: Commit**

```bash
git add src/excel/combinedExcelGenerator.ts tests/excel/combinedExcelGenerator.test.ts
git commit -m "Share one generatedAt timestamp between the Articles and Global sheets"
```

---

### Task 5: Full verification of the Excel letterhead work

- [ ] **Step 1: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors. `scripts/cli.ts` needs no change — `generateCombinedExcel`'s public signature is unchanged, it still only takes `(declaration, outputPath, branding)`.

- [ ] **Step 2: Full test suite**

Run: `npx vitest run`
Expected: all tests pass.

- [ ] **Step 3: Manual smoke test**

Run a real generation through the CLI to visually confirm the third row renders correctly (open the file in Excel/Numbers/LibreOffice and check the letterhead by eye):

```bash
npx tsx scripts/cli.ts <path-to-a-real-liquidation-pdf> <path-to-a-real-dum-pdf> /tmp
```

Open the resulting `/tmp/Declaration.xlsx` and confirm both sheets show a 3-row letterhead ending in "Date de génération : DD/MM/YYYY HH:mm" in a smaller, non-bold font under the two existing rows.

---

### Task 6: Add the "Partager sur WhatsApp" button

**Files:**
- Modify: `src/web/views/upload.html`

- [ ] **Step 1: Add the button's CSS**

In `src/web/views/upload.html`, find the `.success-actions` CSS rules (currently):

```css
    .success-actions { display: flex; gap: 10px; flex-wrap: wrap; }
    .success-actions button { flex: 1; min-width: 130px; }
    .success-actions button { margin-top: 0; }
```

Add immediately after that block:

```css
    /* Hidden by default — only unhidden by JS after feature-detecting
       navigator.canShare({ files: [...] }) support, so it never appears on
       a browser where clicking it could only fail. */
    #whatsappShareBtn { display: none; }
    #whatsappShareBtn.visible { display: flex; }
```

Find the `.cost-exclusive` selector list (currently):

```css
    #successPanel.cost-exclusive .success-icon,
    #successPanel.cost-exclusive .success-title,
    #successPanel.cost-exclusive .success-subtitle,
    #successPanel.cost-exclusive .results-section,
    #successPanel.cost-exclusive #newDeclarationBtn,
    #successPanel.cost-exclusive #downloadAgainBtn,
    #successPanel.cost-exclusive #showResultsBtn { display: none; }
```

Add `#successPanel.cost-exclusive #whatsappShareBtn,` as a new line right after the `#downloadAgainBtn` line, so the WhatsApp button also hides in the cost-only view (same as the download button does):

```css
    #successPanel.cost-exclusive .success-icon,
    #successPanel.cost-exclusive .success-title,
    #successPanel.cost-exclusive .success-subtitle,
    #successPanel.cost-exclusive .results-section,
    #successPanel.cost-exclusive #newDeclarationBtn,
    #successPanel.cost-exclusive #downloadAgainBtn,
    #successPanel.cost-exclusive #whatsappShareBtn,
    #successPanel.cost-exclusive #showResultsBtn { display: none; }
```

- [ ] **Step 2: Add the button's HTML**

Find the `#downloadAgainBtn` button in the `.success-actions` div (currently):

```html
          <button type="button" class="secondary" id="downloadAgainBtn">
            <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 3v10.5M10 13.5l-4-4M10 13.5l4-4M4 16.5h12" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
            Télécharger
          </button>
```

Add immediately after it (still inside `.success-actions`):

```html
          <button type="button" class="secondary" id="whatsappShareBtn">
            <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 2.5c-4.14 0-7.5 3.36-7.5 7.5 0 1.32.35 2.56.96 3.63L2.5 17.5l4-.94A7.47 7.47 0 0 0 10 17.5c4.14 0 7.5-3.36 7.5-7.5S14.14 2.5 10 2.5Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
            WhatsApp
          </button>
```

- [ ] **Step 3: Add the button's JS**

Find where `downloadAgainBtn` is looked up (currently, near the other `document.getElementById` calls):

```js
    const downloadAgainBtn = document.getElementById('downloadAgainBtn');
```

Add immediately after it:

```js
    const whatsappShareBtn = document.getElementById('whatsappShareBtn');
```

Find the `downloadAgainBtn.addEventListener('click', ...)` block (currently):

```js
    downloadAgainBtn.addEventListener('click', async () => {
      downloadAgainBtn.disabled = true;
      try {
        const response = await fetch('/download');
        if (!response.ok) throw new Error('Impossible de retélécharger le fichier.');
        const blob = await response.blob();
        triggerBlobDownload(blob);
      } catch (err) {
        showError(err.message);
      } finally {
        downloadAgainBtn.disabled = false;
      }
    });
```

Add immediately after it:

```js
    const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    function canShareXlsxFile() {
      if (!navigator.canShare) return false;
      try {
        const probe = new File([''], 'probe.xlsx', { type: XLSX_MIME_TYPE });
        return navigator.canShare({ files: [probe] });
      } catch {
        return false;
      }
    }

    if (canShareXlsxFile()) {
      whatsappShareBtn.classList.add('visible');
    }

    whatsappShareBtn.addEventListener('click', async () => {
      whatsappShareBtn.disabled = true;
      try {
        const response = await fetch('/download');
        if (!response.ok) throw new Error('Impossible de préparer le fichier à partager.');
        const blob = await response.blob();
        const file = new File([blob], 'Declaration.xlsx', { type: XLSX_MIME_TYPE });
        await navigator.share({ files: [file], title: 'Déclaration Excel' });
      } catch (err) {
        if (err && err.name === 'AbortError') return; // user cancelled the native share sheet
        showError(err.message || 'Échec du partage — réessayez.');
      } finally {
        whatsappShareBtn.disabled = false;
      }
    });
```

- [ ] **Step 4: Manual smoke test**

No automated test harness covers this file's inline `<script>` (confirmed: no frontend test runner in `package.json`). Verify manually:

1. Start the app: `npm run web` (or the project's existing dev-run command).
2. Open the upload page in Chrome DevTools with device emulation set to a mobile device (or on an actual Android phone) — confirm the "WhatsApp" button appears next to "Télécharger" after generating a declaration.
3. Click it — confirm the native share sheet opens with the `.xlsx` file attached, and selecting WhatsApp from it attaches the file to a new WhatsApp message.
4. Open the same page in desktop Firefox (no file-sharing support) — confirm the WhatsApp button does not appear at all.
5. In a browser where it does appear, click it and cancel the native share sheet — confirm no error message appears (the `AbortError` is swallowed silently).

- [ ] **Step 5: Commit**

```bash
git add src/web/views/upload.html
git commit -m "Add a WhatsApp share button next to Télécharger"
```

---

### Task 7: Final full-repo verification

- [ ] **Step 1: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors.

- [ ] **Step 2: Full test suite**

Run: `npx vitest run`
Expected: all tests pass (no regressions from any of the row-index shifts).

- [ ] **Step 3: Push**

```bash
git push
```

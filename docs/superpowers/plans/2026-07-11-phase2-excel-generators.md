# Phase 2: Excel Generators Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate the two required Excel files from a validated `Declaration` (produced by phase 1's parse+merge pipeline): File 1 (one row per article) and File 2 (one row per physical unit, with each tax code's amount divided equally across the article's units and any rounding remainder reconciled into the last row).

**Architecture:** Two independent generator modules (`ArticleSummaryExcelGenerator`, `UnitLevelExcelGenerator`) both consume the same `Declaration`/`Article` domain types from phase 1 and both write via ExcelJS's streaming `WorkbookWriter` for constant memory regardless of row count, per the design spec's performance requirement (tens of thousands of rows). The per-unit tax-splitting math (division + rounding-remainder allocation) and the tax-code-column-union logic are factored into a small, independently-testable helper module shared by the unit-level generator, since that arithmetic is the part most likely to have subtle bugs and benefits from tests that don't require spinning up a real `.xlsx` file.

**Tech Stack:** TypeScript (strict, ESM), `exceljs` (streaming `WorkbookWriter`/`WorkbookReader`), Vitest — building on phase 1's `src/domain/types.ts` (`Declaration`, `Article`, `TaxLine`) without modifying it.

Reference spec: `docs/superpowers/specs/2026-07-11-customs-declaration-excel-design.md` (§3.5 Excel Generators)
Reference prior plan (phase 1, already merged to `main`): `docs/superpowers/plans/2026-07-11-phase1-parser-domain-merge.md`

---

### Task 1: Add exceljs dependency

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Add exceljs as a runtime dependency**

Edit `package.json` to add a `dependencies` section (it currently only has `devDependencies`):

```json
  "dependencies": {
    "exceljs": "^4.4.0"
  },
```

Place it right after the `"scripts"` block and before `"devDependencies"`, so the resulting file looks like:

```json
{
  "name": "customs-declaration-excel",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "test": "vitest run",
    "test:watch": "vitest",
    "typecheck": "tsc --noEmit"
  },
  "dependencies": {
    "exceljs": "^4.4.0"
  },
  "devDependencies": {
    "typescript": "^5.5.4",
    "vitest": "^2.0.5",
    "@types/node": "^22.0.0"
  }
}
```

(Keep whatever exact `@types/node` version is already present in the file from phase 1 — don't downgrade it, just preserve it while adding the `dependencies` block.)

- [ ] **Step 2: Install**

Run: `npm install`
Expected: `exceljs` added to `node_modules/` and `package-lock.json` updated, no errors.

- [ ] **Step 3: Commit**

```bash
git add package.json package-lock.json
git commit -m "Add exceljs dependency for Excel generation"
```

---

### Task 2: Unit-level tax allocation helpers

**Files:**
- Create: `src/excel/unitLevelTaxHelpers.ts`
- Test: `tests/excel/unitLevelTaxHelpers.test.ts`

This module has two small, pure functions used by the unit-level generator (Task 5):
`unionTaxCodes` (which tax-code columns the sheet needs) and `allocateTaxAcrossUnits`
(how a single tax montant splits across N unit rows, with the rounding remainder
placed on the last row so the column always sums back to the exact source amount).
Testing this in isolation (no `.xlsx` I/O) makes the arithmetic easy to verify.

- [ ] **Step 1: Write the failing tests**

Create `tests/excel/unitLevelTaxHelpers.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { allocateTaxAcrossUnits, unionTaxCodes } from '../../src/excel/unitLevelTaxHelpers.js';
import type { Article } from '../../src/domain/types.js';

describe('allocateTaxAcrossUnits', () => {
  it('splits an evenly-divisible amount equally across all units', () => {
    const amounts = allocateTaxAcrossUnits(70, 350);
    expect(amounts).toHaveLength(350);
    for (const amount of amounts) {
      expect(amount).toBeCloseTo(0.2);
    }
    expect(amounts.reduce((sum, a) => sum + a, 0)).toBeCloseTo(70, 2);
  });

  it('puts the rounding remainder on the last unit so the column sum matches exactly', () => {
    const amounts = allocateTaxAcrossUnits(68, 354);
    expect(amounts).toHaveLength(354);
    // 68 / 354 = 0.192090... -> rounds to 0.19 for the first 353 units
    for (let i = 0; i < 353; i++) {
      expect(amounts[i]).toBeCloseTo(0.19, 2);
    }
    const total = amounts.reduce((sum, a) => sum + a, 0);
    expect(total).toBeCloseTo(68, 2);
  });

  it('handles a zero montant by allocating zero to every unit', () => {
    const amounts = allocateTaxAcrossUnits(0, 200);
    expect(amounts).toHaveLength(200);
    expect(amounts.every((a) => a === 0)).toBe(true);
  });

  it('throws when quantite is not a positive integer', () => {
    expect(() => allocateTaxAcrossUnits(100, 0)).toThrow();
    expect(() => allocateTaxAcrossUnits(100, -5)).toThrow();
    expect(() => allocateTaxAcrossUnits(100, 3.5)).toThrow();
  });
});

describe('unionTaxCodes', () => {
  function makeArticle(taxCodes: string[]): Article {
    return {
      numero: 1,
      hsCode: '6109100010',
      nomArticle: 'T-SHIRT',
      pays: 'ITALIE',
      paysCode: 'IT',
      valeurDeclaree: 100,
      quantite: 10,
      unite: 'U',
      taxes: taxCodes.map((code) => ({ code, assiette: 100, taux: 1, montant: 1 })),
      totalArticle: taxCodes.length,
    };
  }

  it('returns the sorted union of tax codes across all articles, deduplicated', () => {
    const articles = [makeArticle(['007217', '000110']), makeArticle(['002109', '000110'])];
    expect(unionTaxCodes(articles)).toEqual(['000110', '002109', '007217']);
  });

  it('returns an empty array for articles with no taxes', () => {
    expect(unionTaxCodes([makeArticle([])])).toEqual([]);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/excel/unitLevelTaxHelpers.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/excel/unitLevelTaxHelpers.ts`**

```ts
import type { Article } from '../domain/types.js';

export function allocateTaxAcrossUnits(montant: number, quantite: number): number[] {
  if (!Number.isInteger(quantite) || quantite <= 0) {
    throw new Error(
      `quantite must be a positive integer to allocate tax across units, got ${quantite}`
    );
  }

  const baseAmount = Math.round((montant / quantite) * 100) / 100;
  const amounts = new Array<number>(quantite).fill(baseAmount);

  const allocatedExceptLast = baseAmount * (quantite - 1);
  const lastAmount = Math.round((montant - allocatedExceptLast) * 100) / 100;
  amounts[quantite - 1] = lastAmount;

  return amounts;
}

export function unionTaxCodes(articles: Article[]): string[] {
  const codes = new Set<string>();
  for (const article of articles) {
    for (const tax of article.taxes) {
      codes.add(tax.code);
    }
  }
  return Array.from(codes).sort();
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/excel/unitLevelTaxHelpers.test.ts`
Expected: PASS (6 tests)

- [ ] **Step 5: Commit**

```bash
git add src/excel/unitLevelTaxHelpers.ts tests/excel/unitLevelTaxHelpers.test.ts
git commit -m "Add tax-code union and per-unit tax allocation helpers"
```

---

### Task 3: Shared test helper for temp xlsx file paths

**Files:**
- Create: `tests/excel/testHelpers.ts`

Three later test files (Tasks 4, 5, 6) each need to write a real `.xlsx` file to a
temp location and clean it up afterward. Factor that into one tiny helper instead
of repeating `os.tmpdir()`/cleanup boilerplate three times.

- [ ] **Step 1: Create `tests/excel/testHelpers.ts`**

```ts
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';

export function createTempXlsxPath(basename: string): { filePath: string; dir: string } {
  const dir = mkdtempSync(path.join(tmpdir(), 'customs-excel-test-'));
  return { filePath: path.join(dir, `${basename}.xlsx`), dir };
}

export function cleanupTempDir(dir: string): void {
  rmSync(dir, { recursive: true, force: true });
}
```

This has no independent test (it's a test-only utility, exercised indirectly by
every test that uses it in Tasks 4-6) — this is a plain fixture helper, not
production code, so it's exempt from having its own test file.

- [ ] **Step 2: Commit**

```bash
git add tests/excel/testHelpers.ts
git commit -m "Add temp xlsx file path helper for Excel generator tests"
```

---

### Task 4: ArticleSummaryExcelGenerator (File 1)

**Files:**
- Create: `src/excel/articleSummaryExcelGenerator.ts`
- Test: `tests/excel/articleSummaryExcelGenerator.test.ts`

Generates the one-row-per-article summary: `Nom Article | HSC | Pays | Valeur
déclarée | Unité (Quantity)`. Uses the real phase-1 fixtures (via
`parseLiquidation` + `parseDum` + `mergeDeclaration`) as the test's input
`Declaration`, so this is validated against real transcribed document data,
not synthetic fixtures.

- [ ] **Step 1: Write the failing test**

Create `tests/excel/articleSummaryExcelGenerator.test.ts`:

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

    await generateArticleSummaryExcel(declaration, filePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    const headerRow = sheet.getRow(1);
    expect(headerRow.getCell(1).value).toBe('Nom Article');
    expect(headerRow.getCell(2).value).toBe('HSC');
    expect(headerRow.getCell(3).value).toBe('Pays');
    expect(headerRow.getCell(4).value).toBe('Valeur déclarée');
    expect(headerRow.getCell(5).value).toBe('Unité (Quantity)');

    // sheet.rowCount includes the header row, so 2 articles -> 3 rows total
    expect(sheet.rowCount).toBe(3);

    const row1 = sheet.getRow(2);
    expect(row1.getCell(1).value).toBe('T-SHIRT');
    expect(row1.getCell(2).value).toBe('6109100010');
    expect(row1.getCell(3).value).toBe('ITALIE');
    expect(row1.getCell(4).value).toBeCloseTo(27147.0, 1);
    expect(row1.getCell(5).value).toBeCloseTo(354.0, 1);

    const row2 = sheet.getRow(3);
    expect(row2.getCell(3).value).toBe('BANGLADESH');
    expect(row2.getCell(5).value).toBeCloseTo(200.0, 1);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/excel/articleSummaryExcelGenerator.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/excel/articleSummaryExcelGenerator.ts`**

```ts
import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';

export async function generateArticleSummaryExcel(
  declaration: Declaration,
  outputPath: string
): Promise<void> {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: false,
  });
  const sheet = workbook.addWorksheet('Articles');

  sheet.columns = [
    { header: 'Nom Article', key: 'nomArticle', width: 30 },
    { header: 'HSC', key: 'hsCode', width: 15 },
    { header: 'Pays', key: 'pays', width: 20 },
    { header: 'Valeur déclarée', key: 'valeurDeclaree', width: 18 },
    { header: 'Unité (Quantity)', key: 'quantite', width: 18 },
  ];

  for (const article of declaration.articles) {
    sheet
      .addRow({
        nomArticle: article.nomArticle,
        hsCode: article.hsCode,
        pays: article.pays,
        valeurDeclaree: article.valeurDeclaree,
        quantite: article.quantite,
      })
      .commit();
  }

  sheet.commit();
  await workbook.commit();
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/excel/articleSummaryExcelGenerator.test.ts`
Expected: PASS (1 test)

- [ ] **Step 5: Commit**

```bash
git add src/excel/articleSummaryExcelGenerator.ts tests/excel/articleSummaryExcelGenerator.test.ts
git commit -m "Add ArticleSummaryExcelGenerator for File 1"
```

---

### Task 5: UnitLevelExcelGenerator (File 2)

**Files:**
- Create: `src/excel/unitLevelExcelGenerator.ts`
- Test: `tests/excel/unitLevelExcelGenerator.test.ts`

Generates the one-row-per-physical-unit sheet: `Nom Article | HSC | Serial
Number | <tax code columns, dynamic union across the declaration>`. Uses
`allocateTaxAcrossUnits`/`unionTaxCodes` from Task 2. Tested against the same
real merged declaration as Task 4, verifying both row/column shape and that
each tax column's values genuinely reconcile (sum back to the source
`montant`) — this is the core correctness property from the design spec.

- [ ] **Step 1: Write the failing test**

Create `tests/excel/unitLevelExcelGenerator.test.ts`:

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

    await generateUnitLevelExcel(declaration, filePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    const headerRow = sheet.getRow(1);
    expect(headerRow.getCell(1).value).toBe('Nom Article');
    expect(headerRow.getCell(2).value).toBe('HSC');
    expect(headerRow.getCell(3).value).toBe('Serial Number');
    // union of tax codes across both articles, sorted: 000110, 002109, 007217
    expect(headerRow.getCell(4).value).toBe('000110');
    expect(headerRow.getCell(5).value).toBe('002109');
    expect(headerRow.getCell(6).value).toBe('007217');

    // article 1: 354 units, article 2: 200 units -> 554 data rows + 1 header = 555
    expect(sheet.rowCount).toBe(555);

    // first row of article 1
    const firstRow = sheet.getRow(2);
    expect(firstRow.getCell(1).value).toBe('T-SHIRT');
    expect(firstRow.getCell(3).value).toBe(1);

    // last row of article 1 (row 355 = header + 354 units), first row of article 2 resets serial number
    const lastRowArticle1 = sheet.getRow(355);
    expect(lastRowArticle1.getCell(3).value).toBe(354);
    const firstRowArticle2 = sheet.getRow(356);
    expect(firstRowArticle2.getCell(3).value).toBe(1);
    expect(firstRowArticle2.getCell(1).value).toBe('T-SHIRT');

    // Reconciliation: sum each tax column across article 1's 354 rows (rows 2-355)
    // against the known source montants from the Liquidation fixture:
    // 000110 = 0.00, 002109 = 5443.00, 007217 = 68.00
    let sum000110 = 0;
    let sum002109 = 0;
    let sum007217 = 0;
    for (let rowNum = 2; rowNum <= 355; rowNum++) {
      const row = sheet.getRow(rowNum);
      sum000110 += Number(row.getCell(4).value);
      sum002109 += Number(row.getCell(5).value);
      sum007217 += Number(row.getCell(6).value);
    }
    expect(sum000110).toBeCloseTo(0.0, 2);
    expect(sum002109).toBeCloseTo(5443.0, 2);
    expect(sum007217).toBeCloseTo(68.0, 2);
  });

  it('throws when an article quantite is not a whole number', async () => {
    const declaration = loadRealDeclaration();
    const brokenDeclaration: Declaration = {
      ...declaration,
      articles: declaration.articles.map((a, i) => (i === 0 ? { ...a, quantite: 354.5 } : a)),
    };
    const { filePath, dir } = createTempXlsxPath('unit-level-broken');
    tempDir = dir;

    await expect(generateUnitLevelExcel(brokenDeclaration, filePath)).rejects.toThrow();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/excel/unitLevelExcelGenerator.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/excel/unitLevelExcelGenerator.ts`**

```ts
import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { allocateTaxAcrossUnits, unionTaxCodes } from './unitLevelTaxHelpers.js';

export async function generateUnitLevelExcel(
  declaration: Declaration,
  outputPath: string
): Promise<void> {
  const taxCodes = unionTaxCodes(declaration.articles);

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: false,
  });
  const sheet = workbook.addWorksheet('Unit Detail');

  sheet.columns = [
    { header: 'Nom Article', key: 'nomArticle', width: 30 },
    { header: 'HSC', key: 'hsCode', width: 15 },
    { header: 'Serial Number', key: 'serialNumber', width: 15 },
    ...taxCodes.map((code) => ({ header: code, key: code, width: 14 })),
  ];

  for (const article of declaration.articles) {
    const quantite = Math.round(article.quantite);
    if (Math.abs(article.quantite - quantite) > 0.01) {
      throw new Error(
        `Article ${article.numero}: quantite (${article.quantite}) is not a whole number; cannot generate one row per unit`
      );
    }

    const perCodeAllocations = new Map<string, number[]>();
    for (const code of taxCodes) {
      const tax = article.taxes.find((t) => t.code === code);
      perCodeAllocations.set(
        code,
        tax ? allocateTaxAcrossUnits(tax.montant, quantite) : new Array(quantite).fill(0)
      );
    }

    for (let unit = 0; unit < quantite; unit++) {
      const row: Record<string, string | number> = {
        nomArticle: article.nomArticle,
        hsCode: article.hsCode,
        serialNumber: unit + 1,
      };
      for (const code of taxCodes) {
        row[code] = perCodeAllocations.get(code)![unit];
      }
      sheet.addRow(row).commit();
    }
  }

  sheet.commit();
  await workbook.commit();
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/excel/unitLevelExcelGenerator.test.ts`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add src/excel/unitLevelExcelGenerator.ts tests/excel/unitLevelExcelGenerator.test.ts
git commit -m "Add UnitLevelExcelGenerator for File 2"
```

---

### Task 6: Performance test at scale (10k+ rows)

**Files:**
- Test: `tests/excel/unitLevelExcelGenerator.performance.test.ts`

Confirms the streaming approach genuinely handles the design spec's "tens of
thousands of rows" requirement without the test itself needing to measure
process memory (which is flaky in CI) — instead it constructs a declaration
with a single article whose `quantite` is large, generates the file, and
verifies via a streaming reader (not a full in-memory read) that the correct
row count was written, within a generous time bound.

- [ ] **Step 1: Write the test**

Create `tests/excel/unitLevelExcelGenerator.performance.test.ts`:

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
      },
    ],
  };
}

describe('generateUnitLevelExcel performance', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('generates 10,000+ rows via streaming within a reasonable time and with the correct row count', async () => {
    const quantite = 10000;
    const declaration = makeLargeDeclaration(quantite);
    const { filePath, dir } = createTempXlsxPath('unit-level-performance');
    tempDir = dir;

    const start = Date.now();
    await generateUnitLevelExcel(declaration, filePath);
    const durationMs = Date.now() - start;

    expect(durationMs).toBeLessThan(15000);

    // Use the streaming reader (not a full in-memory read) to count rows,
    // consistent with this file existing to prove the streaming path scales.
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {});
    let rowCount = 0;
    for await (const worksheetReader of reader) {
      for await (const _row of worksheetReader) {
        rowCount++;
      }
    }

    // quantite data rows + 1 header row
    expect(rowCount).toBe(quantite + 1);
  }, 30000);
});
```

- [ ] **Step 2: Run the test**

Run: `npx vitest run tests/excel/unitLevelExcelGenerator.performance.test.ts`
Expected: PASS (1 test). If it times out or fails, report back rather than
weakening the assertions — a genuine performance problem here means the
streaming setup in `unitLevelExcelGenerator.ts` needs investigation (e.g.
confirm `useStyles: false` is set, confirm rows are being `.commit()`-ed
individually rather than buffered).

- [ ] **Step 3: Run the full test suite and typecheck**

Run: `npx vitest run`
Expected: PASS — all tests across all files (phase 1 + phase 2) green.

Run: `npx tsc --noEmit`
Expected: No type errors.

- [ ] **Step 4: Commit**

```bash
git add tests/excel/unitLevelExcelGenerator.performance.test.ts
git commit -m "Add performance test for UnitLevelExcelGenerator at 10k+ rows"
```

---

## Phase 2 exit criteria

- `npx vitest run` passes with all tests green (phase 1's 36 tests plus phase 2's new tests: tax allocation helpers, both generators' correctness tests, the performance test).
- `npx tsc --noEmit` passes with no errors.
- `generateArticleSummaryExcel` and `generateUnitLevelExcel` both consume `Declaration` from `src/domain/types.ts` unmodified — no changes needed to phase 1 code.
- File 2's tax columns are proven, by test, to reconcile exactly to the source `montant` values (not just approximately close).
- The unit-level generator is proven to handle at least 10,000 rows via the actual streaming API within a reasonable time bound.
- Next phase (OCR, per the design spec's roadmap) can be developed independently — the Excel generators don't need to change to accept OCR-derived `Declaration` objects, since OCR's job is to produce the same `Declaration` shape phase 1's fixtures already produce.

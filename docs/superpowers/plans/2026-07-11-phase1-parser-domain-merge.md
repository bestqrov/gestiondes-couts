# Phase 1: Domain + Parsers + Merge Layer — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build and test, against the two real sample documents (Liquidation + DUM), the deterministic parsing and merging pipeline that turns their OCR/text-layer output into a validated `Declaration` domain object — no OCR, no Excel generation, no UI yet.

**Architecture:** Two independent line-based parsers (one per document type) produce raw per-document results; a `DeclarationMerger` joins them by reference code + article number into `Article` domain entities; `validateArticle` enforces business rules (quantity > 0, tax sum reconciles to total). Everything is tested against fixture text files transcribed from the two real sample documents, so no OCR dependency is needed for this phase.

**Tech Stack:** TypeScript (strict, ESM), Vitest for tests, Node.js — no other runtime dependencies in this phase.

Reference spec: `docs/superpowers/specs/2026-07-11-customs-declaration-excel-design.md`

---

### Task 1: Project scaffolding

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `vitest.config.ts`
- Create: `.gitignore`

- [ ] **Step 1: Create `package.json`**

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
  "devDependencies": {
    "typescript": "^5.5.4",
    "vitest": "^2.0.5"
  }
}
```

- [ ] **Step 2: Create `tsconfig.json`**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "moduleResolution": "Bundler",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "outDir": "dist",
    "declaration": true,
    "resolveJsonModule": true
  },
  "include": ["src", "tests"]
}
```

- [ ] **Step 3: Create `vitest.config.ts`**

```ts
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
  },
});
```

- [ ] **Step 4: Create `.gitignore`**

```
node_modules/
dist/
*.xlsx
.DS_Store
```

- [ ] **Step 5: Install dependencies**

Run: `cd "/Users/mac/Documents/gestion de couts" && npm install`
Expected: `node_modules/` created, `package-lock.json` created, no errors.

- [ ] **Step 6: Commit**

```bash
git add package.json package-lock.json tsconfig.json vitest.config.ts .gitignore
git commit -m "Scaffold TypeScript project with Vitest"
```

---

### Task 2: Domain types and validation

**Files:**
- Create: `src/domain/types.ts`
- Create: `src/domain/validators.ts`
- Test: `tests/domain/validators.test.ts`

- [ ] **Step 1: Create `src/domain/types.ts`**

```ts
export interface TaxLine {
  code: string;
  assiette: number;
  taux: number;
  montant: number;
}

export interface Article {
  numero: number;
  hsCode: string;
  nomArticle: string;
  pays: string;
  paysCode: string;
  valeurDeclaree: number;
  quantite: number;
  unite: string;
  taxes: TaxLine[];
  totalArticle: number;
}

export interface Declaration {
  code: string;
  redevable: string;
  benNumero: string;
  articles: Article[];
}
```

- [ ] **Step 2: Write the failing test for validation**

Create `tests/domain/validators.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { validateArticle, ValidationError } from '../../src/domain/validators.js';
import type { Article } from '../../src/domain/types.js';

function makeArticle(overrides: Partial<Article> = {}): Article {
  return {
    numero: 1,
    hsCode: '6109100010',
    nomArticle: 'T-SHIRT',
    pays: 'ITALIE',
    paysCode: 'IT',
    valeurDeclaree: 27147,
    quantite: 354,
    unite: 'NOMBRE',
    taxes: [
      { code: '000110', assiette: 27147, taux: 0, montant: 0 },
      { code: '007217', assiette: 27147, taux: 0.25, montant: 68 },
      { code: '002109', assiette: 27215, taux: 20, montant: 5443 },
    ],
    totalArticle: 5511,
    ...overrides,
  };
}

describe('validateArticle', () => {
  it('passes for a well-formed article whose tax sum matches totalArticle', () => {
    expect(() => validateArticle(makeArticle())).not.toThrow();
  });

  it('throws when quantite is zero', () => {
    expect(() => validateArticle(makeArticle({ quantite: 0 }))).toThrow(ValidationError);
  });

  it('throws when tax montants do not sum to totalArticle', () => {
    expect(() => validateArticle(makeArticle({ totalArticle: 9999 }))).toThrow(ValidationError);
  });
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npx vitest run tests/domain/validators.test.ts`
Expected: FAIL — `src/domain/validators.js` (or `.ts`) module not found.

- [ ] **Step 4: Create `src/domain/validators.ts`**

```ts
import type { Article } from './types.js';

export class ValidationError extends Error {
  constructor(message: string, public readonly field: string) {
    super(message);
    this.name = 'ValidationError';
  }
}

const ROUNDING_TOLERANCE = 0.02;

export function validateArticle(article: Article): void {
  if (article.quantite <= 0) {
    throw new ValidationError(
      `Article ${article.numero}: quantite must be > 0, got ${article.quantite}`,
      'quantite'
    );
  }

  const taxSum = article.taxes.reduce((sum, tax) => sum + tax.montant, 0);
  const diff = Math.abs(taxSum - article.totalArticle);
  if (diff > ROUNDING_TOLERANCE) {
    throw new ValidationError(
      `Article ${article.numero}: sum of tax montants (${taxSum.toFixed(2)}) ` +
        `does not match totalArticle (${article.totalArticle.toFixed(2)})`,
      'totalArticle'
    );
  }
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `npx vitest run tests/domain/validators.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 6: Commit**

```bash
git add src/domain/types.ts src/domain/validators.ts tests/domain/validators.test.ts
git commit -m "Add domain types and article validation"
```

---

### Task 3: Shared text-parsing utilities

**Files:**
- Create: `src/parser/shared/text.ts`
- Test: `tests/parser/shared/text.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/parser/shared/text.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { parseFrenchNumber, extractFirst } from '../../../src/parser/shared/text.js';

describe('parseFrenchNumber', () => {
  it('parses space-separated thousands with comma decimal', () => {
    expect(parseFrenchNumber('27 147,00')).toBeCloseTo(27147.0);
  });

  it('parses plain decimal with dot', () => {
    expect(parseFrenchNumber('354.000')).toBeCloseTo(354.0);
  });

  it('throws on unparseable input', () => {
    expect(() => parseFrenchNumber('abc')).toThrow();
  });
});

describe('extractFirst', () => {
  it('returns the first capture group trimmed', () => {
    expect(extractFirst('REDEVABLE :  GLOBAL TRADE LOGISTICS SARL\n', /REDEVABLE\s*:\s*(.+)/)).toBe(
      'GLOBAL TRADE LOGISTICS SARL'
    );
  });

  it('returns undefined when no match', () => {
    expect(extractFirst('no match here', /FOO\s*:\s*(.+)/)).toBeUndefined();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/parser/shared/text.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/parser/shared/text.ts`**

```ts
export function parseFrenchNumber(raw: string): number {
  const cleaned = raw
    .replace(/\s/g, '')
    .replace(/ /g, '')
    .replace(',', '.');
  const value = Number.parseFloat(cleaned);
  if (Number.isNaN(value)) {
    throw new Error(`Cannot parse number from "${raw}"`);
  }
  return value;
}

export function extractFirst(text: string, pattern: RegExp): string | undefined {
  const match = text.match(pattern);
  return match ? match[1].trim() : undefined;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/parser/shared/text.test.ts`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add src/parser/shared/text.ts tests/parser/shared/text.test.ts
git commit -m "Add shared French-number and label-extraction parsing helpers"
```

---

### Task 4: LiquidationParser

**Files:**
- Create: `src/parser/liquidation/liquidationParser.ts`
- Create: `tests/parser/fixtures/liquidation-sample-1.txt`
- Test: `tests/parser/liquidation/liquidationParser.test.ts`

This parser reads the fixed-width tax table format shown in the real Liquidation
sample. Each tax row is a `!`-delimited line with 7 fields: `code`, `assiette`,
`taux`, `s.tva`, `s.fr`, `taux virtuel`, `montant`. Splitting the trimmed line on
`!` gives 9 array elements (empty string before the first `!` and after the last
one), so `code` is index 1 and `montant` is index 7.

- [ ] **Step 1: Create the fixture file `tests/parser/fixtures/liquidation-sample-1.txt`**

```
Type Intervenant :  Operateur
REDEVABLE :  GLOBAL TRADE LOGISTICS SARL
CATEGORIE D'ORDONNANCEMENT :  Credit d'enlevement
CODE : 500001
B E N° : 501 DU : 25/06/2026

N° ET DATE DECLARATION : 700000000000000001 DU : 25/06/2026
N° ET DATE LIQUIDATION : 700000CEE00000000001 DU : 25/06/2026
DATE ECHEANCE PAIEMENT : 24/07/2026

ARTICLE  : 1              NUMERO SH : 6109100010     VALEUR :   27 147,00
QUANTITE : 354.000                UNITE : NOMBRE

TAXE   ! ASSIETTE  ! TAUX ! S.TVA ! S.FR ! TAUX VIRTUEL !  MONTANT
! 000110 !  27147.00 !  0.0 !   T   !      !              !     0,00 !
! 007217 !  27147.00 !  0.25!   T   !      !              !    68,00 !
! 002109 !  27215.00 ! 20.0 !       !      !              !  5 443,00 !
TOTAL ARTICLE :          5 511,00

ARTICLE  : 2              NUMERO SH : 6109100010     VALEUR :   12 892,99
QUANTITE : 200.000                UNITE : NOMBRE

TAXE   ! ASSIETTE  ! TAUX ! S.TVA ! S.FR ! TAUX VIRTUEL !  MONTANT
! 000110 !  12892.99 ! 30.0 !   T   !      !              !  3 868,00 !
! 007217 !  12892.99 !  0.25!   T   !      !              !    33,00 !
! 002109 !  16793.99 ! 20.0 !       !      !              !  3 359,00 !
TOTAL ARTICLE :          7 260,00
```

- [ ] **Step 2: Write the failing test**

Create `tests/parser/liquidation/liquidationParser.test.ts`:

```ts
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseLiquidation } from '../../../src/parser/liquidation/liquidationParser.js';

const fixturePath = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  '../fixtures/liquidation-sample-1.txt'
);

describe('parseLiquidation', () => {
  it('parses header and both articles from the real sample document', () => {
    const text = readFileSync(fixturePath, 'utf-8');
    const result = parseLiquidation(text);

    expect(result.header).toEqual({
      code: '500001',
      redevable: 'GLOBAL TRADE LOGISTICS SARL',
      benNumero: '501',
    });

    expect(result.articles).toHaveLength(2);

    const [article1, article2] = result.articles;

    expect(article1.numero).toBe(1);
    expect(article1.hsCode).toBe('6109100010');
    expect(article1.valeurDeclaree).toBeCloseTo(27147.0);
    expect(article1.quantite).toBeCloseTo(354.0);
    expect(article1.unite).toBe('NOMBRE');
    expect(article1.totalArticle).toBeCloseTo(5511.0);
    expect(article1.taxes).toEqual([
      { code: '000110', assiette: 27147.0, taux: 0.0, montant: 0.0 },
      { code: '007217', assiette: 27147.0, taux: 0.25, montant: 68.0 },
      { code: '002109', assiette: 27215.0, taux: 20.0, montant: 5443.0 },
    ]);

    expect(article2.numero).toBe(2);
    expect(article2.valeurDeclaree).toBeCloseTo(12892.99);
    expect(article2.quantite).toBeCloseTo(200.0);
    expect(article2.totalArticle).toBeCloseTo(7260.0);
    expect(article2.taxes.map((t) => t.code)).toEqual(['000110', '007217', '002109']);
  });

  it('throws when no articles are found', () => {
    expect(() => parseLiquidation('CODE : 123\nREDEVABLE : X\nB E N° : 1')).toThrow(
      'No articles found'
    );
  });
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npx vitest run tests/parser/liquidation/liquidationParser.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 4: Create `src/parser/liquidation/liquidationParser.ts`**

```ts
import type { TaxLine } from '../../domain/types.js';
import { extractFirst, parseFrenchNumber } from '../shared/text.js';

export interface LiquidationHeader {
  code: string;
  redevable: string;
  benNumero: string;
}

export interface LiquidationArticleResult {
  numero: number;
  hsCode: string;
  valeurDeclaree: number;
  quantite: number;
  unite: string;
  taxes: TaxLine[];
  totalArticle: number;
}

export interface LiquidationResult {
  header: LiquidationHeader;
  articles: LiquidationArticleResult[];
}

const ARTICLE_BLOCK_PATTERN = /ARTICLE\s*:\s*(\d+)([\s\S]*?)(?=ARTICLE\s*:\s*\d+|$)/g;

function parseTaxRows(block: string): TaxLine[] {
  const taxes: TaxLine[] = [];
  for (const line of block.split('\n')) {
    const trimmed = line.trim();
    if (!/^!\s*\d{6}\s*!/.test(trimmed)) continue;
    const parts = trimmed.split('!').map((part) => part.trim());
    const [, code, assietteRaw, tauxRaw, , , , montantRaw] = parts;
    taxes.push({
      code,
      assiette: parseFrenchNumber(assietteRaw),
      taux: parseFrenchNumber(tauxRaw),
      montant: parseFrenchNumber(montantRaw),
    });
  }
  return taxes;
}

export function parseLiquidation(text: string): LiquidationResult {
  const code = extractFirst(text, /CODE\s*:\s*(\d+)/);
  const redevable = extractFirst(text, /REDEVABLE\s*:\s*(.+)/);
  const benNumero = extractFirst(text, /B E N°\s*:\s*(\S+)/);

  if (!code || !redevable || !benNumero) {
    throw new Error('Liquidation header fields (CODE, REDEVABLE, B E N°) not found');
  }

  const articles: LiquidationArticleResult[] = [];

  for (const match of text.matchAll(ARTICLE_BLOCK_PATTERN)) {
    const numero = Number.parseInt(match[1], 10);
    const block = match[2];

    const hsCode = extractFirst(block, /NUMERO SH\s*:\s*(\d+)/);
    const valeurRaw = extractFirst(block, /VALEUR\s*:\s*([\d\s.,]+)/);
    const quantiteRaw = extractFirst(block, /QUANTITE\s*:\s*([\d.]+)/);
    const unite = extractFirst(block, /UNITE\s*:\s*(\S+)/);
    const totalRaw = extractFirst(block, /TOTAL ARTICLE\s*:\s*([\d\s.,]+)/);

    if (!hsCode || !valeurRaw || !quantiteRaw || !unite || !totalRaw) {
      throw new Error(`Article ${numero}: missing one of HS code / valeur / quantite / unite / total`);
    }

    const taxes = parseTaxRows(block);
    if (taxes.length === 0) {
      throw new Error(`Article ${numero}: no tax rows found`);
    }

    articles.push({
      numero,
      hsCode,
      valeurDeclaree: parseFrenchNumber(valeurRaw),
      quantite: parseFrenchNumber(quantiteRaw),
      unite,
      taxes,
      totalArticle: parseFrenchNumber(totalRaw),
    });
  }

  if (articles.length === 0) {
    throw new Error('No articles found in Liquidation document');
  }

  return { header: { code, redevable, benNumero }, articles };
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `npx vitest run tests/parser/liquidation/liquidationParser.test.ts`
Expected: PASS (2 tests)

- [ ] **Step 6: Commit**

```bash
git add src/parser/liquidation/liquidationParser.ts tests/parser/liquidation/liquidationParser.test.ts tests/parser/fixtures/liquidation-sample-1.txt
git commit -m "Add LiquidationParser with fixture test from real sample document"
```

---

### Task 5: DumParser

**Files:**
- Create: `src/parser/dum/dumParser.ts`
- Create: `tests/parser/fixtures/dum-sample-1.txt`
- Test: `tests/parser/dum/dumParser.test.ts`

The DUM fixture is a cleaned, canonically-ordered transcription of the fields
from the real sample DUM PDF (raw PDF text-layer extraction interleaves labels
and values inconsistently across the multi-column form; §8 of the design spec
flags that these extraction rules will be refined once more real samples are
available). The parser here matches on labeled `Label : value` lines, which is
what a table-aware text extraction (phase 3) is expected to produce.

- [ ] **Step 1: Create the fixture file `tests/parser/fixtures/dum-sample-1.txt`**

```
DECLARATION

N° d'ordre de l'art. : 1
Colis et désignation des marchandises : T-SHIRT 354.00 NB
Code marchandises : 6109100010
Valeur déclarée : 27 147.000
Unités complémentaires : 354.0 U
Pays d'origine (Nom et code) : ITALIE IT

N° d'ordre de l'art. : 2
Colis et désignation des marchandises : T-SHIRT 200.00 NB
Code marchandises : 6109100010
Valeur déclarée : 12 892.992
Unités complémentaires : 200.0 U
Pays d'origine (Nom et code) : BANGLADESH BD

Renseignements financiers et bancaires : 02 Crédit d'enlèvement 500001099
```

- [ ] **Step 2: Write the failing test**

Create `tests/parser/dum/dumParser.test.ts`:

```ts
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseDum } from '../../../src/parser/dum/dumParser.js';

const fixturePath = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  '../fixtures/dum-sample-1.txt'
);

describe('parseDum', () => {
  it('parses the crédit d\'enlèvement code and both articles from the real sample document', () => {
    const text = readFileSync(fixturePath, 'utf-8');
    const result = parseDum(text);

    expect(result.creditEnlevementCode).toBe('500001099');
    expect(result.articles).toHaveLength(2);

    const [article1, article2] = result.articles;

    expect(article1).toEqual({
      ordre: 1,
      hsCode: '6109100010',
      nomArticle: 'T-SHIRT',
      paysNom: 'ITALIE',
      paysCode: 'IT',
      valeurDeclaree: 27147.0,
      quantite: 354.0,
      unite: 'U',
    });

    expect(article2).toEqual({
      ordre: 2,
      hsCode: '6109100010',
      nomArticle: 'T-SHIRT',
      paysNom: 'BANGLADESH',
      paysCode: 'BD',
      valeurDeclaree: 12892.992,
      quantite: 200.0,
      unite: 'U',
    });
  });

  it('throws when no articles are found', () => {
    expect(() => parseDum("Crédit d'enlèvement 123")).toThrow('No articles found');
  });
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npx vitest run tests/parser/dum/dumParser.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 4: Create `src/parser/dum/dumParser.ts`**

```ts
import { extractFirst, parseFrenchNumber } from '../shared/text.js';

export interface DumArticleResult {
  ordre: number;
  hsCode: string;
  nomArticle: string;
  paysNom: string;
  paysCode: string;
  valeurDeclaree: number;
  quantite: number;
  unite: string;
}

export interface DumResult {
  creditEnlevementCode: string;
  articles: DumArticleResult[];
}

const ARTICLE_BLOCK_PATTERN =
  /N°\s*d'ordre de l'art\.\s*:\s*(\d+)([\s\S]*?)(?=N°\s*d'ordre de l'art\.\s*:\s*\d+|Renseignements financiers|$)/g;

export function parseDum(text: string): DumResult {
  const creditEnlevementCode = extractFirst(text, /Crédit d'enlèvement\D*(\d+)/);
  if (!creditEnlevementCode) {
    throw new Error("Crédit d'enlèvement code not found in DUM document");
  }

  const articles: DumArticleResult[] = [];

  for (const match of text.matchAll(ARTICLE_BLOCK_PATTERN)) {
    const ordre = Number.parseInt(match[1], 10);
    const block = match[2];

    const designationLine = extractFirst(block, /Colis et désignation des marchandises\s*:\s*(.+)/);
    const hsCode = extractFirst(block, /Code marchandises\s*:\s*(\d+)/);
    const valeurRaw = extractFirst(block, /Valeur déclarée\s*:\s*([\d\s.,]+)/);
    const quantiteMatch = block.match(/Unités complémentaires\s*:\s*([\d.,]+)\s*(\S+)/);
    const paysMatch = block.match(/Pays d'origine \(Nom et code\)\s*:\s*(\S+)\s+(\S+)/);

    if (!designationLine || !hsCode || !valeurRaw || !quantiteMatch || !paysMatch) {
      throw new Error(
        `DUM article ${ordre}: missing one of designation / HS code / valeur / quantite / pays`
      );
    }

    const designationMatch = designationLine.match(/^([A-Za-zÀ-ÿ\-\s]+?)\s+[\d.,]+\s+\S+$/);
    const nomArticle = designationMatch ? designationMatch[1].trim() : designationLine.trim();

    articles.push({
      ordre,
      hsCode,
      nomArticle,
      paysNom: paysMatch[1],
      paysCode: paysMatch[2],
      valeurDeclaree: parseFrenchNumber(valeurRaw),
      quantite: parseFrenchNumber(quantiteMatch[1]),
      unite: quantiteMatch[2],
    });
  }

  if (articles.length === 0) {
    throw new Error('No articles found in DUM document');
  }

  return { creditEnlevementCode, articles };
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `npx vitest run tests/parser/dum/dumParser.test.ts`
Expected: PASS (2 tests)

- [ ] **Step 6: Commit**

```bash
git add src/parser/dum/dumParser.ts tests/parser/dum/dumParser.test.ts tests/parser/fixtures/dum-sample-1.txt
git commit -m "Add DumParser with fixture test from real sample document"
```

---

### Task 6: DeclarationMerger

**Files:**
- Create: `src/merge/declarationMerger.ts`
- Test: `tests/merge/declarationMerger.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/merge/declarationMerger.test.ts`:

```ts
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration, MergeError } from '../../src/merge/declarationMerger.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return { liquidation, dum };
}

describe('mergeDeclaration', () => {
  it('merges the real Liquidation and DUM samples into a single Declaration', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const declaration = mergeDeclaration(liquidation, dum);

    expect(declaration.code).toBe('500001');
    expect(declaration.redevable).toBe('GLOBAL TRADE LOGISTICS SARL');
    expect(declaration.articles).toHaveLength(2);

    const [article1, article2] = declaration.articles;

    expect(article1.nomArticle).toBe('T-SHIRT');
    expect(article1.pays).toBe('ITALIE');
    expect(article1.paysCode).toBe('IT');
    expect(article1.hsCode).toBe('6109100010');
    expect(article1.quantite).toBeCloseTo(354.0);
    expect(article1.taxes.map((t) => t.code)).toEqual(['000110', '007217', '002109']);
    expect(article1.totalArticle).toBeCloseTo(5511.0);

    expect(article2.pays).toBe('BANGLADESH');
    expect(article2.paysCode).toBe('BD');
    expect(article2.quantite).toBeCloseTo(200.0);
  });

  it('throws when the DUM crédit d\'enlèvement code does not match the Liquidation code', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const mismatchedDum = { ...dum, creditEnlevementCode: '999999999' };

    expect(() => mergeDeclaration(liquidation, mismatchedDum)).toThrow(MergeError);
  });

  it('throws when an article present in Liquidation is missing from DUM', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const dumMissingArticle = { ...dum, articles: dum.articles.slice(0, 1) };

    expect(() => mergeDeclaration(liquidation, dumMissingArticle)).toThrow(MergeError);
  });

  it('throws when declared value differs beyond tolerance between the two documents', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const skewedDum = {
      ...dum,
      articles: dum.articles.map((a) => (a.ordre === 1 ? { ...a, valeurDeclaree: 999999 } : a)),
    };

    expect(() => mergeDeclaration(liquidation, skewedDum)).toThrow(MergeError);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/merge/declarationMerger.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/merge/declarationMerger.ts`**

```ts
import type { Article, Declaration, TaxLine } from '../domain/types.js';
import type { LiquidationResult } from '../parser/liquidation/liquidationParser.js';
import type { DumResult } from '../parser/dum/dumParser.js';

const VALUE_TOLERANCE = 0.5;

export class MergeError extends Error {}

export function mergeDeclaration(liquidation: LiquidationResult, dum: DumResult): Declaration {
  if (!dum.creditEnlevementCode.startsWith(liquidation.header.code)) {
    throw new MergeError(
      `Liquidation code "${liquidation.header.code}" does not match DUM crédit d'enlèvement code "${dum.creditEnlevementCode}"`
    );
  }

  const dumByOrdre = new Map(dum.articles.map((article) => [article.ordre, article]));

  const articles: Article[] = liquidation.articles.map((liqArticle) => {
    const dumArticle = dumByOrdre.get(liqArticle.numero);
    if (!dumArticle) {
      throw new MergeError(`Article ${liqArticle.numero} present in Liquidation but not found in DUM`);
    }

    if (dumArticle.hsCode !== liqArticle.hsCode) {
      throw new MergeError(
        `Article ${liqArticle.numero}: HS code mismatch (Liquidation "${liqArticle.hsCode}" vs DUM "${dumArticle.hsCode}")`
      );
    }

    if (Math.abs(dumArticle.valeurDeclaree - liqArticle.valeurDeclaree) > VALUE_TOLERANCE) {
      throw new MergeError(
        `Article ${liqArticle.numero}: valeur déclarée mismatch (Liquidation ${liqArticle.valeurDeclaree} vs DUM ${dumArticle.valeurDeclaree})`
      );
    }

    if (Math.abs(dumArticle.quantite - liqArticle.quantite) > VALUE_TOLERANCE) {
      throw new MergeError(
        `Article ${liqArticle.numero}: quantité mismatch (Liquidation ${liqArticle.quantite} vs DUM ${dumArticle.quantite})`
      );
    }

    const taxes: TaxLine[] = liqArticle.taxes.map((tax) => ({ ...tax }));

    return {
      numero: liqArticle.numero,
      hsCode: liqArticle.hsCode,
      nomArticle: dumArticle.nomArticle,
      pays: dumArticle.paysNom,
      paysCode: dumArticle.paysCode,
      valeurDeclaree: dumArticle.valeurDeclaree,
      quantite: dumArticle.quantite,
      unite: dumArticle.unite,
      taxes,
      totalArticle: liqArticle.totalArticle,
    };
  });

  const dumOnlyOrdres = dum.articles
    .map((article) => article.ordre)
    .filter((ordre) => !liquidation.articles.some((a) => a.numero === ordre));

  if (dumOnlyOrdres.length > 0) {
    throw new MergeError(`Article(s) ${dumOnlyOrdres.join(', ')} present in DUM but not found in Liquidation`);
  }

  return {
    code: liquidation.header.code,
    redevable: liquidation.header.redevable,
    benNumero: liquidation.header.benNumero,
    articles,
  };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/merge/declarationMerger.test.ts`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add src/merge/declarationMerger.ts tests/merge/declarationMerger.test.ts
git commit -m "Add DeclarationMerger joining Liquidation and DUM by code and article number"
```

---

### Task 7: End-to-end integration test

**Files:**
- Test: `tests/integration/parseAndMerge.test.ts`

This ties parsing, merging, and domain validation together on the two real
sample fixtures, proving the full phase-1 pipeline produces a valid
`Declaration`.

- [ ] **Step 1: Write the integration test**

Create `tests/integration/parseAndMerge.test.ts`:

```ts
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { validateArticle } from '../../src/domain/validators.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

describe('parse + merge + validate pipeline', () => {
  it('produces a fully valid Declaration from the two real sample documents', () => {
    const liquidationText = readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8');
    const dumText = readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8');

    const liquidation = parseLiquidation(liquidationText);
    const dum = parseDum(dumText);
    const declaration = mergeDeclaration(liquidation, dum);

    expect(declaration.articles).toHaveLength(2);

    for (const article of declaration.articles) {
      expect(() => validateArticle(article)).not.toThrow();
    }

    const totalDeclaredValue = declaration.articles.reduce((sum, a) => sum + a.valeurDeclaree, 0);
    expect(totalDeclaredValue).toBeCloseTo(27147.0 + 12892.992, 1);
  });
});
```

- [ ] **Step 2: Run the full test suite**

Run: `npx vitest run`
Expected: PASS — all tests across all files (validators, text utils, both parsers, merger, integration) green.

- [ ] **Step 3: Run typecheck**

Run: `npx tsc --noEmit`
Expected: No type errors.

- [ ] **Step 4: Commit**

```bash
git add tests/integration/parseAndMerge.test.ts
git commit -m "Add end-to-end integration test for parse + merge + validate pipeline"
```

---

## Phase 1 exit criteria

- `npx vitest run` passes with all tests green (validators, text utils, LiquidationParser, DumParser, DeclarationMerger, integration).
- `npx tsc --noEmit` passes with no errors.
- The two real sample documents (transcribed as fixtures) parse and merge into a valid `Declaration` with correct `nomArticle`, `pays`, `hsCode`, `valeurDeclaree`, `quantite`, and per-code `taxes` for both articles.
- Next phase (Excel Generators, per the design spec's roadmap) consumes `Declaration`/`Article` from `src/domain/types.ts` as-is — no further changes needed to this phase's code to start it.
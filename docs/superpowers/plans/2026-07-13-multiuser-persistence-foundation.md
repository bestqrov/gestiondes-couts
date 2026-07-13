# Multi-User Persistence Foundation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the backend foundation for multi-admin support — SQLite persistence, role-aware authentication (admin/superadmin), extended DUM shipment-cost extraction, and total-landed-cost-per-unit calculation — fully tested independently of any dashboard UI.

**Architecture:** A SQLite database (`better-sqlite3`) holds `users`, `declarations`, and `declaration_articles`. `src/db/*Repository.ts` modules are the only code that touches SQL. `auth.ts` is rewritten to check credentials against the `users` table (bcrypt-hashed passwords) and carry a `role` on each session, replacing the single hardcoded credential pair. `DumParser` gains extraction for shipment-level cost fields already present in the DUM text (currency, invoiced amount, exchange rate, freight, insurance, total declared value) that were not previously parsed. A new `costCalculator` module computes total landed cost per unit per article from those fields plus the existing tax breakdown. This plan wires the new auth into `server.ts`'s existing login/upload/generate/download flow (so the app keeps working end-to-end throughout) but does **not** build the admin/superadmin dashboard UI or the declaration-history save flow — that's a follow-up plan that consumes the repositories and calculator built here.

**Tech Stack:** `better-sqlite3` (synchronous SQLite driver), `bcryptjs` (pure-JS password hashing, no native build step), TypeScript, Vitest.

Reference spec: `docs/superpowers/specs/2026-07-13-multiuser-persistence-dashboards-design.md`

---

### Task 1: Add dependencies

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Add `better-sqlite3` and `bcryptjs` as dependencies, their types as devDependencies**

Edit `package.json`'s `dependencies` block to add:
```json
    "better-sqlite3": "^11.3.0",
    "bcryptjs": "^2.4.3",
```
(insert alphabetically among the existing entries)

Edit `package.json`'s `devDependencies` block to add:
```json
    "@types/better-sqlite3": "^7.6.11",
    "@types/bcryptjs": "^2.4.6",
```
(insert alphabetically among the existing entries)

- [ ] **Step 2: Install**

Run: `npm install`
Expected: no errors; `better-sqlite3` compiles its native binding during install (this is expected and normal — it's a well-established package with prebuilt binaries for common platforms).

- [ ] **Step 3: Commit**

```bash
git add package.json package-lock.json
git commit -m "Add better-sqlite3 and bcryptjs dependencies"
```

---

### Task 2: Database connection and schema

**Files:**
- Create: `src/db/database.ts`
- Test: `tests/db/database.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/db/database.test.ts`:

```ts
import { rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';

describe('createDatabase', () => {
  it('creates the users, declarations, and declaration_articles tables', () => {
    const db = createDatabase(':memory:');

    const tables = db
      .prepare("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
      .all()
      .map((row) => (row as { name: string }).name);

    expect(tables).toEqual(['declaration_articles', 'declarations', 'users']);
    db.close();
  });

  it('does not throw when migrations run twice against the same database file', () => {
    const dbPath = path.join(tmpdir(), `customs-app-test-db-${Date.now()}.sqlite`);
    const db1 = createDatabase(dbPath);
    db1.close();

    expect(() => {
      const db2 = createDatabase(dbPath);
      db2.close();
    }).not.toThrow();

    rmSync(dbPath, { force: true });
    rmSync(`${dbPath}-wal`, { force: true });
    rmSync(`${dbPath}-shm`, { force: true });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/db/database.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/db/database.ts`**

```ts
import Database from 'better-sqlite3';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export function createDatabase(dbPath: string): Database.Database {
  const db = new Database(dbPath);
  if (dbPath !== ':memory:') {
    db.pragma('journal_mode = WAL');
  }
  runMigrations(db);
  return db;
}

function runMigrations(db: Database.Database): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL CHECK (role IN ('admin', 'superadmin')),
      created_at TEXT NOT NULL,
      disabled_at TEXT
    );

    CREATE TABLE IF NOT EXISTS declarations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      owner_user_id INTEGER NOT NULL REFERENCES users(id),
      code TEXT NOT NULL,
      redevable TEXT NOT NULL,
      ben_numero TEXT NOT NULL,
      devise TEXT,
      montant_facture REAL,
      taux_change REAL,
      fret REAL,
      assurance REAL,
      valeur_totale_declaree REAL,
      colis_count INTEGER,
      reference_interne TEXT,
      total_landed_cost REAL NOT NULL,
      cost_estimate_partial INTEGER NOT NULL DEFAULT 0,
      excel_file_path TEXT NOT NULL,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS declaration_articles (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      declaration_id INTEGER NOT NULL REFERENCES declarations(id),
      numero INTEGER NOT NULL,
      hs_code TEXT NOT NULL,
      nom_article TEXT NOT NULL,
      pays TEXT NOT NULL,
      valeur_declaree REAL NOT NULL,
      quantite REAL NOT NULL,
      total_article REAL NOT NULL,
      cost_per_unit REAL NOT NULL,
      taxes_json TEXT NOT NULL
    );
  `);
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_DB_PATH = path.resolve(__dirname, '../../data.db');

let dbInstance: Database.Database | undefined;

export function getDatabase(): Database.Database {
  if (!dbInstance) {
    dbInstance = createDatabase(process.env.DATABASE_PATH ?? DEFAULT_DB_PATH);
  }
  return dbInstance;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/db/database.test.ts`
Expected: PASS (2 tests)

- [ ] **Step 5: Add `data.db*` to `.gitignore`**

Edit `.gitignore`, add a new line:
```
data.db*
```

- [ ] **Step 6: Commit**

```bash
git add src/db/database.ts tests/db/database.test.ts .gitignore
git commit -m "Add SQLite database connection and schema migrations"
```

---

### Task 3: UsersRepository

**Files:**
- Create: `src/db/usersRepository.ts`
- Test: `tests/db/usersRepository.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/db/usersRepository.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';
import {
  createUser,
  findUserByUsername,
  verifyPassword,
  listUsers,
  setUserDisabled,
  seedSuperAdminIfEmpty,
} from '../../src/db/usersRepository.js';

describe('usersRepository', () => {
  it('creates a user with a hashed password and finds it back by username', () => {
    const db = createDatabase(':memory:');

    const created = createUser(db, 'alice', 'hunter2', 'admin');
    expect(created.username).toBe('alice');
    expect(created.role).toBe('admin');
    expect(created.disabledAt).toBeNull();

    const found = findUserByUsername(db, 'alice');
    expect(found).toBeDefined();
    expect(found!.passwordHash).not.toBe('hunter2'); // must be hashed, not plaintext
    expect(verifyPassword(found!.passwordHash, 'hunter2')).toBe(true);
    expect(verifyPassword(found!.passwordHash, 'wrong-password')).toBe(false);

    db.close();
  });

  it('returns undefined for an unknown username', () => {
    const db = createDatabase(':memory:');
    expect(findUserByUsername(db, 'nobody')).toBeUndefined();
    db.close();
  });

  it('lists users ordered by creation time, oldest first', () => {
    const db = createDatabase(':memory:');
    createUser(db, 'first', 'pw', 'admin');
    createUser(db, 'second', 'pw', 'admin');

    const users = listUsers(db);
    expect(users.map((u) => u.username)).toEqual(['first', 'second']);

    db.close();
  });

  it('disables and re-enables a user', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'bob', 'pw', 'admin');

    setUserDisabled(db, user.id, true);
    expect(listUsers(db).find((u) => u.id === user.id)!.disabledAt).not.toBeNull();

    setUserDisabled(db, user.id, false);
    expect(listUsers(db).find((u) => u.id === user.id)!.disabledAt).toBeNull();

    db.close();
  });

  it('seeds a superadmin only when the users table is empty', () => {
    const db = createDatabase(':memory:');

    seedSuperAdminIfEmpty(db, 'root', 'rootpass');
    expect(listUsers(db)).toHaveLength(1);
    expect(listUsers(db)[0].role).toBe('superadmin');

    seedSuperAdminIfEmpty(db, 'root2', 'otherpass');
    expect(listUsers(db)).toHaveLength(1); // unchanged — table wasn't empty
    expect(listUsers(db)[0].username).toBe('root');

    db.close();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/db/usersRepository.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/db/usersRepository.ts`**

```ts
import bcrypt from 'bcryptjs';
import type Database from 'better-sqlite3';

export type UserRole = 'admin' | 'superadmin';

export interface User {
  id: number;
  username: string;
  role: UserRole;
  createdAt: string;
  disabledAt: string | null;
}

export interface UserWithPasswordHash extends User {
  passwordHash: string;
}

interface UserRow {
  id: number;
  username: string;
  password_hash: string;
  role: UserRole;
  created_at: string;
  disabled_at: string | null;
}

function rowToUser(row: UserRow): User {
  return {
    id: row.id,
    username: row.username,
    role: row.role,
    createdAt: row.created_at,
    disabledAt: row.disabled_at,
  };
}

const SALT_ROUNDS = 10;

export function createUser(
  db: Database.Database,
  username: string,
  password: string,
  role: UserRole
): User {
  const passwordHash = bcrypt.hashSync(password, SALT_ROUNDS);
  const createdAt = new Date().toISOString();
  const result = db
    .prepare('INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, ?, ?)')
    .run(username, passwordHash, role, createdAt);
  return { id: Number(result.lastInsertRowid), username, role, createdAt, disabledAt: null };
}

export function findUserByUsername(
  db: Database.Database,
  username: string
): UserWithPasswordHash | undefined {
  const row = db.prepare('SELECT * FROM users WHERE username = ?').get(username) as
    | UserRow
    | undefined;
  if (!row) return undefined;
  return { ...rowToUser(row), passwordHash: row.password_hash };
}

export function verifyPassword(passwordHash: string, password: string): boolean {
  return bcrypt.compareSync(password, passwordHash);
}

export function listUsers(db: Database.Database): User[] {
  const rows = db.prepare('SELECT * FROM users ORDER BY created_at ASC').all() as UserRow[];
  return rows.map(rowToUser);
}

export function countUsers(db: Database.Database): number {
  const row = db.prepare('SELECT COUNT(*) as count FROM users').get() as { count: number };
  return row.count;
}

export function setUserDisabled(db: Database.Database, userId: number, disabled: boolean): void {
  db.prepare('UPDATE users SET disabled_at = ? WHERE id = ?').run(
    disabled ? new Date().toISOString() : null,
    userId
  );
}

export function seedSuperAdminIfEmpty(db: Database.Database, username: string, password: string): void {
  if (countUsers(db) > 0) return;
  createUser(db, username, password, 'superadmin');
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/db/usersRepository.test.ts`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add src/db/usersRepository.ts tests/db/usersRepository.test.ts
git commit -m "Add UsersRepository: create/find/list/disable users, seed initial superadmin"
```

---

### Task 4: DeclarationsRepository

**Files:**
- Create: `src/db/declarationsRepository.ts`
- Test: `tests/db/declarationsRepository.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/db/declarationsRepository.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';
import { createUser } from '../../src/db/usersRepository.js';
import {
  saveDeclaration,
  listDeclarationsForUser,
  listAllDeclarations,
  getDeclarationById,
} from '../../src/db/declarationsRepository.js';
import type { Declaration } from '../../src/domain/types.js';

function makeDeclaration(): Declaration {
  return {
    code: '500001',
    redevable: 'GLOBAL TRADE LOGISTICS SARL',
    benNumero: '501',
    articles: [
      {
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
      },
    ],
  };
}

describe('declarationsRepository', () => {
  it('saves a declaration with its articles and reads it back', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'admin1', 'pw', 'admin');
    const declaration = makeDeclaration();

    const id = saveDeclaration(db, {
      ownerUserId: user.id,
      declaration,
      shipmentCostFields: {
        devise: 'EUR',
        montantFacture: 2981.34,
        tauxChange: 10.6675,
        fret: 7467.0,
        assurance: 118.0,
        valeurTotaleDeclaree: 40039.992,
      },
      articleCosts: [{ numero: 1, costPerUnit: 42.5 }],
      totalLandedCost: 15045.0,
      costEstimatePartial: false,
      colisCount: 6,
      referenceInterne: 'PO-1234',
      excelFilePath: '/data/declaration-1.xlsx',
    });

    expect(id).toBeGreaterThan(0);

    const saved = getDeclarationById(db, id);
    expect(saved).toMatchObject({
      id,
      ownerUserId: user.id,
      code: '500001',
      redevable: 'GLOBAL TRADE LOGISTICS SARL',
      totalLandedCost: 15045.0,
      costEstimatePartial: false,
      colisCount: 6,
      referenceInterne: 'PO-1234',
      excelFilePath: '/data/declaration-1.xlsx',
    });

    db.close();
  });

  it('throws and saves nothing if an article is missing its computed cost', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'admin1', 'pw', 'admin');
    const declaration = makeDeclaration();

    expect(() =>
      saveDeclaration(db, {
        ownerUserId: user.id,
        declaration,
        shipmentCostFields: {},
        articleCosts: [], // missing cost for article 1
        totalLandedCost: 0,
        costEstimatePartial: true,
        excelFilePath: '/data/declaration-x.xlsx',
      })
    ).toThrow('Missing computed cost for article 1');

    expect(listAllDeclarations(db)).toHaveLength(0); // transaction rolled back

    db.close();
  });

  it('lists declarations scoped to their owner, and lists all declarations for the superadmin view', () => {
    const db = createDatabase(':memory:');
    const alice = createUser(db, 'alice', 'pw', 'admin');
    const bob = createUser(db, 'bob', 'pw', 'admin');

    saveDeclaration(db, {
      ownerUserId: alice.id,
      declaration: makeDeclaration(),
      shipmentCostFields: {},
      articleCosts: [{ numero: 1, costPerUnit: 10 }],
      totalLandedCost: 100,
      costEstimatePartial: true,
      excelFilePath: '/data/a.xlsx',
    });
    saveDeclaration(db, {
      ownerUserId: bob.id,
      declaration: makeDeclaration(),
      shipmentCostFields: {},
      articleCosts: [{ numero: 1, costPerUnit: 20 }],
      totalLandedCost: 200,
      costEstimatePartial: true,
      excelFilePath: '/data/b.xlsx',
    });

    expect(listDeclarationsForUser(db, alice.id)).toHaveLength(1);
    expect(listDeclarationsForUser(db, alice.id)[0].excelFilePath).toBe('/data/a.xlsx');
    expect(listDeclarationsForUser(db, bob.id)).toHaveLength(1);
    expect(listAllDeclarations(db)).toHaveLength(2);

    db.close();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/db/declarationsRepository.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/db/declarationsRepository.ts`**

```ts
import type Database from 'better-sqlite3';
import type { Declaration } from '../domain/types.js';

export interface ShipmentCostFields {
  devise?: string;
  montantFacture?: number;
  tauxChange?: number;
  fret?: number;
  assurance?: number;
  valeurTotaleDeclaree?: number;
}

export interface ArticleCost {
  numero: number;
  costPerUnit: number;
}

export interface SaveDeclarationInput {
  ownerUserId: number;
  declaration: Declaration;
  shipmentCostFields: ShipmentCostFields;
  articleCosts: ArticleCost[];
  totalLandedCost: number;
  costEstimatePartial: boolean;
  colisCount?: number;
  referenceInterne?: string;
  excelFilePath: string;
}

export interface SavedDeclarationSummary {
  id: number;
  ownerUserId: number;
  code: string;
  redevable: string;
  totalLandedCost: number;
  costEstimatePartial: boolean;
  colisCount: number | null;
  referenceInterne: string | null;
  excelFilePath: string;
  createdAt: string;
}

interface DeclarationRow {
  id: number;
  owner_user_id: number;
  code: string;
  redevable: string;
  total_landed_cost: number;
  cost_estimate_partial: number;
  colis_count: number | null;
  reference_interne: string | null;
  excel_file_path: string;
  created_at: string;
}

function rowToSummary(row: DeclarationRow): SavedDeclarationSummary {
  return {
    id: row.id,
    ownerUserId: row.owner_user_id,
    code: row.code,
    redevable: row.redevable,
    totalLandedCost: row.total_landed_cost,
    costEstimatePartial: row.cost_estimate_partial === 1,
    colisCount: row.colis_count,
    referenceInterne: row.reference_interne,
    excelFilePath: row.excel_file_path,
    createdAt: row.created_at,
  };
}

export function saveDeclaration(db: Database.Database, input: SaveDeclarationInput): number {
  const createdAt = new Date().toISOString();
  const insertDeclaration = db.prepare(`
    INSERT INTO declarations (
      owner_user_id, code, redevable, ben_numero, devise, montant_facture, taux_change,
      fret, assurance, valeur_totale_declaree, colis_count, reference_interne,
      total_landed_cost, cost_estimate_partial, excel_file_path, created_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const insertArticle = db.prepare(`
    INSERT INTO declaration_articles (
      declaration_id, numero, hs_code, nom_article, pays, valeur_declaree, quantite,
      total_article, cost_per_unit, taxes_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const run = db.transaction((): number => {
    const result = insertDeclaration.run(
      input.ownerUserId,
      input.declaration.code,
      input.declaration.redevable,
      input.declaration.benNumero,
      input.shipmentCostFields.devise ?? null,
      input.shipmentCostFields.montantFacture ?? null,
      input.shipmentCostFields.tauxChange ?? null,
      input.shipmentCostFields.fret ?? null,
      input.shipmentCostFields.assurance ?? null,
      input.shipmentCostFields.valeurTotaleDeclaree ?? null,
      input.colisCount ?? null,
      input.referenceInterne ?? null,
      input.totalLandedCost,
      input.costEstimatePartial ? 1 : 0,
      input.excelFilePath,
      createdAt
    );
    const declarationId = Number(result.lastInsertRowid);

    for (const article of input.declaration.articles) {
      const cost = input.articleCosts.find((c) => c.numero === article.numero);
      if (!cost) {
        throw new Error(`Missing computed cost for article ${article.numero}`);
      }
      insertArticle.run(
        declarationId,
        article.numero,
        article.hsCode,
        article.nomArticle,
        article.pays,
        article.valeurDeclaree,
        article.quantite,
        article.totalArticle,
        cost.costPerUnit,
        JSON.stringify(article.taxes)
      );
    }

    return declarationId;
  });

  return run();
}

export function listDeclarationsForUser(
  db: Database.Database,
  ownerUserId: number
): SavedDeclarationSummary[] {
  const rows = db
    .prepare('SELECT * FROM declarations WHERE owner_user_id = ? ORDER BY created_at DESC')
    .all(ownerUserId) as DeclarationRow[];
  return rows.map(rowToSummary);
}

export function listAllDeclarations(db: Database.Database): SavedDeclarationSummary[] {
  const rows = db
    .prepare('SELECT * FROM declarations ORDER BY created_at DESC')
    .all() as DeclarationRow[];
  return rows.map(rowToSummary);
}

export function getDeclarationById(
  db: Database.Database,
  id: number
): SavedDeclarationSummary | undefined {
  const row = db.prepare('SELECT * FROM declarations WHERE id = ?').get(id) as
    | DeclarationRow
    | undefined;
  return row ? rowToSummary(row) : undefined;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/db/declarationsRepository.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add src/db/declarationsRepository.ts tests/db/declarationsRepository.test.ts
git commit -m "Add DeclarationsRepository: save declaration + articles transactionally, list/get"
```

---

### Task 5: Extend DumParser for shipment-level cost fields

**Files:**
- Modify: `src/parser/dum/dumParser.ts`
- Modify: `tests/parser/dum/dumParser.test.ts`

The DUM's raw extracted text contains a recognizable cluster —
`"EUR   2 981.340   10.667500   7 467.000  21  118.000   40 039.992 24   06   2026"`
— that positionally encodes: currency code, invoiced amount, exchange rate,
freight, (a form field-number label, ignored), insurance, total declared
value, then a date (`DD MM YYYY`, used only as an anchor to terminate the
match reliably). This has been verified against the real sample fixture.

- [ ] **Step 1: Write the failing test**

Add to `tests/parser/dum/dumParser.test.ts` (append inside the existing `describe('parseDum', ...)` block, before its closing `});`):

```ts
  it('extracts shipment-level cost fields (devise, montant facturé, taux de change, fret, assurance, valeur totale déclarée) from the real sample', () => {
    const text = readFileSync(fixturePath, 'utf-8');
    const result = parseDum(text);

    expect(result.shipmentCost).toEqual({
      devise: 'EUR',
      montantFacture: 2981.34,
      tauxChange: 10.6675,
      fret: 7467.0,
      assurance: 118.0,
      valeurTotaleDeclaree: 40039.992,
    });
  });

  it('leaves shipmentCost undefined (not a hard failure) when the cluster is not found', () => {
    const text = `Crédit d'enlèvement 700002123
6109100010   1 000.000 5.00   AP 10.0 U MAROC   MA  COLIS  CHEMISE 10.00 NB 1`;
    const result = parseDum(text);

    expect(result.shipmentCost).toBeUndefined();
  });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/parser/dum/dumParser.test.ts`
Expected: FAIL — `result.shipmentCost` is `undefined` in the first new test (property doesn't exist yet).

- [ ] **Step 3: Modify `src/parser/dum/dumParser.ts`**

Add a new exported interface and extraction logic. First, add this import to the top of the file (alongside the existing import):

```ts
import { extractFirst, parseFrenchNumber } from '../shared/text.js';
```

(this import already exists — no change needed there). Now add the new interface right after `DumArticleResult`:

```ts
export interface DumShipmentCost {
  devise: string;
  montantFacture: number;
  tauxChange: number;
  fret: number;
  assurance: number;
  valeurTotaleDeclaree: number;
}
```

Add `shipmentCost?: DumShipmentCost;` to the `DumResult` interface:

```ts
export interface DumResult {
  creditEnlevementCode: string;
  articles: DumArticleResult[];
  shipmentCost?: DumShipmentCost;
}
```

Add the new pattern constant near `ARTICLE_PATTERN`:

```ts
// See design spec §4 — this cluster of shipment-level values (currency,
// invoiced amount, exchange rate, freight, [a form field-number label,
// ignored], insurance, total declared value) appears together in the raw
// extracted text, terminated by a DD MM YYYY date used only as a reliable
// anchor. Optional: if not found, shipmentCost is simply omitted rather
// than treated as a parse failure — the app's core function doesn't depend
// on it.
const SHIPMENT_COST_PATTERN =
  /\b(EUR|MAD|USD|GBP)\s+(\d[\d\s.,]*?\d)\s+([\d.]+)\s+(\d[\d\s.,]*?\d)\s+\d+\s+([\d.]+)\s+(\d[\d\s.,]*?\d)\s+\d{2}\s+\d{2}\s+\d{4}/;

function extractShipmentCost(text: string): DumShipmentCost | undefined {
  const match = text.match(SHIPMENT_COST_PATTERN);
  if (!match) return undefined;

  const [, devise, montantFactureRaw, tauxChangeRaw, fretRaw, assuranceRaw, valeurTotaleRaw] = match;
  return {
    devise,
    montantFacture: parseFrenchNumber(montantFactureRaw),
    tauxChange: parseFrenchNumber(tauxChangeRaw),
    fret: parseFrenchNumber(fretRaw),
    assurance: parseFrenchNumber(assuranceRaw),
    valeurTotaleDeclaree: parseFrenchNumber(valeurTotaleRaw),
  };
}
```

Finally, modify the `return` statement at the end of `parseDum` to include it:

```ts
  return { creditEnlevementCode, articles, shipmentCost: extractShipmentCost(text) };
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/parser/dum/dumParser.test.ts`
Expected: PASS (7 tests)

- [ ] **Step 5: Run the full suite to confirm no regression**

Run: `npm test`
Expected: all tests pass (this change is additive — every existing `DumResult` consumer ignores the new optional field).

- [ ] **Step 6: Commit**

```bash
git add src/parser/dum/dumParser.ts tests/parser/dum/dumParser.test.ts
git commit -m "Extend DumParser to extract shipment-level cost fields (devise, montant facturé, taux de change, fret, assurance, valeur totale déclarée)"
```

---

### Task 6: Cost calculator

**Files:**
- Create: `src/domain/costCalculator.ts`
- Test: `tests/domain/costCalculator.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/domain/costCalculator.test.ts`:

```ts
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { calculateLandedCost } from '../../src/domain/costCalculator.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dumText = readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8');
  const dum = parseDum(dumText);
  return { declaration: mergeDeclaration(liquidation, dum), shipmentCost: dum.shipmentCost! };
}

describe('calculateLandedCost', () => {
  it('computes a full (non-partial) landed cost whose total equals shipment-cost + total taxes, regardless of per-article split', () => {
    const { declaration, shipmentCost } = loadRealDeclaration();

    const result = calculateLandedCost(declaration, shipmentCost);

    expect(result.partial).toBe(false);
    expect(result.articleCosts).toHaveLength(2);

    const totalTaxes = declaration.articles.reduce(
      (sum, a) => sum + a.taxes.reduce((s, t) => s + t.montant, 0),
      0
    );
    const expectedTotal =
      (shipmentCost.montantFacture + shipmentCost.fret + shipmentCost.assurance) *
        shipmentCost.tauxChange +
      totalTaxes;

    expect(result.totalLandedCost).toBeCloseTo(expectedTotal, 2);

    // Reconciliation: summing (costPerUnit * quantite) per article must equal the total.
    const reconciledTotal = result.articleCosts.reduce((sum, c) => {
      const article = declaration.articles.find((a) => a.numero === c.numero)!;
      return sum + c.costPerUnit * article.quantite;
    }, 0);
    expect(reconciledTotal).toBeCloseTo(expectedTotal, 2);

    // Every article's cost per unit must be positive and finite.
    for (const cost of result.articleCosts) {
      expect(cost.costPerUnit).toBeGreaterThan(0);
      expect(Number.isFinite(cost.costPerUnit)).toBe(true);
    }
  });

  it('falls back to duty-only-per-unit and marks the result partial when shipment cost fields are unavailable', () => {
    const { declaration } = loadRealDeclaration();

    const result = calculateLandedCost(declaration, {});

    expect(result.partial).toBe(true);

    const article1 = declaration.articles.find((a) => a.numero === 1)!;
    const article1TaxTotal = article1.taxes.reduce((sum, t) => sum + t.montant, 0);
    const cost1 = result.articleCosts.find((c) => c.numero === 1)!;
    expect(cost1.costPerUnit).toBeCloseTo(article1TaxTotal / article1.quantite, 4);
  });

  it('treats a partially-populated shipment cost (e.g. missing tauxChange) as partial, not a crash', () => {
    const { declaration, shipmentCost } = loadRealDeclaration();
    const { tauxChange, ...incomplete } = shipmentCost;

    const result = calculateLandedCost(declaration, incomplete);

    expect(result.partial).toBe(true);
    for (const cost of result.articleCosts) {
      expect(Number.isFinite(cost.costPerUnit)).toBe(true);
    }
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/domain/costCalculator.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Create `src/domain/costCalculator.ts`**

```ts
import type { Declaration } from './types.js';

export interface PartialShipmentCost {
  devise?: string;
  montantFacture?: number;
  tauxChange?: number;
  fret?: number;
  assurance?: number;
  valeurTotaleDeclaree?: number;
}

export interface ArticleCostResult {
  numero: number;
  costPerUnit: number;
}

export interface CostCalculationResult {
  articleCosts: ArticleCostResult[];
  totalLandedCost: number;
  partial: boolean;
}

// Per design spec §5: each article's share of shipment-level costs
// (invoiced amount + freight + insurance, converted via the DUM's own
// exchange rate) is allocated proportionally to its share of the
// shipment's total declared value — the same principle already used for
// tax-per-unit splitting, applied one level up. If any of the four
// shipment-level fields is missing, shipment cost is treated as zero and
// the result is flagged `partial` so callers don't present an incomplete
// number as a complete one.
export function calculateLandedCost(
  declaration: Declaration,
  shipment: PartialShipmentCost
): CostCalculationResult {
  const { montantFacture, fret, assurance, tauxChange } = shipment;
  const hasFullShipmentData =
    montantFacture !== undefined &&
    fret !== undefined &&
    assurance !== undefined &&
    tauxChange !== undefined;

  const totalDeclaredValue = declaration.articles.reduce((sum, a) => sum + a.valeurDeclaree, 0);

  const articleCosts: ArticleCostResult[] = declaration.articles.map((article) => {
    const taxTotal = article.taxes.reduce((sum, tax) => sum + tax.montant, 0);

    let shipmentCostForArticle = 0;
    if (hasFullShipmentData && totalDeclaredValue > 0) {
      const share = article.valeurDeclaree / totalDeclaredValue;
      shipmentCostForArticle = (montantFacture + fret + assurance) * tauxChange * share;
    }

    const totalArticleCost = shipmentCostForArticle + taxTotal;
    const costPerUnit = article.quantite > 0 ? totalArticleCost / article.quantite : 0;

    return { numero: article.numero, costPerUnit };
  });

  const totalLandedCost = articleCosts.reduce((sum, cost) => {
    const article = declaration.articles.find((a) => a.numero === cost.numero)!;
    return sum + cost.costPerUnit * article.quantite;
  }, 0);

  return { articleCosts, totalLandedCost, partial: !hasFullShipmentData };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/domain/costCalculator.test.ts`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add src/domain/costCalculator.ts tests/domain/costCalculator.test.ts
git commit -m "Add landed-cost-per-unit calculator (purchase + freight + insurance + duties)"
```

---

### Task 7: Rewrite auth.ts for DB-backed, role-aware sessions

**Files:**
- Modify: `src/web/auth.ts`
- Modify: `tests/web/auth.test.ts`

- [ ] **Step 1: Write the failing test (full rewrite of the test file)**

Replace the entire contents of `tests/web/auth.test.ts` with:

```ts
import { describe, it, expect } from 'vitest';
import { createSession, requireAuth, requireSuperAdmin, type SessionInfo } from '../../src/web/auth.js';
import type { NextFunction, Request, Response } from 'express';

function makeReqRes(cookieHeader: string | undefined, method = 'GET') {
  const req = { headers: { cookie: cookieHeader }, method } as unknown as Request;
  const redirectCalls: string[] = [];
  const jsonCalls: unknown[] = [];
  const sendCalls: unknown[] = [];
  let statusCode: number | undefined;
  const res = {
    redirect: (url: string) => redirectCalls.push(url),
    status: (code: number) => {
      statusCode = code;
      return res;
    },
    json: (body: unknown) => jsonCalls.push(body),
    send: (body: unknown) => sendCalls.push(body),
  } as unknown as Response;
  let nextCalled = false;
  const next: NextFunction = () => {
    nextCalled = true;
  };
  return {
    req,
    res,
    next,
    redirectCalls,
    jsonCalls,
    sendCalls,
    wasNextCalled: () => nextCalled,
    getStatusCode: () => statusCode,
  };
}

describe('requireAuth', () => {
  it('calls next() and attaches the session for a valid session cookie', () => {
    const sessionId = createSession({ userId: 1, username: 'alice', role: 'admin' });
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes(`session=${sessionId}`);

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(true);
    expect(redirectCalls).toHaveLength(0);
    expect((req as Request & { session?: SessionInfo }).session).toEqual({
      userId: 1,
      username: 'alice',
      role: 'admin',
    });
  });

  it('redirects to /login for a GET request with no session cookie', () => {
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes(undefined);

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(redirectCalls).toEqual(['/login']);
  });

  it('redirects to /login for a GET request with an unknown/expired session id', () => {
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes('session=not-a-real-session');

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(redirectCalls).toEqual(['/login']);
  });

  it('responds with JSON 401 instead of redirecting for an unauthenticated POST request', () => {
    const { req, res, next, wasNextCalled, redirectCalls, jsonCalls, getStatusCode } = makeReqRes(
      undefined,
      'POST'
    );

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(redirectCalls).toHaveLength(0);
    expect(getStatusCode()).toBe(401);
    expect(jsonCalls).toEqual([{ success: false, error: 'Session expirée, veuillez vous reconnecter.' }]);
  });
});

describe('requireSuperAdmin', () => {
  it('calls next() when the session role is superadmin', () => {
    const sessionId = createSession({ userId: 2, username: 'root', role: 'superadmin' });
    const { req, res, next, wasNextCalled, sendCalls } = makeReqRes(`session=${sessionId}`);

    requireSuperAdmin(req, res, next);

    expect(wasNextCalled()).toBe(true);
    expect(sendCalls).toHaveLength(0);
  });

  it('responds 403 when the session role is admin (not superadmin)', () => {
    const sessionId = createSession({ userId: 3, username: 'alice', role: 'admin' });
    const { req, res, next, wasNextCalled, getStatusCode, sendCalls } = makeReqRes(`session=${sessionId}`);

    requireSuperAdmin(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(getStatusCode()).toBe(403);
    expect(sendCalls).toHaveLength(1);
  });

  it('responds 403 when there is no session at all', () => {
    const { req, res, next, wasNextCalled, getStatusCode } = makeReqRes(undefined);

    requireSuperAdmin(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(getStatusCode()).toBe(403);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/web/auth.test.ts`
Expected: FAIL — `createSession` currently takes no arguments and `requireSuperAdmin` doesn't exist.

- [ ] **Step 3: Replace the entire contents of `src/web/auth.ts`**

```ts
import { randomUUID } from 'node:crypto';
import type { NextFunction, Request, Response } from 'express';
import type { UserRole } from '../db/usersRepository.js';

// Per design spec §2: sessions are still held in memory (not persisted to
// SQLite) — restarting the server logs everyone out, same tradeoff as
// before. What's new is that sessions now carry a role, and the
// credentials they were created from live in the `users` table (bcrypt
// password hashes) instead of a single hardcoded pair.
const SESSION_COOKIE_NAME = 'session';

export interface SessionInfo {
  userId: number;
  username: string;
  role: UserRole;
}

const activeSessions = new Map<string, SessionInfo>();

export function createSession(user: SessionInfo): string {
  const sessionId = randomUUID();
  activeSessions.set(sessionId, user);
  return sessionId;
}

function getSessionIdFromCookie(cookieHeader: string | undefined): string | undefined {
  if (!cookieHeader) return undefined;
  const match = cookieHeader
    .split(';')
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${SESSION_COOKIE_NAME}=`));
  return match?.slice(SESSION_COOKIE_NAME.length + 1);
}

export function setSessionCookie(res: Response, sessionId: string): void {
  res.setHeader('Set-Cookie', `${SESSION_COOKIE_NAME}=${sessionId}; HttpOnly; Path=/; SameSite=Lax`);
}

function getSession(req: Request): SessionInfo | undefined {
  const sessionId = getSessionIdFromCookie(req.headers.cookie);
  if (!sessionId) return undefined;
  return activeSessions.get(sessionId);
}

export function requireAuth(req: Request, res: Response, next: NextFunction): void {
  const session = getSession(req);
  if (session) {
    (req as Request & { session?: SessionInfo }).session = session;
    next();
    return;
  }
  // The client's fetch()-based upload flow expects JSON back from POST
  // /generate; redirecting it to the (HTML) /login page here made the
  // client's response.json() throw a confusing "Unexpected token '<'"
  // instead of the real problem. Only redirect for normal page navigation
  // (GET); respond with JSON for API-style POST requests.
  if (req.method === 'POST') {
    res.status(401).json({ success: false, error: 'Session expirée, veuillez vous reconnecter.' });
    return;
  }
  res.redirect('/login');
}

export function requireSuperAdmin(req: Request, res: Response, next: NextFunction): void {
  const existing = (req as Request & { session?: SessionInfo }).session;
  const session = existing ?? getSession(req);
  if (session?.role === 'superadmin') {
    (req as Request & { session?: SessionInfo }).session = session;
    next();
    return;
  }
  res.status(403).send('Accès refusé — réservé au superadmin.');
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/web/auth.test.ts`
Expected: PASS (7 tests)

- [ ] **Step 5: Commit**

```bash
git add src/web/auth.ts tests/web/auth.test.ts
git commit -m "Rewrite auth.ts for DB-backed, role-aware sessions (admin/superadmin)"
```

---

### Task 8: Wire the new auth into server.ts's login flow, seed superadmin at boot

**Files:**
- Modify: `src/web/server.ts`

This is the task that makes the app runnable end-to-end again — `server.ts`
currently calls the old `checkCredentials(username, password)` /
`createSession()` (no-arg) API that Task 7 replaced. The
upload/generate/download flow itself is untouched; only the login handler
and startup sequence change. Declaration history saving and the
admin/superadmin dashboards are NOT part of this task — they're a follow-up
plan that builds on the repositories and calculator from Tasks 3, 4, and 6.

- [ ] **Step 1: Update the imports and add DB/user imports**

In `src/web/server.ts`, find this line:

```ts
import { checkCredentials, createSession, requireAuth, setSessionCookie } from './auth.js';
```

Replace it with:

```ts
import { createSession, requireAuth, setSessionCookie } from './auth.js';
import { getDatabase } from '../db/database.js';
import { findUserByUsername, verifyPassword, seedSuperAdminIfEmpty } from '../db/usersRepository.js';
```

- [ ] **Step 2: Seed the superadmin at boot**

Find this line (near the top of the file, after the `UPLOAD_DIR`/`OUTPUT_DIR` setup):

```ts
mkdirSync(UPLOAD_DIR, { recursive: true });
mkdirSync(OUTPUT_DIR, { recursive: true });
```

Add immediately after it:

```ts
const db = getDatabase();
const superAdminUsername = process.env.SUPERADMIN_USERNAME ?? 'redwan';
const superAdminPassword = process.env.SUPERADMIN_PASSWORD ?? 'redwan2026';
if (!process.env.SUPERADMIN_USERNAME || !process.env.SUPERADMIN_PASSWORD) {
  console.warn(
    'SUPERADMIN_USERNAME/SUPERADMIN_PASSWORD not set — falling back to default credentials for initial setup. Set these in production.'
  );
}
seedSuperAdminIfEmpty(db, superAdminUsername, superAdminPassword);
```

- [ ] **Step 3: Replace the login POST handler**

Find:

```ts
app.post('/login', (req, res) => {
  const { username, password } = req.body as { username?: string; password?: string };

  if (!username || !password || !checkCredentials(username, password)) {
    const errorBlock =
      '<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>Identifiant ou mot de passe incorrect.</span></div>';
    res.status(401).send(loginHtml.replace('{{ERROR_BLOCK}}', errorBlock));
    return;
  }

  const sessionId = createSession();
  setSessionCookie(res, sessionId);
  res.redirect('/');
});
```

Replace it with:

```ts
app.post('/login', (req, res) => {
  const { username, password } = req.body as { username?: string; password?: string };
  const errorBlock =
    '<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>Identifiant ou mot de passe incorrect.</span></div>';

  if (!username || !password) {
    res.status(401).send(loginHtml.replace('{{ERROR_BLOCK}}', errorBlock));
    return;
  }

  const user = findUserByUsername(db, username);
  if (!user || user.disabledAt || !verifyPassword(user.passwordHash, password)) {
    res.status(401).send(loginHtml.replace('{{ERROR_BLOCK}}', errorBlock));
    return;
  }

  const sessionId = createSession({ userId: user.id, username: user.username, role: user.role });
  setSessionCookie(res, sessionId);
  res.redirect('/');
});
```

- [ ] **Step 4: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors.

- [ ] **Step 5: Run the full test suite**

Run: `npm test`
Expected: all tests pass (server.ts itself has no direct unit tests, but this
confirms nothing else broke).

- [ ] **Step 6: Manual smoke test**

Run: `rm -f data.db && npm run web`
Then in another terminal:
```bash
curl -s -c /tmp/cookies.txt -X POST http://localhost:4310/login \
  -d "username=redwan&password=redwan2026" -o /dev/null -w "login:%{http_code}\n"
curl -s -b /tmp/cookies.txt http://localhost:4310/ -o /dev/null -w "upload-page:%{http_code}\n"
```
Expected: `login:302` (redirect to `/`) then `upload-page:200`. Stop the server
afterward (`Ctrl+C` in its terminal) and delete the smoke-test database:
`rm -f data.db data.db-wal data.db-shm`.

- [ ] **Step 7: Commit**

```bash
git add src/web/server.ts
git commit -m "Wire DB-backed auth into server.ts login flow; seed superadmin at boot"
```

---

## Phase exit criteria

- `npm test` passes with all tests green across all files, including the new
  `tests/db/*`, `tests/domain/costCalculator.test.ts`, and the rewritten
  `tests/web/auth.test.ts`.
- `npx tsc --noEmit` passes with no errors.
- The existing login → upload → generate → download flow works end-to-end
  against the new DB-backed auth (verified in Task 8's manual smoke test).
- `UsersRepository`, `DeclarationsRepository`, and `calculateLandedCost` are
  fully tested and ready for the next plan (admin/superadmin dashboards) to
  wire into HTTP routes and UI — no changes to their public APIs should be
  needed for that follow-up plan to consume them.
- Declaration history is **not yet** saved anywhere by the running app (no
  route calls `saveDeclaration` yet) — that wiring, plus the colis/référence
  entry form and both dashboards, is explicitly the next plan's job, per
  spec §3.6 (out of scope note) is not applicable here; this is simply
  sequencing, not a scope cut.

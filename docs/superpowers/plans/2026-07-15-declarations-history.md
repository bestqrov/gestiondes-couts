# Historique des déclarations Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the superadmin's "Coût de produit" single-transaction view with a paginated, filterable history of every saved declaration, plus a per-declaration detail page.

**Architecture:** A new `listTransactionsPaginated` query in `transactionsRepository.ts` (sort by `createdAt` desc, optional redevable + date-range filters, skip/limit pagination) backs a rewritten `GET /superadmin/costs` list route and a new `GET /superadmin/costs/:id` detail route. The list/detail rendering lives in `renderSuperAdminDashboard.ts`, reusing the existing totals-card markup (`renderSearchResultCard`) for the detail view. See `docs/superpowers/specs/2026-07-15-declarations-history-design.md` for the approved design.

**Tech Stack:** TypeScript, Express 5, MongoDB driver, Vitest.

---

### Task 1: Extend the fake Mongo collection test helper (skip + query operators)

**Files:**
- Modify: `tests/helpers/fakeMongoCollection.ts`

The new pagination query needs `.find(filter).sort().skip().limit().toArray()`, and the filter will use `$regex`/`$options` (redevable search, already used by `searchTransactionsByRedevable`) and `$gte`/`$lte` (date range) — the fake's `matches()` only supports exact equality today.

- [ ] **Step 1: Add operator support to `matches()`**

Replace the `matches` function:

```ts
  function valueMatches(actual: unknown, expected: unknown): boolean {
    if (expected !== null && typeof expected === 'object' && !(expected instanceof ObjectId)) {
      const ops = expected as Record<string, unknown>;
      if ('$regex' in ops) {
        const flags = typeof ops.$options === 'string' ? ops.$options : '';
        return new RegExp(String(ops.$regex), flags).test(String(actual));
      }
      let ok = true;
      if ('$gte' in ops) ok = ok && (actual as string) >= (ops.$gte as string);
      if ('$lte' in ops) ok = ok && (actual as string) <= (ops.$lte as string);
      return ok;
    }
    return actual === expected;
  }

  function matches(doc: T, filter: Record<string, unknown>): boolean {
    return Object.entries(filter).every(([key, value]) => {
      if (key === '_id') return idsEqual(doc._id, value);
      return valueMatches((doc as Record<string, unknown>)[key], value);
    });
  }
```

- [ ] **Step 2: Add `.skip()` to the cursor returned by `find()`**

In the `find` method, add a `skip` step between `sort` and `toArray`:

```ts
    find: (filter: Record<string, unknown> = {}) => {
      let results = docs.filter((d) => matches(d, filter));
      let skipCount = 0;
      let limitCount: number | undefined;
      const cursor = {
        sort(sortSpec: Record<string, number>) {
          const [field, dir] = Object.entries(sortSpec)[0];
          results = [...results].sort((a, b) => {
            const av = (a as Record<string, unknown>)[field] as string | number;
            const bv = (b as Record<string, unknown>)[field] as string | number;
            if (av < bv) return -dir;
            if (av > bv) return dir;
            return 0;
          });
          return cursor;
        },
        skip(n: number) {
          skipCount = n;
          return cursor;
        },
        limit(n: number) {
          limitCount = n;
          return cursor;
        },
        toArray: async () => {
          const sliced = results.slice(skipCount, limitCount !== undefined ? skipCount + limitCount : undefined);
          return sliced.map((d) => ({ ...d }));
        },
      };
      return cursor;
    },
```

Note: `getMostRecentTransaction` already calls `.find().sort().limit(1).toArray()` — check whether `limit` already existed before this change (it didn't; that function must currently be untested against the fake, only against real MongoDB per the project's existing manual-smoke-test pattern). Adding `limit` here is additive and doesn't change any other behavior.

- [ ] **Step 3: Run the existing test suite to confirm nothing regressed**

Run: `npx vitest run`
Expected: all currently-passing tests still pass (this step only adds capabilities to the fake, doesn't change its existing behavior for exact-match filters).

- [ ] **Step 4: Commit**

```bash
git add tests/helpers/fakeMongoCollection.ts
git commit -m "Add skip/limit and query operators ($regex, $gte, $lte) to fake Mongo collection"
```

---

### Task 2: Add `listTransactionsPaginated` and `getTransactionById` to transactionsRepository.ts

**Files:**
- Modify: `src/db/transactionsRepository.ts`
- Test: `tests/db/transactionsRepository.listTransactionsPaginated.test.ts` (new)

- [ ] **Step 1: Add an optional `_id` field to `TransactionDocument`**

In `src/db/transactionsRepository.ts`, add the import and field:

```ts
import { ObjectId, type Collection } from 'mongodb';
```

(replacing the existing `import type { Collection } from 'mongodb';`)

```ts
export interface TransactionDocument {
  _id?: ObjectId;
  ownerUserId: string;
  code: string;
  redevable: string;
  valeurTotaleDeclaree: number | null;
  totalTaxes: number;
  totalLandedCost: number;
  costEstimatePartial: boolean;
  colisCount: number | null;
  referenceInterne: string | null;
  articles: TransactionArticle[];
  createdAt: string;
}
```

- [ ] **Step 2: Write the failing test for `listTransactionsPaginated`**

Create `tests/db/transactionsRepository.listTransactionsPaginated.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { createFakeCollection } from '../helpers/fakeMongoCollection.js';
import {
  saveTransaction,
  listTransactionsPaginated,
  type TransactionDocument,
} from '../../src/db/transactionsRepository.js';

function makeCollection() {
  return createFakeCollection<TransactionDocument>();
}

async function seed(
  collection: ReturnType<typeof makeCollection>,
  entries: Array<{ redevable: string; createdAt: string }>
) {
  for (const entry of entries) {
    await saveTransaction(collection, {
      ownerUserId: 'u1',
      code: '500001',
      redevable: entry.redevable,
      valeurTotaleDeclaree: null,
      totalTaxes: 0,
      totalLandedCost: 0,
      costEstimatePartial: false,
      articles: [],
    });
    // saveTransaction stamps createdAt with `new Date().toISOString()`, so
    // overwrite it directly on the fake's stored doc via a second save
    // isn't possible — instead, seed with pre-set createdAt by inserting
    // through the collection directly for full control in this test.
  }
}

describe('listTransactionsPaginated', () => {
  it('sorts newest first and paginates with the given page size', async () => {
    const collection = makeCollection();
    for (let i = 0; i < 5; i++) {
      await collection.insertOne({
        ownerUserId: 'u1',
        code: `50000${i}`,
        redevable: `SOCIETE ${i}`,
        valeurTotaleDeclaree: null,
        totalTaxes: 0,
        totalLandedCost: 0,
        costEstimatePartial: false,
        colisCount: null,
        referenceInterne: null,
        articles: [],
        createdAt: `2026-01-0${i + 1}T00:00:00.000Z`,
      });
    }

    const page1 = await listTransactionsPaginated(collection, { page: 1, pageSize: 2 });
    expect(page1.items.map((t) => t.code)).toEqual(['500004', '500003']);
    expect(page1.total).toBe(5);

    const page2 = await listTransactionsPaginated(collection, { page: 2, pageSize: 2 });
    expect(page2.items.map((t) => t.code)).toEqual(['500002', '500001']);

    const page3 = await listTransactionsPaginated(collection, { page: 3, pageSize: 2 });
    expect(page3.items.map((t) => t.code)).toEqual(['500000']);
  });

  it('filters by redevable (case-insensitive substring)', async () => {
    const collection = makeCollection();
    await collection.insertOne({
      ownerUserId: 'u1', code: 'A', redevable: 'Global Trade SARL', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-01-01T00:00:00.000Z',
    });
    await collection.insertOne({
      ownerUserId: 'u1', code: 'B', redevable: 'Other Company', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-01-02T00:00:00.000Z',
    });

    const result = await listTransactionsPaginated(collection, {
      page: 1,
      pageSize: 20,
      redevable: 'global',
    });
    expect(result.items.map((t) => t.code)).toEqual(['A']);
    expect(result.total).toBe(1);
  });

  it('filters by date range (dateFrom/dateTo, inclusive)', async () => {
    const collection = makeCollection();
    await collection.insertOne({
      ownerUserId: 'u1', code: 'A', redevable: 'X', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-01-01T00:00:00.000Z',
    });
    await collection.insertOne({
      ownerUserId: 'u1', code: 'B', redevable: 'X', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-01-15T00:00:00.000Z',
    });
    await collection.insertOne({
      ownerUserId: 'u1', code: 'C', redevable: 'X', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-02-01T00:00:00.000Z',
    });

    const result = await listTransactionsPaginated(collection, {
      page: 1,
      pageSize: 20,
      dateFrom: '2026-01-10',
      dateTo: '2026-01-31',
    });
    expect(result.items.map((t) => t.code)).toEqual(['B']);
  });

  it('returns an empty list with total 0 when nothing matches', async () => {
    const collection = makeCollection();
    const result = await listTransactionsPaginated(collection, { page: 1, pageSize: 20 });
    expect(result.items).toEqual([]);
    expect(result.total).toBe(0);
  });
});
```

- [ ] **Step 3: Run the test to verify it fails**

Run: `npx vitest run tests/db/transactionsRepository.listTransactionsPaginated.test.ts`
Expected: FAIL — `listTransactionsPaginated` is not exported yet.

- [ ] **Step 4: Implement `listTransactionsPaginated` and `getTransactionById`**

In `src/db/transactionsRepository.ts`, add after `getMostRecentTransaction`:

```ts
export interface ListTransactionsOptions {
  page: number; // 1-indexed
  pageSize: number;
  redevable?: string;
  dateFrom?: string; // ISO date (yyyy-mm-dd), inclusive
  dateTo?: string; // ISO date (yyyy-mm-dd), inclusive
}

export interface ListTransactionsResult {
  items: TransactionDocument[];
  total: number;
  page: number;
  pageSize: number;
}

// Powers the superadmin "Historique" page — every saved transaction across
// all admins, newest first, with an optional redevable substring filter and
// an optional createdAt date range, paginated. dateTo is widened to the end
// of that calendar day (23:59:59.999) since createdAt is a precise
// timestamp but the filter inputs are plain dates.
export async function listTransactionsPaginated(
  collection: Collection<TransactionDocument>,
  options: ListTransactionsOptions
): Promise<ListTransactionsResult> {
  const filter: Record<string, unknown> = {};
  if (options.redevable) {
    const escaped = options.redevable.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    filter.redevable = { $regex: escaped, $options: 'i' };
  }
  if (options.dateFrom || options.dateTo) {
    const range: Record<string, string> = {};
    if (options.dateFrom) range.$gte = `${options.dateFrom}T00:00:00.000Z`;
    if (options.dateTo) range.$lte = `${options.dateTo}T23:59:59.999Z`;
    filter.createdAt = range;
  }

  const [items, total] = await Promise.all([
    collection
      .find(filter)
      .sort({ createdAt: -1 })
      .skip((options.page - 1) * options.pageSize)
      .limit(options.pageSize)
      .toArray(),
    collection.countDocuments(filter),
  ]);

  return { items, total, page: options.page, pageSize: options.pageSize };
}

export async function getTransactionById(
  collection: Collection<TransactionDocument>,
  id: string
): Promise<TransactionDocument | null> {
  return collection.findOne({ _id: new ObjectId(id) });
}
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `npx vitest run tests/db/transactionsRepository.listTransactionsPaginated.test.ts`
Expected: PASS (4 tests)

- [ ] **Step 6: Run the full suite and typecheck**

Run: `npx tsc --noEmit && npx vitest run`
Expected: no type errors, all tests pass (the existing `aggregateCountryProductCounts` tests and any other transactionsRepository consumers are unaffected — `_id` is optional and additive).

- [ ] **Step 7: Commit**

```bash
git add src/db/transactionsRepository.ts tests/db/transactionsRepository.listTransactionsPaginated.test.ts
git commit -m "Add paginated, filterable transaction listing for the declarations history page"
```

---

### Task 3: Rewrite the `/superadmin/costs` routes in server.ts

**Files:**
- Modify: `src/web/server.ts`

- [ ] **Step 1: Update the import from `transactionsRepository.js`**

Find (near the top of `src/web/server.ts`):

```ts
import {
  saveTransaction,
  getMostRecentTransaction,
  countTransactions,
  getCountryProductCounts,
  searchTransactionsByRedevable,
  TRANSACTIONS_COLLECTION,
  type TransactionDocument,
  type CountryProductCount,
} from '../db/transactionsRepository.js';
```

Replace with:

```ts
import {
  saveTransaction,
  countTransactions,
  getCountryProductCounts,
  listTransactionsPaginated,
  getTransactionById,
  TRANSACTIONS_COLLECTION,
  type TransactionDocument,
  type CountryProductCount,
} from '../db/transactionsRepository.js';
```

(`getMostRecentTransaction` and `searchTransactionsByRedevable` are no longer called directly by server.ts once this task is done — leave them exported from the repository file itself since removing them isn't necessary and isn't part of this task's scope.)

- [ ] **Step 2: Update the `renderSuperAdminCosts`/`renderSuperAdminPlaceholder` import**

Find:

```ts
import {
  renderSuperAdminOverview,
  renderSuperAdminUsers,
  renderSuperAdminPlaceholder,
  renderSuperAdminCosts,
  renderSuperAdminSettings,
  renderSuperAdminGenerate,
} from './renderSuperAdminDashboard.js';
```

Replace with:

```ts
import {
  renderSuperAdminOverview,
  renderSuperAdminUsers,
  renderSuperAdminPlaceholder,
  renderSuperAdminCosts,
  renderSuperAdminCostDetail,
  renderSuperAdminSettings,
  renderSuperAdminGenerate,
} from './renderSuperAdminDashboard.js';
```

- [ ] **Step 3: Replace the `GET /superadmin/costs` handler**

Find the existing handler (currently reads `mostRecent`/`searchQuery`/`searchResults`) and replace its body with pagination + date-range parsing:

```ts
app.get('/superadmin/costs', requireSuperAdmin, async (req, res) => {
  let collection: Collection<TransactionDocument>;
  try {
    const mongoDb = await getMongoDb();
    collection = mongoDb.collection<TransactionDocument>(TRANSACTIONS_COLLECTION);
  } catch (mongoError) {
    console.error('Failed to reach MongoDB for Historique:', mongoError);
    res.status(503).send(
      renderSuperAdminPlaceholder(
        'Historique',
        "Impossible de se connecter à la base de données pour le moment. Réessayez plus tard.",
        DEFAULT_APP_SETTINGS
      )
    );
    return;
  }

  const total = await countTransactions(collection);
  const settings = await getAppSettings(await getSettingsCollection());
  if (total === 0) {
    res.send(
      renderSuperAdminPlaceholder(
        'Historique',
        "Aucune déclaration n'a encore été générée sur l'application. L'historique s'affichera ici après une génération.",
        settings
      )
    );
    return;
  }

  const page = Math.max(1, Number(req.query.page) || 1);
  const pageSize = 20;
  const redevable = typeof req.query.q === 'string' ? req.query.q.trim() : '';
  const dateFrom = typeof req.query.dateFrom === 'string' ? req.query.dateFrom : '';
  const dateTo = typeof req.query.dateTo === 'string' ? req.query.dateTo : '';

  const result = await listTransactionsPaginated(collection, {
    page,
    pageSize,
    redevable: redevable || undefined,
    dateFrom: dateFrom || undefined,
    dateTo: dateTo || undefined,
  });

  res.send(renderSuperAdminCosts(result, settings, { q: redevable, dateFrom, dateTo }));
});

app.get('/superadmin/costs/:id', requireSuperAdmin, async (req, res) => {
  let collection: Collection<TransactionDocument>;
  try {
    const mongoDb = await getMongoDb();
    collection = mongoDb.collection<TransactionDocument>(TRANSACTIONS_COLLECTION);
  } catch (mongoError) {
    console.error('Failed to reach MongoDB for a declaration detail:', mongoError);
    res.status(503).send(
      renderSuperAdminPlaceholder(
        'Historique',
        "Impossible de se connecter à la base de données pour le moment. Réessayez plus tard.",
        DEFAULT_APP_SETTINGS
      )
    );
    return;
  }

  const transaction = await getTransactionById(collection, req.params.id);
  const settings = await getAppSettings(await getSettingsCollection());
  if (!transaction) {
    res.status(404).send(
      renderSuperAdminPlaceholder('Historique', 'Déclaration introuvable.', settings)
    );
    return;
  }

  res.send(renderSuperAdminCostDetail(transaction, settings));
});
```

Note: `req.params.id` will throw inside `getTransactionById` (via `new ObjectId(id)`) if it isn't a valid 24-hex-char ObjectId string — e.g. someone hand-typing `/superadmin/costs/abc` in the URL. Express 5 forwards a thrown/rejected error from an async handler to its default error middleware (500), which is an acceptable baseline (matches how this codebase already treats other malformed-input edge cases) — no extra validation needed for this plan's scope.

- [ ] **Step 4: Typecheck**

Run: `npx tsc --noEmit`
Expected: errors pointing at `renderSuperAdminCosts`'s and `renderSuperAdminCostDetail`'s call signatures not matching yet — expected, fixed in Task 4.

- [ ] **Step 5: Commit** (after Task 4 makes this typecheck-clean — see Task 4's final step for the combined commit)

---

### Task 4: Rewrite the Historique list/detail UI in renderSuperAdminDashboard.ts

**Files:**
- Modify: `src/web/renderSuperAdminDashboard.ts`

- [ ] **Step 1: Update the import from `transactionsRepository.js`**

Find:

```ts
import type { TransactionDocument, CountryProductCount } from '../db/transactionsRepository.js';
```

Replace with:

```ts
import type {
  TransactionDocument,
  CountryProductCount,
  ListTransactionsResult,
} from '../db/transactionsRepository.js';
```

- [ ] **Step 2: Add "Historique" back to the sidebar nav**

Correction from the design doc: `NAV_ITEMS` currently has **no `'costs'` entry at all** — an earlier, unrelated request ("hyed cout de produit mn menu") removed it from the sidebar while leaving the route reachable but unlinked, back when the page only showed the single most-recent transaction. Now that it's a real paginated history, it should be navigable again — add it as a 5th item.

Find the `NAV_ITEMS` array's type declaration:

```ts
const NAV_ITEMS: Array<{
  page: SuperAdminPage;
  href: string;
  label: string;
  icon: string;
  color: 'indigo' | 'green' | 'amber' | 'pink';
}> = [
```

Change the `color` union to add a 5th value:

```ts
  color: 'indigo' | 'green' | 'amber' | 'pink' | 'teal';
```

Then find the closing `];` of the `NAV_ITEMS` array (right after the `'settings'` entry) and insert a new entry for `'costs'` before it — order it between `'generate'` and `'users'` (declarations are generated, then reviewed in the history, then accounts/settings):

```ts
  {
    page: 'generate',
    href: '/superadmin/generate',
    label: 'Générer une déclaration',
    icon: '<path d="M10 3v10.5M10 13.5l-4-4M10 13.5l4-4M4 16.5h12" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>',
    color: 'green',
  },
  {
    page: 'costs',
    href: '/superadmin/costs',
    label: 'Historique',
    icon: '<path d="M10 5.5v5l3.5 2M17 10a7 7 0 1 1-2.05-4.95" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>',
    color: 'teal',
  },
  {
    page: 'users',
```

(This shows the `'generate'` and `'users'` entries only as anchors to locate the insertion point — don't duplicate or modify them otherwise.)

- [ ] **Step 3: Add the `.nav-icon-teal` CSS variant**

Find the existing nav-icon color rules:

```css
  .nav-icon-indigo { background: rgba(99, 102, 241, 0.18); color: #a5b4fc; }
  .nav-icon-green { background: rgba(16, 185, 129, 0.18); color: #6ee7b7; }
  .nav-icon-amber { background: rgba(245, 158, 11, 0.18); color: #fcd34d; }
  .nav-icon-pink { background: rgba(236, 72, 153, 0.18); color: #f9a8d4; }
```

Add a 5th line right after:

```css
  .nav-icon-teal { background: rgba(20, 184, 166, 0.18); color: #5eead4; }
```

- [ ] **Step 4: Replace `renderCostsSearchForm` with a combined search + date-range filter form**

Find:

```ts
function renderCostsSearchForm(searchQuery: string): string {
  const clearLink = searchQuery
    ? `<a href="/superadmin/costs" class="search-clear">Effacer</a>`
    : '';
  return `
    <form method="get" action="/superadmin/costs" class="search-form">
      <input type="text" name="q" placeholder="Rechercher par nom / société (redevable)..." value="${escapeHtml(searchQuery)}" />
      <button type="submit" class="search-submit">Rechercher</button>
      ${clearLink}
    </form>
  `;
}
```

Replace with:

```ts
interface CostsFilters {
  q: string;
  dateFrom: string;
  dateTo: string;
}

function renderCostsFilterForm(filters: CostsFilters): string {
  const hasFilters = filters.q || filters.dateFrom || filters.dateTo;
  const clearLink = hasFilters ? `<a href="/superadmin/costs" class="search-clear">Effacer</a>` : '';
  return `
    <form method="get" action="/superadmin/costs" class="search-form">
      <input type="text" name="q" placeholder="Rechercher par nom / société (redevable)..." value="${escapeHtml(filters.q)}" />
      <input type="date" name="dateFrom" value="${escapeHtml(filters.dateFrom)}" aria-label="Date de début" />
      <input type="date" name="dateTo" value="${escapeHtml(filters.dateTo)}" aria-label="Date de fin" />
      <button type="submit" class="search-submit">Filtrer</button>
      ${clearLink}
    </form>
  `;
}
```

- [ ] **Step 5: Rewrite `renderSuperAdminCosts` as the paginated list**

Find the whole existing `renderSuperAdminCosts` function (from `export function renderSuperAdminCosts(` through its closing `}` — it currently branches on `searchQuery && searchResults` vs. the single "most recent" detail view) and replace it entirely with:

```ts
function transactionId(doc: TransactionDocument): string {
  return String((doc as TransactionDocument & { _id?: { toString(): string } })._id);
}

function paginationLink(filters: CostsFilters, page: number): string {
  const params = new URLSearchParams();
  if (filters.q) params.set('q', filters.q);
  if (filters.dateFrom) params.set('dateFrom', filters.dateFrom);
  if (filters.dateTo) params.set('dateTo', filters.dateTo);
  params.set('page', String(page));
  return `/superadmin/costs?${params.toString()}`;
}

// The superadmin's full declaration history — every saved transaction
// across all admins, newest first, filterable by redevable and/or date
// range, 20 per page. Each row links to /superadmin/costs/:id for the
// totals-card detail view (see renderSuperAdminCostDetail below).
export function renderSuperAdminCosts(
  result: ListTransactionsResult,
  settings: AppSettings,
  filters: CostsFilters
): string {
  const filterForm = renderCostsFilterForm(filters);
  const totalPages = Math.max(1, Math.ceil(result.total / result.pageSize));

  const rows = result.items
    .map(
      (item) => `<tr class="clickable-row" onclick="window.location='/superadmin/costs/${transactionId(item)}'">
        <td>${formatDate(item.createdAt)}</td>
        <td>${escapeHtml(item.redevable)}</td>
        <td>${escapeHtml(item.code)}</td>
      </tr>`
    )
    .join('');

  const emptyState =
    result.items.length === 0
      ? `<div class="card placeholder-card"><p>Aucune déclaration ne correspond à ces filtres.</p></div>`
      : '';

  const table =
    result.items.length > 0
      ? `<div class="card">
          <table>
            <thead><tr><th>Date</th><th>Redevable</th><th>Code</th></tr></thead>
            <tbody>${rows}</tbody>
          </table>
          <div class="pagination">
            ${
              result.page > 1
                ? `<a href="${paginationLink(filters, result.page - 1)}" class="pagination-link">← Précédent</a>`
                : '<span class="pagination-link pagination-disabled">← Précédent</span>'
            }
            <span class="pagination-status">Page ${result.page} / ${totalPages}</span>
            ${
              result.page < totalPages
                ? `<a href="${paginationLink(filters, result.page + 1)}" class="pagination-link">Suivant →</a>`
                : '<span class="pagination-link pagination-disabled">Suivant →</span>'
            }
          </div>
        </div>`
      : emptyState;

  const body = `
    <p class="lede">Historique de toutes les déclarations générées sur l'application (${result.total} au total).</p>
    ${filterForm}
    ${table}
  `;
  return renderShell('costs', 'Historique', body, settings, COSTS_SEARCH_STYLE);
}
```

- [ ] **Step 6: Add `renderSuperAdminCostDetail`**

Directly after the new `renderSuperAdminCosts`, add:

```ts
// The detail view for a single declaration, reached by clicking a row in
// the Historique list — reuses the same totals card the old single-
// transaction view and search results already used.
export function renderSuperAdminCostDetail(
  transaction: TransactionDocument,
  settings: AppSettings
): string {
  const body = `
    <a href="/superadmin/costs" class="back-link">← Retour à l'historique</a>
    ${renderSearchResultCard(transaction)}
  `;
  return renderShell('costs', 'Historique', body, settings, COSTS_SEARCH_STYLE);
}
```

(`renderSearchResultCard` already exists above this point in the file and is unchanged.)

- [ ] **Step 7: Add CSS for the clickable rows, pagination, date inputs, and back link**

Find the `COSTS_SEARCH_STYLE` constant and add these rules inside its template literal (anywhere in the block is fine — append at the end before the closing backtick):

```css
  .clickable-row { cursor: pointer; }
  .clickable-row:hover { background: var(--input-bg); }
  .search-form input[type="date"] { width: auto; }
  .pagination {
    display: flex; align-items: center; justify-content: center; gap: 16px;
    padding-top: 16px; margin-top: 4px; border-top: 1px solid var(--line-soft);
  }
  .pagination-link { color: var(--brand-600); text-decoration: none; font-weight: 600; font-size: 13px; }
  .pagination-link:hover { text-decoration: underline; }
  .pagination-disabled { color: var(--ink-400); pointer-events: none; }
  .pagination-status { font-size: 12.5px; color: var(--ink-500); }
  .back-link {
    display: inline-block; margin-bottom: 16px; color: var(--brand-600);
    text-decoration: none; font-weight: 600; font-size: 13.5px;
  }
  .back-link:hover { text-decoration: underline; }
```

- [ ] **Step 8: Typecheck**

Run: `npx tsc --noEmit`
Expected: no errors (this closes out the type mismatches flagged at the end of Task 3).

- [ ] **Step 9: Run the full test suite**

Run: `npx vitest run`
Expected: all tests pass — no existing test exercises `renderSuperAdminCosts`'s exact signature, so this is primarily a regression check on everything else.

- [ ] **Step 10: Commit**

```bash
git add src/web/server.ts src/web/renderSuperAdminDashboard.ts
git commit -m "Replace single-transaction Coût de produit view with a paginated, filterable Historique"
```

---

### Task 5: Manual end-to-end smoke test against a real MongoDB test database

**Files:** none (verification only)

- [ ] **Step 1: Confirm a MongoDB connection string is available**

This requires `MONGODB_URI` pointing at a real reachable MongoDB (Atlas or otherwise) — check `process.env.MONGODB_URI` / `.env`. If unavailable in this environment, skip straight to Step 5 and note in the final report that live-server verification wasn't possible here (the automated tests in Tasks 1–4 already cover the logic; only the live HTTP flow is unverified).

- [ ] **Step 2: Start the app against a disposable test database**

Run (adjust the database name suffix in the URI to something like `customs_app_test_history`, a throwaway database — never point this at the real production database):

```bash
DATABASE_NAME_SUFFIX_NOTE="use a disposable test db name in MONGODB_URI, e.g. .../customs_app_test_history?..." 
SUPERADMIN_USERNAME=redwan SUPERADMIN_PASSWORD=redwan2026 PORT=5790 npx tsx src/web/server.ts &
```

- [ ] **Step 3: Seed a handful of transactions and log in**

Use curl (with a cookie jar) to log in as `redwan`/`redwan2026`, then generate 2–3 declarations via `/generate` with different sample Liquidation/DUM fixture pairs (or directly `insertOne` a few documents into the test database's `transactions` collection with varied `redevable`/`createdAt` values via a short throwaway Node script, whichever is faster) so the list has enough rows to exercise pagination (e.g. 25+ rows) and filtering.

- [ ] **Step 4: Verify the list, filters, pagination, and detail page**

- `GET /superadmin/costs` — confirm the sidebar label reads "Historique", the table shows Date/Redevable/Code, and pagination controls appear once there are more than 20 rows.
- `GET /superadmin/costs?q=<partial redevable>` — confirm filtering narrows the list correctly.
- `GET /superadmin/costs?dateFrom=...&dateTo=...` — confirm date-range filtering works.
- Click a row (or `GET /superadmin/costs/<id>` directly) — confirm the detail card (coût total, taxes, valeur déclarée) renders with a working "← Retour à l'historique" link.
- Confirm the zero-transactions and Mongo-unreachable placeholder states still render (temporarily unset `MONGODB_URI` to check the latter).

- [ ] **Step 5: Clean up the test database**

```bash
node -e "
const { MongoClient } = require('mongodb');
(async () => {
  const client = new MongoClient(process.env.MONGODB_URI);
  await client.connect();
  await client.db().dropDatabase();
  await client.close();
})();
"
```

Then stop the local server process.

- [ ] **Step 6: Final report**

Summarize in the conversation: what was verified live vs. only unit-tested, and push confirmation that Tasks 1–4's commits are already on `main`.

# Historique des déclarations — Design

## Context

Part 1 of 5 requested features (the others — export groupé, notifications,
audit log, limites d'usage — are separate future sub-projects, built one at
a time). The superadmin's "Coût de produit" page currently shows only the
single most-recently-generated transaction, plus an ad-hoc search by
redevable. This replaces it with a real paginated history.

## Scope

- Replace the content (and sidebar label) of the existing
  `/superadmin/costs` route with a paginated, filterable list of all saved
  transactions.
- Add a detail route for viewing one transaction's cost breakdown.
- Preserve the existing search-by-redevable capability.
- Add a new date-range filter (période).

Out of scope (future sub-projects): grouped Excel/PDF export over a date
range, email/WhatsApp notifications, an audit log, and usage limits/alerts.
This design does lay groundwork the export feature will reuse (the paginated
+ date-filtered query), but does not build export itself.

## Data layer

New function in `src/db/transactionsRepository.ts`:

```ts
export interface ListTransactionsOptions {
  page: number; // 1-indexed
  pageSize: number;
  redevable?: string; // case-insensitive substring match, same as existing search
  dateFrom?: string; // ISO date, inclusive
  dateTo?: string; // ISO date, inclusive
}

export interface ListTransactionsResult {
  items: TransactionDocument[];
  total: number;
  page: number;
  pageSize: number;
}

export async function listTransactionsPaginated(
  collection: Collection<TransactionDocument>,
  options: ListTransactionsOptions
): Promise<ListTransactionsResult>
```

Sorted by `createdAt` descending. Filter combines redevable regex (reusing
the existing escape-and-case-insensitive approach from
`searchTransactionsByRedevable`) and a `createdAt` range on the ISO string
(lexicographic comparison works since it's always full ISO 8601). Uses
`.skip((page-1)*pageSize).limit(pageSize)` plus a separate `countDocuments`
call for `total`.

## Routes (`src/web/server.ts`)

- `GET /superadmin/costs` — reads `page`, `q` (redevable), `dateFrom`,
  `dateTo` from query params (defaults: page=1, pageSize=20, no filters).
  Calls `listTransactionsPaginated`, renders the list. Existing
  Mongo-unreachable / zero-transactions-ever placeholder handling is kept
  as-is.
- `GET /superadmin/costs/:id` — looks up one transaction by `_id`, renders
  it via the existing `renderSearchResultCard` markup (same card already
  used for search results today), with a "← Retour à l'historique" link
  back to `/superadmin/costs` (preserving the referring page/filters via a
  `?from=` query param would be nice-to-have; plain back-to-page-1 link is
  the baseline).

## UI (`src/web/renderSuperAdminDashboard.ts`)

- `renderSuperAdminCosts` is rewritten to take a `ListTransactionsResult`
  instead of a single "most recent" transaction.
- Filter bar: existing redevable search input, plus two new `<input
  type="date">` fields (dateFrom/dateTo), all as one GET form.
- Table: Date | Redevable | Code, each row a link to
  `/superadmin/costs/:id`.
- Pagination: "Précédent" / "Suivant" links under the table, preserving
  current filters via hidden query params; disabled/hidden at the first
  and last page respectively.
- New `renderSuperAdminCostDetail` function for the `:id` route, wrapping
  the existing card-rendering logic.
- `NAV_ITEMS` label for this page changes from "Coût de produit" to
  "Historique".

## Testing

- Repository unit test for `listTransactionsPaginated` (fake collection):
  pagination math, date-range filtering, redevable filtering, combined
  filters, empty results.
- Manual smoke test against a disposable real MongoDB test database
  (existing project pattern): seed a handful of transactions across
  different dates/redevables, verify list/filter/pagination/detail-view
  all behave correctly, then drop the test database.

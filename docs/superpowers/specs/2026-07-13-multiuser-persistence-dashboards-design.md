# Multi-User Persistence, Cost Calculation & Dashboards — Design Spec

Date: 2026-07-13
Status: Approved for implementation

## 1. Problem & Goal

The app is currently a single-user, one-shot tool: one hardcoded login
(`redwan`/`redwan2026`), no persistence, no history. This phase turns it
into a small multi-admin tool:

- Multiple admins, each with their own login credentials.
- A superadmin role that manages admin accounts and sees aggregate analytics
  across all admins' work.
- Every generated declaration is saved (not just downloaded once), so an
  admin can revisit their history and re-download past Excel files.
- The Excel output gains a real business number: **total landed cost per
  physical unit** (purchase price + freight + insurance + customs duties),
  not just the duty/tax breakdown it already computes.

This explicitly reverses phase 1's "one-shot, no persistence" decision —
that was correct for the original single-user prototype; it no longer fits
a multi-admin tool that needs history and cross-admin analytics.

## 2. Roles & Auth

Two roles: `admin` and `superadmin`. A `superadmin` can do everything an
`admin` can, plus manage users and view analytics.

- `users` table replaces the hardcoded credential check in `src/web/auth.ts`.
  Passwords are hashed (bcrypt).
- Exactly one superadmin account is seeded on first boot from
  `SUPERADMIN_USERNAME` / `SUPERADMIN_PASSWORD` environment variables (if
  the `users` table is empty). All other accounts — additional admins, or
  additional superadmins — are created from the superadmin dashboard, not
  from environment variables.
- Session-based auth (same mechanism as today — signed cookie session
  referencing a session record), extended to carry `userId` and `role`.
- Two middleware guards: `requireAuth` (any logged-in user — replaces
  today's `requireAuth`) and `requireSuperAdmin` (superadmin only, for
  `/superadmin/*` routes).
- A disabled admin (`disabled_at` set) can no longer log in, but their past
  declarations remain in the history/analytics (not deleted).

## 3. Persistence

SQLite via `better-sqlite3` (synchronous API, no separate DB server,
proven to deploy fine in this project's Coolify/Nixpacks setup the same
way `sharp`'s native bindings already do). Single file, `data.db`, created
on boot if absent; path configurable via `DATABASE_PATH` env var (defaults
to a path under the project root, alongside the existing `.tmp-output`
convention).

```sql
CREATE TABLE users (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('admin', 'superadmin')),
  created_at TEXT NOT NULL,
  disabled_at TEXT
);

CREATE TABLE declarations (
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
  cost_estimate_partial INTEGER NOT NULL DEFAULT 0, -- 1 if fret/assurance/montant_facture/taux_change were unavailable
  excel_file_path TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TABLE declaration_articles (
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
```

`declaration_articles` denormalizes each article's computed `cost_per_unit`
(so the dashboard can list it without recomputing) and stores the full tax
breakdown as JSON (`taxes_json`) since it's only ever read back whole, never
queried by tax code.

Generated `.xlsx` files are written to a persistent directory
(`OUTPUT_DIR`, currently `.tmp-output` — renamed conceptually but not
necessarily on disk to reflect that files now persist rather than being
overwritten per-request) with a unique filename per declaration
(`declaration-<id>.xlsx`), referenced by `excel_file_path`.

## 4. Extended DUM Extraction

The DUM PDF already contains, as plain extractable text (confirmed against
the real sample), fields we don't currently parse: currency + invoiced
amount (`"EUR 2 981.340"`), exchange rate (`"10.667500"`), freight
(`"7 467.000"`), insurance (`"118.000"`), and total declared value
(`"40 039.992"`). `DumParser` gains a new top-level (document-level, not
per-article) extraction for these five fields, using the same
positional-matching strategy as the article extraction (§ DumParser in the
phase-1/2 spec) — anchored on the recognizable `"EUR|MAD|USD ... rate ...
freight-value"` numeric cluster near the currency code, not on scattered
labels. These become optional fields on `DumResult` (`undefined` if the
document doesn't contain them or they don't match — never a hard parse
failure, since the app's core function doesn't depend on them).

## 5. Cost Per Unit Calculation

For each article, its share of shipment-level costs (invoiced amount +
freight + insurance) is allocated proportionally to its share of the
shipment's total declared value — the same proportional-allocation
principle already used for tax-per-unit splitting, applied one level up:

```
article_share = article.valeurDeclaree / sum(all articles' valeurDeclaree)
article_shipment_cost = (montantFacture + fret + assurance) × tauxChange × article_share
article_total_cost = article_shipment_cost + sum(article.taxes[].montant)
cost_per_unit = article_total_cost / article.quantite
```

If `montantFacture`/`fret`/`assurance`/`tauxChange` are unavailable (DUM
extraction didn't find them), `cost_per_unit` falls back to
duty-only-per-unit (today's existing calculation) and the declaration is
flagged (a boolean column, `cost_estimate_partial`) so the UI can show "coût
partiel (fret/assurance non détectés)" rather than presenting an
incomplete number as if it were complete.

## 6. Manual Admin Inputs

Two fields the documents don't contain, entered by the admin after
generation and before the declaration is saved to history:
- **Colis** (package count) — integer.
- **Référence interne** — free-text string, for the admin's own
  cross-referencing (e.g. a PO number).

These are entered on a short intermediate form shown after Excel generation
succeeds, before the record is written to `declarations`. Both are
optional (an admin can save without filling them, and edit them later from
the history view).

## 7. UI Flows

**`/login`** — unchanged visual design, now checks `users` table.

**`/admin/dashboard`** (default landing page for both roles after login):
- "Nouvelle déclaration" button → today's upload flow (`/`), unchanged
  through generation.
- After generation succeeds: a short form (Colis, Référence interne) before
  the record is saved — replaces today's "download and done" ending.
- **History table**: this admin's own declarations only (code, redevable,
  date, total landed cost, colis, référence, download button). Superadmins
  also have their own personal history here (they can generate
  declarations too, same as any admin).

**`/superadmin/dashboard`** (superadmin only, linked from the admin
dashboard's nav for superadmin users):
- **Analytics**: total declarations, total landed cost (all admins
  combined), a simple bar chart by month and by admin. Server-rendered
  (an HTML `<canvas>` with vanilla JS drawing, or plain HTML/CSS bars — no
  charting library dependency needed at this data scale).
- **User management**: table of all admin accounts (username, created date,
  active/disabled), a form to create a new admin (username + password), a
  disable/enable toggle per account. Superadmin cannot disable their own
  account (guard against self-lockout).

## 8. What's Explicitly Out of Scope (this phase)

- Editing/deleting a saved declaration's parsed data (history is
  append-only; correcting a mistake means generating a new declaration).
- Per-admin permissions finer than admin/superadmin (e.g. "admin A can see
  admin B's history") — history is strictly own-data-only for admins.
- Password reset flow (superadmin resets a stuck admin's password
  manually by disabling + recreating the account, for now).
- Multi-currency conversion beyond using the DUM's own `tauxChange` as-is
  (no external exchange-rate API).

## 9. Migration Notes

- The existing `checkCredentials`/session code in `src/web/auth.ts` is
  replaced, not extended — the hardcoded `APP_USERNAME`/`APP_PASSWORD` env
  vars are removed in favor of `SUPERADMIN_USERNAME`/`SUPERADMIN_PASSWORD`
  (first-boot seed only).
- `combinedExcelGenerator`, `articleSummaryExcelGenerator`, and
  `unitLevelExcelGenerator` are unchanged — cost-per-unit is a new column
  computed and stored alongside the existing generation flow, added to
  `declaration_articles`, not a change to the Excel sheets themselves in
  this phase. (Adding a cost-per-unit column to the Excel output itself is
  a natural follow-up but is not required for this phase's stated goal —
  the dashboard is where landed cost is surfaced.)

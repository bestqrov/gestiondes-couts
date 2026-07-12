# Customs Declaration → Excel Automation — Design Spec

Date: 2026-07-11
Status: Approved for implementation (phase 1 starting)

## 1. Problem & Goal

An importing company receives, per shipment, two fixed-format Moroccan customs PDF
documents that never change layout:

1. **Liquidation Douanière (BE — Bulletin d'Engagement / Crédit d'enlèvement)**
   — per-article tax breakdown (TAXE code, ASSIETTE, TAUX, MONTANT, TOTAL ARTICLE).
2. **DUM (Déclaration Unique de Marchandises)**
   — the master declaration: per-article HS code, country of origin, declared
   value, quantity, and item description.

The two documents describe the **same shipment** and the **same set of articles**,
linked by:
- a shared reference code (the "Crédit d'enlèvement" number, e.g. `500001` /
  `500001099` — the DUM's field includes a 3-digit suffix not present in the
  Liquidation header code), and
- article ordering / article number (`ARTICLE : 1` in Liquidation ↔
  `N° d'ordre de l'art. : 1` in DUM).

The user uploads both documents; the app must parse both, **merge them per
article**, let the user correct any low-confidence/missing field, and generate
two Excel files. The end goal (business purpose) is per-unit **cost
calculation**: knowing the true landed cost of each physical piece in a
shipment, including its share of duties/taxes.

## 2. Inputs

- Exactly 2 files per run: one Liquidation PDF, one DUM PDF (or scanned image
  of either). No other input formats (e.g. packing-list Excel) are consumed by
  the system — those exist only as external reference data the user may look
  at separately.
- The UI must let the user pick which uploaded file is which (or the app
  detects it — see §5.2) since a user could upload them in either order.

## 3. Architecture

Five independent layers; each depends only on the interface of the layer below.

```
Input (2 PDFs/images: Liquidation + DUM)
   → OCR Layer            (per-document text extraction)
   → Parser Layer          (deterministic per-document-type parsing → raw fields + confidence)
   → Merge Layer            (join Liquidation + DUM records by ref code + article number)
   → Domain Layer            (validated Declaration/Article/TaxLine entities, business rules)
   → Excel Generator Layer    (streaming writers for File 1 and File 2)
   → UI Layer                  (file picker, correction screen, export)
```

The **Merge Layer** is new since the two-document discovery (§ history) — it is
the piece that reconciles two independently-parsed documents into one
validated `Article`. It is intentionally separate from both parsers (each
parser only knows its own document type) and from the Domain layer (which only
knows about already-merged, valid data).

### 3.1 OCR Layer

- **Born-digital PDFs**: extract embedded text layer via `pdfjs-dist` — fast,
  ~100% accurate, tried first for each of the two input files independently.
- **Scanned PDFs/images**: rasterize pages, then OCR via `tesseract.js`
  (bundled WASM, no system Tesseract install needed — important for Electron
  packaging), French+numeric language pack.
- Output: normalized `OcrResult` — `{ text, confidence, page }[]`. Text-layer
  extraction always reports confidence `1.0`.
- Each of the two uploaded files goes through this layer independently and
  produces its own `OcrResult`.

### 3.2 Parser Layer

Two parsers, one per document type, sharing a common line-based
state-machine approach (not coordinate/positional matching) driven by a
declarative rule config (label anchors + regex), so a future layout variant
is a new rule set, not new logic:

- **LiquidationParser** → `LiquidationDeclaration { code, redevable, benNumero, dates, articles: LiquidationArticle[] }`
  where `LiquidationArticle { numero, hsCode, valeur, quantite, taxes: TaxLine[], totalArticle }`
  and `TaxLine { code, assiette, taux, montant }` — tax `code` set is fully
  dynamic (whatever appears in the TAXE column: `000110`, `007217`, `002109`,
  others).
- **DumParser** → `DumDeclaration { creditEnlevementCode, articles: DumArticle[] }`
  where `DumArticle { ordre, codeMarchandises (HS), designation (nom article), paysOrigine, paysCode, valeurDeclaree, quantite, unite }`.

Every extracted field carries `{ value, confidence, source: 'ocr'|'derived'|'missing' }`.

Document-type detection: each parser first checks for its own anchor markers
(`"LIQUIDATION DES DROITS ET TAXES"` header repeated per-article block vs.
`"DECLARATION"` / `"Mod. D.U.M"` header) so the app can auto-detect which
uploaded file is which regardless of upload order.

### 3.3 Merge Layer

`DeclarationMerger.merge(liquidation, dum) → MergedArticle[]`

- Matches the two documents by reference code: DUM's `creditEnlevementCode`
  (e.g. `500001099`) must start with / contain Liquidation's `code`
  (e.g. `500001`) — exact matching rule confirmed against real samples in
  phase 1 tests, with a hard validation error (not silent skip) if no match is
  found, since mismatched shipments must never be silently combined.
- Matches individual articles by `numero` (Liquidation) == `ordre` (DUM).
- Cross-validates overlapping fields when both docs provide them (e.g. HS
  code, valeur, quantité must agree between the two sources within tolerance);
  disagreement is a flagged validation error surfaced in the correction UI,
  never silently resolved by picking one source.
- Produces one `MergedArticle` per matched article pair, carrying: HS code,
  nom article (DUM), pays + pays code (DUM), valeur déclarée (DUM, cross-
  checked against Liquidation), quantité + unité (DUM, cross-checked), and the
  full `taxes: TaxLine[]` + `totalArticle` (Liquidation).
- An article present in only one document is a **validation error**, not a
  partial success — surfaced to the user for correction, not silently dropped.

### 3.4 Domain Layer

```ts
Declaration { code, redevable, benNumero, dates..., articles: Article[] }
Article { numero, hsCode, nomArticle, pays, paysCode, valeurDeclaree, quantite, unite, taxes: TaxLine[], totalArticle }
TaxLine { code, assiette, taux, montant }
```

Validation rules enforced here (in addition to the Merge Layer's cross-doc
checks):
- `quantite > 0` (required for row expansion in File 2).
- `sum(taxes.montant) == totalArticle` within rounding tolerance; mismatch is
  flagged, not silently accepted.
- All required fields present (no `missing` source flags remaining) or
  explicitly confirmed by the user in the correction UI.

This is where "accuracy for an importing company" is enforced: nothing
reaches Excel generation without passing validation or being explicitly
user-confirmed.

### 3.5 Excel Generators

Both generators consume the validated `Declaration` domain object and use
`ExcelJS`'s streaming `WorkbookWriter` (constant memory regardless of row
count).

**File 1 — Article Summary** (one row per article):
| Nom Article | HSC | Pays | Valeur déclarée | Unité (Quantity) |

**File 2 — Unit-Level Cost Detail** (one row per physical unit):
- Columns: `Nom Article | HSC | Serial Number (1..Quantité) | <tax code 1> | <tax code 2> | ... `
- Tax code columns are the **union of all codes seen across the merged
  declaration's articles** (fully dynamic — not a fixed set of 4), so the
  sheet has consistent columns even when articles have different code sets
  (article missing a given code ⇒ 0 for that column).
- Per confirmed requirement: for each article, generate exactly `quantite`
  rows; each tax code's `montant` is divided equally across those rows
  (`montant / quantite`, rounded to 2 decimals). Rounding remainder policy:
  any leftover cents from rounding are added to the **last** row of the
  article so column sums always reconcile exactly to the source `montant`.

### 3.6 Correction UI (Electron + React + TypeScript)

Flow: **Login** → **Upload screen** (drag-and-drop or file picker for the 2
files) → auto-detect which is Liquidation vs DUM (with manual override if
detection is ambiguous) → processing screen (OCR → parse → merge → validate,
with progress) → **Correction screen**: split view (original page image
alongside an editable form of merged fields per article), low-confidence/
missing/mismatched fields highlighted with their flag reason (e.g. "OCR
confidence low", "value differs between the two documents: 27147.000 vs
27146.000") → confirm → **Generate** button → **Results screen**.

State: plain React state/context — single declaration in flight, no history
(confirmed: one-shot workflow, no persistence/database).

#### 3.6.1 Login screen

Single hardcoded credential pair (`redwan` / `redwan2026`), checked
client-side against an env-configured value — no user database, no
multi-user support, no password reset flow. This is a simple access gate for
an internal single-user tool, not a security boundary; it must not be
mistaken for one if the app is ever exposed outside a trusted local network.
Session persists in memory for the app's lifetime (re-prompted on restart).

#### 3.6.2 Upload screen

Drag-and-drop zone (plus a fallback file picker button) accepting exactly 2
files (PDF or image). Each dropped/selected file is auto-classified as
Liquidation or DUM per the §3.2 detection rule, shown back to the user with
its detected type and a manual override control in case detection is wrong
or ambiguous. Upload is blocked from proceeding until exactly one Liquidation
and one DUM file are present.

#### 3.6.3 Results screen

After the user confirms the correction screen and clicks **Generate**, the
app renders a screen split into two panels — one per output file:

- **Left: File 1 (Article Summary)** — full HTML table preview (small,
  one row per article) plus a "Download Excel" button.
- **Right: File 2 (Unit-Level Cost Detail)** — HTML table preview capped to
  the first N rows (e.g. 200) with a note ("showing 200 of 12,400 rows —
  download for full detail") since this file can reach tens of thousands of
  rows and must never be fully rendered in the DOM, plus a "Download Excel"
  button. The preview reads from the same in-memory generation result used
  for the full-file download — it does not re-run generation.

## 4. Tech Stack

- Electron + React + TypeScript, Vite bundler
- `pdfjs-dist` (PDF text-layer + rasterization), `tesseract.js` (scanned-image
  OCR, no native binary dependency)
- `exceljs` streaming `WorkbookWriter`
- `vitest` for unit tests

No database — one-shot workflow.

## 5. Folder Structure

```
src/
  main/                    Electron main process: file dialogs, IPC
  ocr/                     PdfTextExtractor, TesseractOcrEngine, OcrResult types
  parser/
    liquidation/           LiquidationParser + rule config
    dum/                    DumParser + rule config
  merge/                    DeclarationMerger
  domain/                   entities, validators
  excel/                     ArticleSummaryExcelGenerator, UnitLevelExcelGenerator
  ui/                         React app: UploadScreen, CorrectionScreen, ExportScreen
  shared/                     types shared across processes
tests/
  parser/fixtures/            sample OCR text + expected parsed output, per doc type
  merge/fixtures/              paired liquidation+dum fixtures with expected merged output
docs/
  Architecture.md
  superpowers/specs/           this file and future specs
```

## 6. Testing Strategy

- Parser: fixture-based unit tests (raw text in → per-document parsed object
  out) for both LiquidationParser and DumParser, using the two real sample
  documents provided (transcribed as text fixtures).
- Merge: fixture tests pairing a Liquidation + DUM fixture → expected merged
  `Article[]`, including a mismatch case (to verify validation errors fire).
- Excel generators: snapshot tests on small inputs (row/column shape) plus a
  performance test asserting streaming memory stays flat at 10k+ rows, and a
  rounding-reconciliation test (column sums equal source `montant` exactly).
- OCR layer: smoke-tested manually against the real sample documents (OCR
  accuracy isn't meaningfully unit-testable).

## 7. Roadmap (phased)

1. **Domain + Parsers (both) + Merge Layer** — built and tested against text
   fixtures transcribed from the two real sample documents. No OCR/UI yet.
2. **Excel Generators** — wire merged domain output into both streaming
   generators; verify with generated sample files, including the rounding
   reconciliation behavior.
3. **OCR Layer** — PDF text-layer extraction + Tesseract fallback, wired to
   real files end-to-end (parsing real PDFs, not just fixture text).
4. **Electron/React UI** — dual file upload with auto-detection, correction
   screen with side-by-side image + form, export flow.

## 8. Open Items / Assumptions Carried Into Phase 1

- Exact string-matching rule between Liquidation's `code` and DUM's
  `creditEnlevementCode` (e.g. `500001` vs `500001099`) will be finalized and
  covered by a test once phase 1 fixtures are built from the two real samples
  — the current assumption is prefix-match.
- Only these two document types are in scope; the packing-list Excel shown
  during design discussion is external reference data, not a system input.

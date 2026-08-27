import ExcelJS from 'exceljs';
import type { PackingListRow } from '../parser/packingList/packingListParser.js';
import type { Declaration } from '../domain/types.js';
import { normalizeCountryName } from '../domain/countryNames.js';
import {
  allocateTaxAcrossUnits,
  unionTaxCodes,
  taxCodeDesignation,
} from './unitLevelTaxHelpers.js';
import {
  styleDataRow,
  styleHeaderRowGrouped,
  styleColumnsFill,
  addSheetTitleRows,
  resolveBrandArgb,
  resolveBrandDarkArgb,
  resolveCompanyName,
  resolveDocumentTitle,
  SUM_HT_ROW_FILL_ARGB,
  SUM_TTC_ROW_FILL_ARGB,
  type BrandingInfo,
  type ColumnGroup,
} from './excelStyling.js';

// TVA IMPORT AUTRE PDS — excluded from Somme DD HT (Hors Taxe: duties
// without VAT) but still counted in Somme DD TTC (Toutes Taxes Comprises),
// same HT/TTC distinction the two column pairs' names already promise.
const VAT_TAX_CODE = '002109';

// Manually-entered shipment-wide costs (not present in any parsed document)
// that get spread across every HS total row by that row's matched PRORATA
// share, one column each, right after PRORATA — same "total x prorata"
// spread the PRORATA column itself documents. Order here is the order the
// columns appear in on the sheet. These 6 are always present; a superadmin
// can add further custom ones (persisted in AppSettings.extraCostFields —
// see appSettingsRepository.ts), which get appended after these in the
// order they were added.
export const BASE_EXTRA_COST_FIELDS = [
  { key: 'fraisTransport', label: 'Montant Frais transport' },
  { key: 'assurance', label: 'Montant Assurance' },
  { key: 'fraisLocaux', label: 'Frais locaux passage mead' },
  { key: 'transit', label: 'Transit' },
  { key: 'transportNational', label: 'Transport national' },
  { key: 'mcia', label: 'MCIA' },
] as const;

export interface ExtraCostField {
  key: string;
  label: string;
}

export type ExtraCostFieldKey = string;

export type ExtraCosts = Record<string, number>;

const BASE_COLUMN_COUNT = 8;
const DESCRIPTION_COLUMN = 2;
const PIECES_COLUMN = 4;
const UNIT_COLUMN = 5;
const TOTAL_COLUMN = 6;
const BASE_COLUMN_GROUPS: ColumnGroup[] = [
  { kind: 'identity', from: 1, to: 3 }, // item, DESCRIPTION, color
  { kind: 'quantity', from: 4, to: 4 }, // pieces
  { kind: 'value', from: 5, to: 6 }, // unit, total
  { kind: 'identity', from: 7, to: 8 }, // origin, HS CODE
];

// Only the first 6 digits (the HS "position") need to agree to consider a
// packing-list row and a declaration article the same product — matches the
// HS_CODE_COMPARISON_LENGTH convention already used to reconcile the
// Liquidation and DUM's HS codes in declarationMerger.ts. Product names
// differ too much between the supplier's packing list and the customs
// documents' OCR'd text to match reliably by name.
const HS_CODE_MATCH_LENGTH = 6;

function hsCodePrefix(code: string): string {
  return code.slice(0, HS_CODE_MATCH_LENGTH);
}

// Same HS position can be sourced from more than one country, each with its
// own tax montants (e.g. two variants of the same product, one from China,
// one from Bangladesh) — so, when that actually happens, the group total a
// packing-list row picks up needs to key on HS code prefix AND country of
// origin together, not HS code alone, or a row would silently pick up
// another country's tax montants.
function hsAndOriginKey(hsCode: string, pays: string): string {
  return `${hsCodePrefix(hsCode)}|${normalizeCountryName(pays)}`;
}

// The per-unit tax montants of an article's very first physical unit (the
// same allocation Global's row 1 for this article shows) — not the article's
// full total. Requested explicitly: a packing-list row picks up a single
// unit's tax value from the matching Global rows, not the sum across every
// matching unit/article. Also folds in the declaration-wide-only
// ordonnancement taxes (e.g. REDV.INF., RI SEGMA, REMISES CREDIT) using the
// same montant × Prorata treatment Global gives them, so every tax visible
// in Global has a matching value here. `prorata` is Global's own Prorata
// value for this article (same on every one of its unit rows), picked up
// the same "first matching line, not a sum" way for the PRORATA column.
interface FirstUnitData {
  taxes: Map<string, number>;
  prorata: number;
  // The matched declaration article's own full HS code (Global's HSC),
  // shown in place of the packing list's own hsCode — the packing list's
  // code is often shorter/less precise (only what the supplier entered),
  // while the declaration's is the authoritative customs one Global itself
  // displays.
  hsCode: string;
  // The declaration's own country of origin and product name — used as the
  // group's canonical key and display values once every packing-list row
  // matching this HS+origin group is merged into one HS-total row.
  pays: string;
  nomArticle: string;
}

function firstUnitTaxes(
  article: Declaration['articles'][number],
  extraOrdonnancementTaxes: Declaration['ordonnancementTaxes'],
  declarationValeurDeclareeTotal: number
): FirstUnitData {
  const quantite = Math.round(article.quantite);
  const perCode = new Map<string, number>();
  for (const tax of article.taxes) {
    perCode.set(tax.code, quantite > 0 ? allocateTaxAcrossUnits(tax.montant, quantite)[0] : 0);
  }
  const prorata =
    quantite > 0 && declarationValeurDeclareeTotal > 0
      ? article.valeurDeclaree / quantite / declarationValeurDeclareeTotal
      : 0;
  if (quantite > 0 && declarationValeurDeclareeTotal > 0) {
    for (const tax of extraOrdonnancementTaxes) {
      perCode.set(tax.code, tax.montant * prorata);
    }
  } else {
    for (const tax of extraOrdonnancementTaxes) {
      perCode.set(tax.code, 0);
    }
  }
  return { taxes: perCode, prorata, hsCode: article.hsCode, pays: article.pays, nomArticle: article.nomArticle };
}

// Maps HS code prefix + country of origin to the first-encountered matching
// article's first-unit tax montants — multiple articles (e.g. one per
// color/variant) can share the same HS position and origin, the same way
// multiple packing-list rows do, but only the first one encountered
// (declaration.articles order, matching Global's row order) is used, per
// the same "just the first matching line" rule.
function firstUnitTaxesByHsAndOrigin(
  declaration: Declaration,
  extraOrdonnancementTaxes: Declaration['ordonnancementTaxes'],
  declarationValeurDeclareeTotal: number
): Map<string, FirstUnitData> {
  const result = new Map<string, FirstUnitData>();
  for (const article of declaration.articles) {
    const key = hsAndOriginKey(article.hsCode, article.pays);
    if (result.has(key)) continue;
    result.set(key, firstUnitTaxes(article, extraOrdonnancementTaxes, declarationValeurDeclareeTotal));
  }
  return result;
}

// Same as firstUnitTaxesByHsAndOrigin, but grouped by HS code prefix alone —
// the fallback used for the (overwhelmingly common) case where an HS prefix
// has only one country of origin in the whole declaration, so origin
// spelling (packing list vs DUM/Liquidation, potentially a country name
// outside countryNames.ts's known aliases) can't cause a false non-match.
function firstUnitTaxesByHsPrefix(
  declaration: Declaration,
  extraOrdonnancementTaxes: Declaration['ordonnancementTaxes'],
  declarationValeurDeclareeTotal: number
): Map<string, FirstUnitData> {
  const result = new Map<string, FirstUnitData>();
  for (const article of declaration.articles) {
    const prefix = hsCodePrefix(article.hsCode);
    if (result.has(prefix)) continue;
    result.set(prefix, firstUnitTaxes(article, extraOrdonnancementTaxes, declarationValeurDeclareeTotal));
  }
  return result;
}

// An article's own Prorata share of the whole declaration's declared value —
// summing this across every one of an article's unit rows in Global (each
// showing valeurDeclareePerUnit / declarationValeurDeclareeTotal) gives back
// exactly this same fraction, so it's the group sum this article contributes
// to hsAndOriginProrataSums/hsPrefixProrataSums below.
function articleProrataShare(
  article: Declaration['articles'][number],
  declarationValeurDeclareeTotal: number
): number {
  return declarationValeurDeclareeTotal > 0 ? article.valeurDeclaree / declarationValeurDeclareeTotal : 0;
}

// Maps HS code prefix + country of origin to the SUM of every matching
// article's own Prorata share — requested explicitly: a packing-list row's
// PRORATA should reflect its whole HS+origin group's combined share of the
// declaration (e.g. several color variants of the same product/origin), not
// just the first matching article's own share the way firstUnitTaxes' taxes
// still do.
function hsAndOriginProrataSums(
  declaration: Declaration,
  declarationValeurDeclareeTotal: number
): Map<string, number> {
  const result = new Map<string, number>();
  for (const article of declaration.articles) {
    const key = hsAndOriginKey(article.hsCode, article.pays);
    const share = articleProrataShare(article, declarationValeurDeclareeTotal);
    result.set(key, (result.get(key) ?? 0) + share);
  }
  return result;
}

// Same as hsAndOriginProrataSums, but grouped by HS code prefix alone — the
// fallback used for the same unambiguous-prefix case firstUnitTaxesByHsPrefix
// handles.
function hsPrefixProrataSums(
  declaration: Declaration,
  declarationValeurDeclareeTotal: number
): Map<string, number> {
  const result = new Map<string, number>();
  for (const article of declaration.articles) {
    const prefix = hsCodePrefix(article.hsCode);
    const share = articleProrataShare(article, declarationValeurDeclareeTotal);
    result.set(prefix, (result.get(prefix) ?? 0) + share);
  }
  return result;
}

// Sums every article's own quantite/valeurDeclaree into its HS+origin
// group — used to size a synthetic row standing in for a whole declaration
// group the packing list has no line for at all (see the synthetic-row loop
// in addPackingListSheet below), the same way a real packing-list row's own
// pieces/total size a matched group's row.
function quantiteAndValeurTotalsByHsAndOrigin(
  declaration: Declaration
): Map<string, { quantite: number; valeur: number }> {
  const result = new Map<string, { quantite: number; valeur: number }>();
  for (const article of declaration.articles) {
    const key = hsAndOriginKey(article.hsCode, article.pays);
    const totals = result.get(key) ?? { quantite: 0, valeur: 0 };
    totals.quantite += Math.round(article.quantite);
    totals.valeur += article.valeurDeclaree;
    result.set(key, totals);
  }
  return result;
}

// HS prefixes with more than one distinct (normalized) country of origin
// among the declaration's articles — only for these does origin actually
// need to be part of the match; every other prefix is unambiguous by HS
// code alone, and forcing an origin match there would only risk a spelling
// mismatch losing a row's taxes for no benefit.
function hsPrefixesWithMultipleOrigins(declaration: Declaration): Set<string> {
  const originsByPrefix = new Map<string, Set<string>>();
  for (const article of declaration.articles) {
    const prefix = hsCodePrefix(article.hsCode);
    const origins = originsByPrefix.get(prefix) ?? new Set<string>();
    origins.add(normalizeCountryName(article.pays));
    originsByPrefix.set(prefix, origins);
  }
  const ambiguous = new Set<string>();
  for (const [prefix, origins] of originsByPrefix) {
    if (origins.size > 1) ambiguous.add(prefix);
  }
  return ambiguous;
}

// Adds the "HS total" sheet — one row per declaration HS+origin group,
// merging every uploaded packing-list row matched to it (e.g. several
// color/variant lines all sharing the same product/origin used to each show
// up as their own row, all repeating that group's same Prorata — reported as
// confusing/duplicated). pieces/total sum across every matching row; PRORATA
// is that group's already-summed share (see hsAndOriginProrataSums). A
// packing-list row matching no declaration group still shows up
// individually afterward, under its own raw HS code/origin.
export async function addPackingListSheet(
  workbook: ExcelJS.Workbook,
  rows: PackingListRow[],
  declaration: Declaration,
  branding: BrandingInfo,
  generatedAt: Date,
  extraCosts: ExtraCosts = {},
  // Which extra-cost columns to render, and in what order — defaults to the
  // 6 built-in fields; a superadmin's custom fields (AppSettings.extraCostFields)
  // get passed in appended after those by the caller.
  extraCostFields: readonly ExtraCostField[] = BASE_EXTRA_COST_FIELDS,
  sheetName = 'HS total'
): Promise<void> {
  const taxCodes = unionTaxCodes(declaration.articles);
  // Rubriques from the Liquidation's RECAPITULATION table that never appear
  // on any individual article (e.g. REDV.INF., RI SEGMA, REMISES CREDIT) —
  // same declaration-wide-only taxes Global shows via montant × Prorata.
  // Included here too so every tax visible in Global also appears in HS
  // total.
  const extraOrdonnancementTaxes = declaration.ordonnancementTaxes.filter(
    (tax) => !taxCodes.includes(tax.code)
  );
  const extraCodes = extraOrdonnancementTaxes.map((tax) => tax.code);
  const allTaxCodes = [...taxCodes, ...extraCodes];
  const allTaxDesignations = [
    ...taxCodes.map(taxCodeDesignation),
    ...extraOrdonnancementTaxes.map((tax) => tax.designation),
  ];
  const declarationValeurDeclareeTotal = declaration.articles.reduce(
    (sum, article) => sum + article.valeurDeclaree,
    0
  );
  // One "<abbreviation> Total" column per tax (montant x pieces for the
  // row), right after HS CODE — same order as allTaxCodes, so column N here
  // always corresponds to allTaxCodes[N].
  const taxTotalColumns = allTaxCodes.map((_, i) => BASE_COLUMN_COUNT + 1 + i);
  // Somme DD HT / DD unitaire HT sit right after the "<tax> Total" columns,
  // immediately before Somme DD TTC / DD unitaire TTC — the individual tax
  // montant columns that used to sit right after HS CODE (DTS IMPORT
  // NORMAL, TVA IMPORT AUTRE PDS, etc.) are intentionally not shown on this
  // sheet; they're still computed internally (see the row loop below) to
  // feed Somme DD HT/TTC and the "<tax> Total" columns, just not displayed.
  const sommeDdColumn = BASE_COLUMN_COUNT + allTaxCodes.length + 1;
  const ddUnitaireColumn = sommeDdColumn + 1;
  // Somme DD TTC / DD unitaire TTC close out the sheet — the row's taxes
  // summed (montant x pieces, i.e. the sum of the "<tax> Total" columns)
  // and that sum spread per piece, same relationship as the HT pair.
  const sommeDdTtcColumn = ddUnitaireColumn + 1;
  const ddUnitaireTtcColumn = sommeDdTtcColumn + 1;
  // PRORATA closes out the sheet — the matched Global article's own Prorata
  // value (see firstUnitTaxes), not derived from this row's own quantities.
  const prorataColumn = ddUnitaireTtcColumn + 1;
  // One column per manually-entered extra cost field, right after PRORATA,
  // same order as extraCostFields.
  const extraCostColumns = extraCostFields.map((_, i) => prorataColumn + 1 + i);
  const columnCount = extraCostColumns[extraCostColumns.length - 1];
  const columnGroups: ColumnGroup[] = [
    ...BASE_COLUMN_GROUPS,
    // Somme DD HT / DD unitaire HT get their own green; the "<tax> Total"
    // columns get their own rose color; the closing Somme DD TTC / DD
    // unitaire TTC pair gets a second, distinct shade of green from the HT
    // pair, so all three read as visibly distinct from each other; PRORATA
    // and the extra cost columns get the same emerald "value" color Global
    // uses for its own Prorata column.
    { kind: 'taxTotal' as const, from: BASE_COLUMN_COUNT + 1, to: sommeDdColumn - 1 },
    { kind: 'sumHt' as const, from: sommeDdColumn, to: ddUnitaireColumn },
    { kind: 'sumTtc' as const, from: sommeDdTtcColumn, to: ddUnitaireTtcColumn },
    { kind: 'value' as const, from: prorataColumn, to: columnCount },
  ];

  const firstUnitTaxesByOrigin = firstUnitTaxesByHsAndOrigin(
    declaration,
    extraOrdonnancementTaxes,
    declarationValeurDeclareeTotal
  );
  const firstUnitTaxesByPrefix = firstUnitTaxesByHsPrefix(
    declaration,
    extraOrdonnancementTaxes,
    declarationValeurDeclareeTotal
  );
  const ambiguousHsPrefixes = hsPrefixesWithMultipleOrigins(declaration);
  const prorataSumsByOrigin = hsAndOriginProrataSums(declaration, declarationValeurDeclareeTotal);
  const prorataSumsByPrefix = hsPrefixProrataSums(declaration, declarationValeurDeclareeTotal);
  const groupTotalsByOrigin = quantiteAndValeurTotalsByHsAndOrigin(declaration);
  // Somme DD and the "<tax> Total" columns are zero-padded to at least 4
  // digits before the decimal separator, with no thousands separator —
  // requested so every value in these columns lines up the same width (e.g.
  // "0000,00" rather than "0,00"), instead of the shared money format.
  const zeroPaddedColumns = new Set<number>([sommeDdColumn, ...taxTotalColumns, sommeDdTtcColumn]);

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
    // Wide enough to fit the full "<tax designation> Total" header text
    // (designations vary a lot in length), not the fixed 18 used elsewhere.
    ...allTaxCodes.map((code, i) => ({
      key: `${code}_total`,
      width: Math.max(18, `${allTaxDesignations[i]} Total`.length + 2),
    })),
    { key: 'sommeDd', width: 18 },
    { key: 'ddUnitaire', width: 18 },
    { key: 'sommeDdTtc', width: 18 },
    { key: 'ddUnitaireTtc', width: 18 },
    { key: 'prorata', width: 18 },
    ...extraCostFields.map((field) => ({ key: field.key, width: Math.max(18, field.label.length + 2) })),
  ];

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

  const headerRow = sheet.addRow([
    'item',
    'DESCRIPTION',
    'color',
    'pieces',
    'unit',
    'total',
    'origin',
    'HS CODE',
    ...allTaxDesignations.map((designation) => `${designation} Total`),
    'Somme DD HT',
    'DD unitaire HT',
    'Somme DD TTC',
    'DD unitaire TTC',
    'PRORATA',
    ...extraCostFields.map((field) => field.label),
  ]);
  styleHeaderRowGrouped(headerRow, columnCount, columnGroups);

  // Adds one HS-total row: `identity` is the item/description/color/pieces/
  // unit/total/origin/hsCode set (either a merged group's combined totals or
  // an individual unmatched packing-list row's own), `taxes` are the group's
  // per-unit tax montants (undefined for an unmatched row), and `prorata` is
  // the already-resolved PRORATA fraction to show/spread extra costs by.
  function addHsTotalRow(
    identity: {
      item: string;
      description: string;
      color: string;
      pieces: number;
      unit: number;
      total: number;
      origin: string;
      hsCode: string;
    },
    taxes: Map<string, number> | undefined,
    prorata: number,
    index: number
  ): void {
    const rowValues: Record<string, string | number> = { ...identity };
    // Not displayed as its own columns (see the comment on sommeDdColumn
    // above), but still needed per-code to build Somme DD HT/TTC and the
    // "<tax> Total" columns below.
    const taxMontants = new Map<string, number>();
    for (const code of allTaxCodes) {
      taxMontants.set(code, taxes?.get(code) ?? 0);
    }
    // "<abbreviation> Total" columns — each tax's per-unit montant spread
    // across this row's piece count, the same montant x pieces pattern as
    // the "total" column (unit price x pieces).
    let sommeDdTtc = 0;
    for (const code of allTaxCodes) {
      const taxTotal = taxMontants.get(code)! * identity.pieces;
      rowValues[`${code}_total`] = taxTotal;
      sommeDdTtc += taxTotal;
    }
    // "Somme DD TTC" / "DD unitaire TTC" — the row's "<tax> Total" columns
    // summed, and that sum spread per piece.
    rowValues.sommeDdTtc = sommeDdTtc;
    rowValues.ddUnitaireTtc = identity.pieces > 0 ? sommeDdTtc / identity.pieces : 0;
    // "Somme DD HT" — Somme DD TTC minus the TVA IMPORT AUTRE PDS Total
    // column, i.e. every tax's total except VAT. "DD unitaire" spreads that
    // over this row's piece count, same per-unit pattern as Global's Valeur
    // Déclarée / Unité.
    const vatTotal = (taxMontants.get(VAT_TAX_CODE) ?? 0) * identity.pieces;
    const sommeDd = sommeDdTtc - vatTotal;
    rowValues.sommeDd = sommeDd;
    rowValues.ddUnitaire = identity.pieces > 0 ? sommeDd / identity.pieces : 0;
    rowValues.prorata = prorata;
    // Each extra cost field's shipment-wide total spread across this row by
    // its matched PRORATA share, same "total x prorata" split Global uses
    // for its declaration-wide-only taxes (see firstUnitTaxes above).
    for (const field of extraCostFields) {
      rowValues[field.key] = (extraCosts[field.key] ?? 0) * prorata;
    }

    const excelRow = sheet.addRow(rowValues);
    styleDataRow(
      excelRow,
      columnCount,
      index,
      new Set([
        UNIT_COLUMN,
        TOTAL_COLUMN,
        sommeDdColumn,
        ...taxTotalColumns,
        sommeDdTtcColumn,
        ...extraCostColumns,
      ]),
      new Set([prorataColumn])
    );
    // DD unitaire keeps 6 decimal digits rather than the shared 2-decimal
    // money format — it's a per-piece fraction, not a currency total. The
    // integer part is zero-padded to at least 4 digits, same as the tax
    // columns, but the 6 decimal digits after the comma are left untouched.
    // No thousands separator (plain "0000.000000"), since a grouping comma
    // next to 6 decimals read as a formatting glitch. Cloning the style
    // (rather than mutating cell.numFmt directly) matters here: styleDataRow
    // hands out shared style objects by reference, so mutating one cell's
    // numFmt in place would silently change every other cell using that
    // same banded/plain style throughout the sheet.
    const ddUnitaireCell = excelRow.getCell(ddUnitaireColumn);
    ddUnitaireCell.style = { ...ddUnitaireCell.style, numFmt: '0000.000000' };
    const ddUnitaireTtcCell = excelRow.getCell(ddUnitaireTtcColumn);
    ddUnitaireTtcCell.style = { ...ddUnitaireTtcCell.style, numFmt: '0000.000000' };
    // Pieces is a plain whole count — no thousands separator, no decimals.
    const piecesCell = excelRow.getCell(PIECES_COLUMN);
    piecesCell.style = { ...piecesCell.style, numFmt: '0' };
    // DESCRIPTION is free-form product text — left-aligned like a normal
    // text cell, not centered like the rest of the sheet's columns.
    const descriptionCell = excelRow.getCell(DESCRIPTION_COLUMN);
    descriptionCell.style = {
      ...descriptionCell.style,
      alignment: { vertical: 'middle', horizontal: 'left' },
    };
    for (const column of zeroPaddedColumns) {
      const cell = excelRow.getCell(column);
      cell.style = { ...cell.style, numFmt: '0000.00' };
    }
    // Somme DD HT/DD unitaire HT and Somme DD TTC/DD unitaire TTC each
    // carry their own green tint down every data row too, not just the
    // header — applied last so it isn't overwritten by the numFmt
    // overrides above.
    styleColumnsFill(excelRow, [sommeDdColumn, ddUnitaireColumn], SUM_HT_ROW_FILL_ARGB);
    styleColumnsFill(excelRow, [sommeDdTtcColumn, ddUnitaireTtcColumn], SUM_TTC_ROW_FILL_ARGB);
  }

  // Sums every matched packing-list row's own pieces/total into its
  // declaration HS+origin group, instead of a separate HS-total row per
  // packing-list row — a group with, say, several color/variant lines used
  // to show each one individually, all repeating that same group's summed
  // Prorata. Grouped by the matched article's own canonical HS+origin key
  // (not the row's own raw hsCode/origin spelling), so packing-list rows
  // whose origin spelling varies slightly still merge into the same group.
  interface GroupAggregate {
    pieces: number;
    total: number;
    matchedData: FirstUnitData;
  }
  const matchedByGroup = new Map<string, GroupAggregate>();
  const unmatchedRows: PackingListRow[] = [];
  for (const row of rows) {
    const rowHsPrefix = hsCodePrefix(row.hsCode);
    // Only prefixes with more than one country of origin in the declaration
    // need the origin-qualified match; everything else matches by HS prefix
    // alone (the pre-existing, reliable behavior), with the origin-qualified
    // lookup as a fallback if the row's origin spelling doesn't line up with
    // any of that prefix's known origins.
    const matchedData = ambiguousHsPrefixes.has(rowHsPrefix)
      ? (firstUnitTaxesByOrigin.get(hsAndOriginKey(row.hsCode, row.origin)) ??
        firstUnitTaxesByPrefix.get(rowHsPrefix))
      : firstUnitTaxesByPrefix.get(rowHsPrefix);
    if (!matchedData) {
      unmatchedRows.push(row);
      continue;
    }
    const key = hsAndOriginKey(matchedData.hsCode, matchedData.pays);
    const aggregate = matchedByGroup.get(key);
    if (aggregate) {
      aggregate.pieces += row.pieces;
      aggregate.total += row.total;
    } else {
      matchedByGroup.set(key, { pieces: row.pieces, total: row.total, matchedData });
    }
  }

  let index = 0;
  // One row per matched HS+origin group, in the order each group was first
  // encountered in the packing list.
  for (const [key, aggregate] of matchedByGroup) {
    // The group's already-summed Prorata (see hsAndOriginProrataSums) —
    // every packing-list row in this group contributes the same group total,
    // not its own individual share. `key` is built the same
    // hsAndOriginKey(hsCode, pays) way hsAndOriginProrataSums itself keys by,
    // and always corresponds to a real declaration article, so this lookup
    // is always defined.
    const prorata = Math.round((prorataSumsByOrigin.get(key) ?? 0) * 10000) / 10000;
    addHsTotalRow(
      {
        // No single packing-list "item"/"color" applies once several
        // packing-list rows (e.g. one per color/variant) are merged into
        // this one group row — the declaration's own product name is the
        // identity that still makes sense here.
        item: '',
        description: aggregate.matchedData.nomArticle,
        color: '',
        pieces: aggregate.pieces,
        unit: aggregate.pieces > 0 ? aggregate.total / aggregate.pieces : 0,
        total: aggregate.total,
        origin: aggregate.matchedData.pays,
        hsCode: aggregate.matchedData.hsCode,
      },
      aggregate.matchedData.taxes,
      prorata,
      index
    );
    index += 1;
  }

  // Packing-list rows that matched no declaration group at all (wrong/
  // unknown HS code, or a real origin that doesn't correspond to any
  // declaration article for that HS position) still show up individually,
  // under their own raw HS code/origin, so it's still visible what didn't
  // match and why.
  for (const row of unmatchedRows) {
    addHsTotalRow(
      {
        item: row.item,
        description: row.description,
        color: row.color,
        pieces: row.pieces,
        unit: row.unit,
        total: row.total,
        origin: row.origin,
        hsCode: row.hsCode,
      },
      undefined,
      0,
      index
    );
    index += 1;
  }

  // A declaration group (HS+origin) the packing list has no line for at
  // all — different from an unmatched packing-list row, which at least
  // shows up (under its own HS code) even without a match. Without this, a
  // product present in Articles/Global but missing from the supplier's
  // packing list simply never appears anywhere in HS total, with no
  // indication why (reported: an HSC visible in Global with no row at all
  // in HS total). One synthetic row per missing group stands in for the
  // whole group, sized from its own combined quantite/valeur declarée, with
  // the same summed-Prorata-share-of-the-declaration PRORATA a matched
  // group's row would show.
  for (const [key, data] of firstUnitTaxesByOrigin) {
    if (matchedByGroup.has(key)) continue;
    const totals = groupTotalsByOrigin.get(key) ?? { quantite: 0, valeur: 0 };
    const prorata = Math.round((prorataSumsByOrigin.get(key) ?? 0) * 10000) / 10000;
    addHsTotalRow(
      {
        item: '',
        description: `${data.nomArticle} (absent de la packing list)`,
        color: '',
        pieces: totals.quantite,
        unit: totals.quantite > 0 ? totals.valeur / totals.quantite : 0,
        total: totals.valeur,
        origin: data.pays,
        hsCode: data.hsCode,
      },
      data.taxes,
      prorata,
      index
    );
    index += 1;
  }
}

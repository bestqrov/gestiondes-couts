import ExcelJS from 'exceljs';
import type { PackingListRow } from '../parser/packingList/packingListParser.js';
import type { Declaration } from '../domain/types.js';
import { unionTaxCodes, taxCodeDesignation, hsAndOriginKey } from './unitLevelTaxHelpers.js';
import { createPackingListMatcher } from './packingListMatcher.js';
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

// Adds the "HS total" sheet — one row per declaration HS+origin group,
// summing every uploaded packing-list row matched to it (pieces/total/
// PRORATA all add up across however many packing-list lines, e.g. color
// variants, share that group), plus one row per packing-list row that
// matched no group at all. Declaration-driven grouping, not a raw copy of
// the packing list's own rows.
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
  // PRORATA closes out the sheet — the matched article's own per-unit
  // Prorata (see firstUnitTaxes) × this row's pieces (see the row loop
  // below), so PRORATA sums to ~100% across every row sharing the same
  // pays + HS code, in proportion to how many physical units each row
  // stands for.
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

  const matcher = createPackingListMatcher(declaration);
  const { firstUnitTaxesByOrigin, resolveRowMatch } = matcher;
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

  // Builds and styles one HS-total data row — shared by real packing-list
  // rows and by the synthetic rows added below for a declaration group the
  // packing list has no line for at all (see coveredGroupKeys).
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
    matchedTaxes: Map<string, number> | undefined,
    prorata: number,
    index: number
  ): void {
    const rowValues: Record<string, string | number> = { ...identity };
    // Not displayed as its own columns (see the comment on sommeDdColumn
    // above), but still needed per-code to build Somme DD HT/TTC and the
    // "<tax> Total" columns below.
    const taxMontants = new Map<string, number>();
    for (const code of allTaxCodes) {
      taxMontants.set(code, matchedTaxes?.get(code) ?? 0);
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
  // declaration HS+origin group, instead of keeping one HS-total row per
  // packing-list row — a group with, say, 71 packing-list lines (one per
  // color/variant) used to show its per-unit Prorata (e.g. 1/71 =
  // 1.4085%) repeated 71 times; it's now one row per group, with PRORATA
  // the sum of those repeats (up to 100%, or less if the packing list
  // doesn't fully itemize the group — see the "À vérifier" report for
  // exactly which units are missing and why).
  const matchedByGroup = new Map<string, { pieces: number; total: number }>();
  const unmatchedRows: PackingListRow[] = [];
  for (const row of rows) {
    const match = resolveRowMatch(row);
    if (!match) {
      unmatchedRows.push(row);
      continue;
    }
    const key = hsAndOriginKey(match.hsCode, match.pays);
    const aggregate = matchedByGroup.get(key) ?? { pieces: 0, total: 0 };
    aggregate.pieces += row.pieces;
    aggregate.total += row.total;
    matchedByGroup.set(key, aggregate);
  }

  let index = 0;
  // One row per declaration HS+origin group — every group gets a row
  // regardless of how many (if any) packing-list rows matched it, so a
  // product present in Articles/Global but missing from the packing list
  // still shows up here (with 0 pieces/PRORATA) instead of silently never
  // appearing at all.
  for (const [key, data] of firstUnitTaxesByOrigin) {
    const aggregate = matchedByGroup.get(key) ?? { pieces: 0, total: 0 };
    // PRORATA — the matched article's own per-unit Prorata (Global's Valeur
    // Déclarée / unit, over its HS+origin group's combined Valeur Déclarée)
    // × this group's total matched pieces, i.e. the sum of every matching
    // packing-list row's own share. Rounded to match the PRORATA column's
    // 0.00% display (4 decimal places of the raw fraction) so the
    // extra-cost columns below, which multiply by this same value, derive
    // from what's actually shown, not a longer float tail hidden by
    // display-only formatting.
    const prorata = Math.round(data.prorata * aggregate.pieces * 10000) / 10000;
    addHsTotalRow(
      {
        // No single packing-list "item"/"color" applies once several
        // packing-list rows (e.g. one per color/variant) are merged into
        // this one group row — the declaration's own product name is the
        // identity that still makes sense here.
        item: '',
        description: data.nomArticle,
        color: '',
        pieces: aggregate.pieces,
        unit: aggregate.pieces > 0 ? aggregate.total / aggregate.pieces : 0,
        total: aggregate.total,
        origin: data.pays,
        hsCode: data.hsCode,
      },
      data.taxes,
      prorata,
      index
    );
    index += 1;
  }

  // Packing-list rows that matched no declaration group at all (wrong/
  // unknown HS code, or a real origin that doesn't correspond to any
  // declaration article for that HS position) still show up individually,
  // under their own raw HS code/origin, so it's still visible what didn't
  // match and why — same rows the "À vérifier" report lists separately.
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
}

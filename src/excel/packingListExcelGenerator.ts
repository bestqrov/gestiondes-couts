import ExcelJS from 'exceljs';
import type { PackingListRow } from '../parser/packingList/packingListParser.js';
import type { Declaration } from '../domain/types.js';
import { normalizeCountryName } from '../domain/countryNames.js';
import { unionTaxCodes, taxCodeDesignation } from './unitLevelTaxHelpers.js';
import {
  styleDataRow,
  styleHeaderRowGrouped,
  addSheetTitleRows,
  resolveBrandArgb,
  resolveBrandDarkArgb,
  resolveCompanyName,
  resolveDocumentTitle,
  type BrandingInfo,
  type ColumnGroup,
} from './excelStyling.js';

const BASE_COLUMN_COUNT = 8;
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

// Sums every declaration article's tax montants, grouped by HS code prefix +
// country of origin — multiple articles (e.g. one per color/variant) can
// share the same HS position and origin, the same way multiple packing-list
// rows do, so the group total is what "the tax value for this HS code and
// origin" means.
function sumTaxesByHsAndOrigin(declaration: Declaration): Map<string, Map<string, number>> {
  const totals = new Map<string, Map<string, number>>();
  for (const article of declaration.articles) {
    const key = hsAndOriginKey(article.hsCode, article.pays);
    const perCode = totals.get(key) ?? new Map<string, number>();
    for (const tax of article.taxes) {
      perCode.set(tax.code, (perCode.get(tax.code) ?? 0) + tax.montant);
    }
    totals.set(key, perCode);
  }
  return totals;
}

// Same as sumTaxesByHsAndOrigin, but grouped by HS code prefix alone — the
// fallback used for the (overwhelmingly common) case where an HS prefix has
// only one country of origin in the whole declaration, so origin spelling
// (packing list vs DUM/Liquidation, potentially a country name outside
// countryNames.ts's known aliases) can't cause a false non-match.
function sumTaxesByHsPrefix(declaration: Declaration): Map<string, Map<string, number>> {
  const totals = new Map<string, Map<string, number>>();
  for (const article of declaration.articles) {
    const prefix = hsCodePrefix(article.hsCode);
    const perCode = totals.get(prefix) ?? new Map<string, number>();
    for (const tax of article.taxes) {
      perCode.set(tax.code, (perCode.get(tax.code) ?? 0) + tax.montant);
    }
    totals.set(prefix, perCode);
  }
  return totals;
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

// Adds the "HS total" sheet — a direct, unaggregated copy of the uploaded
// packing list's rows, in their original order, independent of the
// Liquidation/DUM merge/validation pipeline (`declaration` is only used for
// the letterhead's document-title line, same as every other sheet).
export async function addPackingListSheet(
  workbook: ExcelJS.Workbook,
  rows: PackingListRow[],
  declaration: Declaration,
  branding: BrandingInfo,
  generatedAt: Date,
  sheetName = 'HS total'
): Promise<void> {
  const taxCodes = unionTaxCodes(declaration.articles);
  // Somme DD / DD unitaire sit right after the tax code columns — derived
  // sums of those tax values, so grouped and colored with them.
  const sommeDdColumn = BASE_COLUMN_COUNT + taxCodes.length + 1;
  const ddUnitaireColumn = sommeDdColumn + 1;
  const columnCount = ddUnitaireColumn;
  const taxColumns = new Set<number>(taxCodes.map((_, i) => BASE_COLUMN_COUNT + 1 + i));
  const columnGroups: ColumnGroup[] = [
    ...BASE_COLUMN_GROUPS,
    { kind: 'tax' as const, from: BASE_COLUMN_COUNT + 1, to: columnCount },
  ];

  const taxTotalsByHsAndOrigin = sumTaxesByHsAndOrigin(declaration);
  const taxTotalsByHsPrefix = sumTaxesByHsPrefix(declaration);
  const ambiguousHsPrefixes = hsPrefixesWithMultipleOrigins(declaration);
  // The tax columns (DTS IMPORT NORMAL, TVA IMPORT AUTRE PDS, etc.) and
  // Somme DD are zero-padded to at least 4 digits before the decimal
  // separator, with no thousands separator — requested so every value in
  // these columns lines up the same width (e.g. "0000,00" rather than
  // "0,00"), instead of the shared money format.
  const zeroPaddedColumns = new Set<number>([...taxColumns, sommeDdColumn]);

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
    ...taxCodes.map((code) => ({ key: code, width: 24 })),
    { key: 'sommeDd', width: 18 },
    { key: 'ddUnitaire', width: 18 },
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
    ...taxCodes.map(taxCodeDesignation),
    'Somme DD',
    'DD unitaire',
  ]);
  styleHeaderRowGrouped(headerRow, columnCount, columnGroups);

  rows.forEach((row, index) => {
    const rowHsPrefix = hsCodePrefix(row.hsCode);
    // Only prefixes with more than one country of origin in the declaration
    // need the origin-qualified match; everything else matches by HS prefix
    // alone (the pre-existing, reliable behavior), with the origin-qualified
    // lookup as a fallback if the row's origin spelling doesn't line up with
    // any of that prefix's known origins.
    const matchedTaxTotals = ambiguousHsPrefixes.has(rowHsPrefix)
      ? (taxTotalsByHsAndOrigin.get(hsAndOriginKey(row.hsCode, row.origin)) ??
        taxTotalsByHsPrefix.get(rowHsPrefix))
      : taxTotalsByHsPrefix.get(rowHsPrefix);
    const rowValues: Record<string, string | number> = {
      item: row.item,
      description: row.description,
      color: row.color,
      pieces: row.pieces,
      unit: row.unit,
      total: row.total,
      origin: row.origin,
      // The full HS code is shown here; only the 6-digit prefix is used
      // above to match against the declaration's tax totals.
      hsCode: row.hsCode,
    };
    let sommeDd = 0;
    for (const code of taxCodes) {
      const montant = matchedTaxTotals?.get(code) ?? 0;
      rowValues[code] = montant;
      sommeDd += montant;
    }
    // "DD unitaire" — Somme DD spread over this row's piece count, the same
    // per-unit pattern as Global's Valeur Déclarée / Unité.
    rowValues.sommeDd = sommeDd;
    rowValues.ddUnitaire = row.pieces > 0 ? sommeDd / row.pieces : 0;

    const excelRow = sheet.addRow(rowValues);
    styleDataRow(
      excelRow,
      columnCount,
      index,
      new Set([UNIT_COLUMN, TOTAL_COLUMN, ...taxColumns, sommeDdColumn])
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
    // Pieces is a plain whole count — no thousands separator, no decimals.
    const piecesCell = excelRow.getCell(PIECES_COLUMN);
    piecesCell.style = { ...piecesCell.style, numFmt: '0' };
    for (const column of zeroPaddedColumns) {
      const cell = excelRow.getCell(column);
      cell.style = { ...cell.style, numFmt: '0000.00' };
    }
  });
}

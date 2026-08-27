import type { PackingListRow } from '../parser/packingList/packingListParser.js';
import type { Declaration } from '../domain/types.js';
import { normalizeCountryName } from '../domain/countryNames.js';
import {
  allocateTaxAcrossUnits,
  hsAndOriginKey,
  valeurDeclareeTotalsByHsAndOrigin,
} from './unitLevelTaxHelpers.js';

// Only the first 6 digits (the HS "position") need to agree to consider a
// packing-list row and a declaration article the same product — matches the
// HS_CODE_COMPARISON_LENGTH convention already used to reconcile the
// Liquidation and DUM's HS codes in declarationMerger.ts. Product names
// differ too much between the supplier's packing list and the customs
// documents' OCR'd text to match reliably by name.
const HS_CODE_MATCH_LENGTH = 6;

export function hsCodePrefix(code: string): string {
  // Strips any non-digit characters (dots, spaces, dashes) before taking the
  // prefix — see unitLevelTaxHelpers.ts's matching hsCodePrefix.
  return code.replace(/\D/g, '').slice(0, HS_CODE_MATCH_LENGTH);
}

// The per-unit tax montants of an article's very first physical unit (the
// same allocation Global's row 1 for this article shows) — not the article's
// full total. Also folds in the declaration-wide-only ordonnancement taxes
// (e.g. REDV.INF., RI SEGMA, REMISES CREDIT) using the same montant ×
// Prorata treatment Global gives them. `prorata` is Global's own per-unit
// Prorata for this article, `hsCode`/`pays` are its own (Global's HSC/
// origin), for display in place of a matched packing-list row's own,
// shorter/less precise values.
export interface FirstUnitData {
  taxes: Map<string, number>;
  prorata: number;
  hsCode: string;
  pays: string;
  nomArticle: string;
}

function firstUnitTaxes(
  article: Declaration['articles'][number],
  extraOrdonnancementTaxes: Declaration['ordonnancementTaxes'],
  groupTotal: number,
  declarationTotal: number
): FirstUnitData {
  const quantite = Math.round(article.quantite);
  const perCode = new Map<string, number>();
  for (const tax of article.taxes) {
    perCode.set(tax.code, quantite > 0 ? allocateTaxAcrossUnits(tax.montant, quantite)[0] : 0);
  }
  const prorata = quantite > 0 && groupTotal > 0 ? article.valeurDeclaree / quantite / groupTotal : 0;
  const declarationShare =
    quantite > 0 && declarationTotal > 0 ? article.valeurDeclaree / quantite / declarationTotal : 0;
  if (declarationShare > 0) {
    for (const tax of extraOrdonnancementTaxes) {
      perCode.set(tax.code, tax.montant * declarationShare);
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
  valeurDeclareeTotalsByGroup: Map<string, number>,
  declarationTotal: number
): Map<string, FirstUnitData> {
  const result = new Map<string, FirstUnitData>();
  for (const article of declaration.articles) {
    const key = hsAndOriginKey(article.hsCode, article.pays);
    if (result.has(key)) continue;
    const groupTotal = valeurDeclareeTotalsByGroup.get(key) ?? 0;
    result.set(key, firstUnitTaxes(article, extraOrdonnancementTaxes, groupTotal, declarationTotal));
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
  valeurDeclareeTotalsByGroup: Map<string, number>,
  declarationTotal: number
): Map<string, FirstUnitData> {
  const result = new Map<string, FirstUnitData>();
  for (const article of declaration.articles) {
    const prefix = hsCodePrefix(article.hsCode);
    if (result.has(prefix)) continue;
    const groupTotal = valeurDeclareeTotalsByGroup.get(hsAndOriginKey(article.hsCode, article.pays)) ?? 0;
    result.set(prefix, firstUnitTaxes(article, extraOrdonnancementTaxes, groupTotal, declarationTotal));
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

export interface PackingListMatcher {
  resolveRowMatch(row: PackingListRow): FirstUnitData | undefined;
  // Sum of Valeur Déclarée / physical units (quantite), per HS+origin group
  // (hsAndOriginKey) — shared "HS total" and the verification report's own
  // coverage-gap computation both need these.
  valeurDeclareeTotalsByGroup: Map<string, number>;
  quantiteTotalsByGroup: Map<string, number>;
  firstUnitTaxesByOrigin: Map<string, FirstUnitData>;
}

// Builds everything needed to match packing-list rows to declaration
// articles by HS code prefix + country of origin — shared by "HS total"
// (packingListExcelGenerator.ts) and the data-verification report
// (verificationReportExcelGenerator.ts), so both agree on exactly which
// rows/groups do and don't match.
export function createPackingListMatcher(declaration: Declaration): PackingListMatcher {
  const taxCodes = new Set(declaration.articles.flatMap((article) => article.taxes.map((tax) => tax.code)));
  const extraOrdonnancementTaxes = declaration.ordonnancementTaxes.filter((tax) => !taxCodes.has(tax.code));
  const declarationValeurDeclareeTotal = declaration.articles.reduce(
    (sum, article) => sum + article.valeurDeclaree,
    0
  );
  const valeurDeclareeTotalsByGroup = valeurDeclareeTotalsByHsAndOrigin(declaration.articles);
  const quantiteTotalsByGroup = new Map<string, number>();
  for (const article of declaration.articles) {
    const key = hsAndOriginKey(article.hsCode, article.pays);
    quantiteTotalsByGroup.set(key, (quantiteTotalsByGroup.get(key) ?? 0) + Math.round(article.quantite));
  }
  const firstUnitTaxesByOrigin = firstUnitTaxesByHsAndOrigin(
    declaration,
    extraOrdonnancementTaxes,
    valeurDeclareeTotalsByGroup,
    declarationValeurDeclareeTotal
  );
  const firstUnitTaxesByPrefix = firstUnitTaxesByHsPrefix(
    declaration,
    extraOrdonnancementTaxes,
    valeurDeclareeTotalsByGroup,
    declarationValeurDeclareeTotal
  );
  const ambiguousHsPrefixes = hsPrefixesWithMultipleOrigins(declaration);
  // Every distinct (normalized) country of origin present anywhere in the
  // declaration — used to tell apart two different reasons an HS-prefix
  // match's own origin might not equal a row's: an unrecognized/misspelled
  // name (safe to ignore — see hsPrefixesWithMultipleOrigins's regression
  // fix), vs. a real, known country the row is genuinely reporting, that
  // just isn't this HS position's (e.g. a row honestly marked "INDE" for an
  // HS prefix the declaration only has a "CHINE" article for) — the latter
  // must not be silently pooled into the wrong article's taxes/values.
  const declarationOrigins = new Set(declaration.articles.map((article) => normalizeCountryName(article.pays)));
  function originSafeToIgnore(rowOrigin: string, matchedPays: string): boolean {
    const normalizedRowOrigin = normalizeCountryName(rowOrigin);
    return normalizedRowOrigin === normalizeCountryName(matchedPays) || !declarationOrigins.has(normalizedRowOrigin);
  }
  // Resolves a packing-list row to its matched declaration data — an
  // HS+origin match when ambiguous, the bare HS prefix otherwise —
  // undefined when no declaration article matches this row at all, or when
  // its own origin is a real, different one from the matched article's
  // (see originSafeToIgnore).
  function resolveRowMatch(row: PackingListRow): FirstUnitData | undefined {
    const prefix = hsCodePrefix(row.hsCode);
    if (!ambiguousHsPrefixes.has(prefix)) {
      const data = firstUnitTaxesByPrefix.get(prefix);
      if (!data) return undefined;
      return originSafeToIgnore(row.origin, data.pays) ? data : undefined;
    }
    const originKey = hsAndOriginKey(row.hsCode, row.origin);
    const originData = firstUnitTaxesByOrigin.get(originKey);
    if (originData) return originData;
    const prefixData = firstUnitTaxesByPrefix.get(prefix);
    if (!prefixData) return undefined;
    return originSafeToIgnore(row.origin, prefixData.pays) ? prefixData : undefined;
  }

  return { resolveRowMatch, valeurDeclareeTotalsByGroup, quantiteTotalsByGroup, firstUnitTaxesByOrigin };
}

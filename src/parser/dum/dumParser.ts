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

export interface DumShipmentCost {
  devise: string;
  montantFacture: number;
  tauxChange: number;
  fret: number;
  assurance: number;
  valeurTotaleDeclaree: number;
}

export interface DumResult {
  creditEnlevementCode: string;
  articles: DumArticleResult[];
  shipmentCost?: DumShipmentCost;
}

// pdfjs-dist extracts a DUM page's text items in the PDF's internal content-
// stream order, not visual reading order — labels and their corresponding
// values end up scattered non-adjacently in the resulting string (confirmed
// against a real sample DUM PDF), so label-anchored regexes like
// "Code marchandises\s*:\s*(\d+)" don't work here. What IS reliable is that
// each article's DATA VALUES appear together, in a fixed relative order:
// HS code, valeur déclarée, poids net, AP/SP, unités complémentaires, pays
// (nom + code), then — after a short run of variable boilerplate
// (déclaration sommaire / colis-type text, which differs per article) — the
// product designation and its true piece quantity/unit/ordre number. This
// pattern is matched positionally instead of via labels.
//
// Some HS codes carry a "(PTI:OUI)" suffix directly after the code, with no
// whitespace before it — matched as an optional, non-capturing group so it
// doesn't break the match for those articles.
//
// The "unités complémentaires" field (right after poids net / AP-SP) is
// read but NOT used as the article's quantity — for most goods it happens
// to equal the piece count (unit "U"), but for some HS codes (e.g. perfume,
// classified by net weight) it's a genuinely different measure in "KG",
// which would both mismatch the Liquidation's piece quantity and fail to
// match the literal "U" this pattern used to require. The real piece
// quantity/unit is instead captured from the désignation line ("EAU DE
// TOILETTE 480.00 NB"), which reliably matches the Liquidation regardless
// of what the complementary-unit field measures.
//
// The désignation capture requires a run of 2+ spaces immediately before it
// (a genuine field-boundary marker in this jumbled text) so that preceding
// boilerplate words (e.g. "COLIS", "MARCHANDISES NON EMBALLEE") — which are
// also capitalized and adjacent — aren't swept into the product name.
const ARTICLE_PATTERN =
  /(\d{10})(?:\([^)]*\))?\s+(\d[\d\s.,]*?\d)\s+[\d.]+\s+(?:AP|SP)\s+[\d.]+\s+(?:U|KG)\s+([A-Z][A-Z]*)\s+([A-Z]{2})\b[\s\S]{0,120}?\s{2,}([A-Z][A-Z-]*(?:\s[A-Z][A-Z-]*)*)\s+([\d.,]+)\s+(NB|U|PAIRE)\s+(\d+)\b/g;

// See design spec §4 — this cluster of shipment-level values (currency,
// invoiced amount, exchange rate, freight, [a form field-number label,
// ignored], insurance, total declared value) appears together in the raw
// extracted text, terminated by a DD MM YYYY date used only as a reliable
// anchor. Optional: if not found, shipmentCost is simply omitted rather
// than treated as a parse failure — the app's core function doesn't depend
// on it.
//
// The currency alternation (EUR|MAD|USD|GBP) is derived from the single real
// sample this pattern was built against. A shipment invoiced in a currency
// outside this list won't match and falls through to the same "not found" /
// undefined result as a genuinely malformed document, with no distinct
// signal for which case occurred. Broaden this list if a shipment in another
// currency is encountered.
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

export function parseDum(text: string): DumResult {
  // Bounded to a short gap (\D{0,20}) rather than unbounded \D*, so that if OCR noise ever
  // separates the label from its value by more than a short run of non-digits, this fails to
  // match (and throws "not found" below) instead of silently skipping ahead to grab an
  // unrelated digit run as the code.
  const creditEnlevementCode = extractFirst(text, /Crédit d'enlèvement\D{0,20}(\d+)/);
  if (!creditEnlevementCode) {
    throw new Error("Crédit d'enlèvement code not found in DUM document");
  }

  // Note on failure mode: unlike the label-based approach this replaces,
  // ARTICLE_PATTERN either matches an article fully or not at all — there is
  // no per-field "missing X" error for a partially-matching article; a
  // malformed one is simply absent from the result. This risks silently
  // returning fewer articles than the document actually has. That risk is
  // caught downstream: DeclarationMerger cross-checks the DUM's article set
  // against the Liquidation's and throws if they don't match one-to-one, so
  // a silently-dropped DUM article surfaces as a merge error rather than
  // silently reaching Excel generation.
  const articles: DumArticleResult[] = [];

  for (const match of text.matchAll(ARTICLE_PATTERN)) {
    const [, hsCode, valeurRaw, paysNom, paysCode, nomArticleRaw, quantiteRaw, unite, ordreRaw] =
      match;

    articles.push({
      ordre: Number.parseInt(ordreRaw, 10),
      hsCode,
      nomArticle: nomArticleRaw.trim(),
      paysNom,
      paysCode,
      valeurDeclaree: parseFrenchNumber(valeurRaw),
      quantite: parseFrenchNumber(quantiteRaw),
      unite,
    });
  }

  if (articles.length === 0) {
    throw new Error('No articles found in DUM document');
  }

  return { creditEnlevementCode, articles, shipmentCost: extractShipmentCost(text) };
}
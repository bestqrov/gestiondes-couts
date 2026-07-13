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

export interface DumResult {
  creditEnlevementCode: string;
  articles: DumArticleResult[];
}

// pdfjs-dist extracts a DUM page's text items in the PDF's internal content-
// stream order, not visual reading order — labels and their corresponding
// values end up scattered non-adjacently in the resulting string (confirmed
// against a real sample DUM PDF), so label-anchored regexes like
// "Code marchandises\s*:\s*(\d+)" don't work here. What IS reliable is that
// each article's DATA VALUES appear together, in a fixed relative order:
// HS code, valeur déclarée, poids net, AP/SP, quantité, unité, pays (nom +
// code), then — after a short run of variable boilerplate (déclaration
// sommaire / colis-type text, which differs per article) — the product
// designation and its ordre number. This pattern is matched positionally
// instead of via labels.
//
// The désignation capture requires a run of 2+ spaces immediately before it
// (a genuine field-boundary marker in this jumbled text) so that preceding
// boilerplate words (e.g. "COLIS", "MARCHANDISES NON EMBALLEE") — which are
// also capitalized and adjacent — aren't swept into the product name.
const ARTICLE_PATTERN =
  /(\d{10})\s+(\d[\d\s.,]*?\d)\s+[\d.]+\s+(?:AP|SP)\s+([\d.]+)\s+(U)\s+([A-Z][A-Z]*)\s+([A-Z]{2})\b[\s\S]{0,120}?\s{2,}([A-Z][A-Z-]*(?:\s[A-Z][A-Z-]*)*)\s+[\d.,]+\s+(?:NB|U)\s+(\d+)\b/g;

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
    const [, hsCode, valeurRaw, quantiteRaw, unite, paysNom, paysCode, nomArticleRaw, ordreRaw] =
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

  return { creditEnlevementCode, articles };
}
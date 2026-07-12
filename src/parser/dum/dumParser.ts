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

const ARTICLE_BLOCK_PATTERN =
  /N°\s*d'ordre de l'art\.\s*:\s*(\d+)([\s\S]*?)(?=N°\s*d'ordre de l'art\.\s*:\s*\d+|Renseignements financiers|$)/g;

export function parseDum(text: string): DumResult {
  // Bounded to a short gap (\D{0,20}) rather than unbounded \D*, so that if OCR noise ever
  // separates the label from its value by more than a short run of non-digits, this fails to
  // match (and throws "not found" below) instead of silently skipping ahead to grab an
  // unrelated digit run as the code.
  const creditEnlevementCode = extractFirst(text, /Crédit d'enlèvement\D{0,20}(\d+)/);
  if (!creditEnlevementCode) {
    throw new Error("Crédit d'enlèvement code not found in DUM document");
  }

  const articles: DumArticleResult[] = [];

  for (const match of text.matchAll(ARTICLE_BLOCK_PATTERN)) {
    const ordre = Number.parseInt(match[1], 10);
    const block = match[2];

    const designationLine = extractFirst(block, /Colis et désignation des marchandises\s*:\s*(.+)/);
    const hsCode = extractFirst(block, /Code marchandises\s*:\s*(\d+)/);
    const valeurRaw = extractFirst(block, /Valeur déclarée\s*:\s*([\d\s.,]+)/);
    const quantiteMatch = block.match(/Unités complémentaires\s*:\s*([\d.,]+)\s*(\S+)/);
    const paysMatch = block.match(/Pays d'origine \(Nom et code\)\s*:\s*(\S+)\s+(\S+)/);

    if (!designationLine || !hsCode || !valeurRaw || !quantiteMatch || !paysMatch) {
      throw new Error(
        `DUM article ${ordre}: missing one of designation / HS code / valeur / quantite / pays`
      );
    }

    // The trailing "QTY UNIT" captured here (e.g. "354.00 NB") is intentionally discarded and
    // not cross-validated against the "Unités complémentaires" quantity/unit below (e.g.
    // "354.0 U") — the two use different unit vocabularies and are not reconciled, by design
    // for now.
    const designationMatch = designationLine.match(/^([A-Za-zÀ-ÿ\-\s]+?)\s+[\d.,]+\s+\S+$/);
    if (!designationMatch) {
      throw new Error(
        `DUM article ${ordre}: designation line does not match the expected "NAME QTY UNIT" format: "${designationLine}"`
      );
    }
    const nomArticle = designationMatch[1].trim();

    articles.push({
      ordre,
      hsCode,
      nomArticle,
      paysNom: paysMatch[1],
      paysCode: paysMatch[2],
      valeurDeclaree: parseFrenchNumber(valeurRaw),
      quantite: parseFrenchNumber(quantiteMatch[1]),
      unite: quantiteMatch[2],
    });
  }

  if (articles.length === 0) {
    throw new Error('No articles found in DUM document');
  }

  return { creditEnlevementCode, articles };
}

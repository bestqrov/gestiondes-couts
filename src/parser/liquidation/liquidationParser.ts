import type { TaxLine } from '../../domain/types.js';
import { extractFirst, parseFrenchNumber } from '../shared/text.js';

export interface LiquidationHeader {
  code: string;
  redevable: string;
  benNumero: string;
}

export interface LiquidationArticleResult {
  numero: number;
  hsCode: string;
  valeurDeclaree: number;
  quantite: number;
  unite: string;
  taxes: TaxLine[];
  totalArticle: number;
}

export interface LiquidationResult {
  header: LiquidationHeader;
  articles: LiquidationArticleResult[];
}

// The negative lookbehind (?<!TOTAL ) excludes "TOTAL ARTICLE :" lines, which also match the
// naive "ARTICLE\s*:\s*(\d+)" pattern and would otherwise be mistaken for a new article-block
// boundary, corrupting how blocks are split.
const ARTICLE_BLOCK_PATTERN =
  /(?<!TOTAL )ARTICLE\s*:\s*(\d+)([\s\S]*?)(?=(?<!TOTAL )ARTICLE\s*:\s*\d+|$)/g;

const TAX_ROW_PATTERN = /^!\s*\d{6}\s*!/;

function parseTaxRows(block: string): TaxLine[] {
  const taxes: TaxLine[] = [];
  for (const line of block.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed.startsWith('!')) continue;

    if (!TAX_ROW_PATTERN.test(trimmed)) {
      // The header row (e.g. "! TAXE ! ASSIETTE ! ...") is expected and skipped; any other
      // line starting with "!" that doesn't look like a valid tax row is malformed input
      // (e.g. OCR-garbled code) and must not be silently dropped.
      if (trimmed.includes('TAXE')) continue;
      throw new Error(`Malformed tax row (unrecognized format): "${trimmed}"`);
    }

    const parts = trimmed.split('!').map((part) => part.trim());
    if (parts.length !== 9) {
      throw new Error(
        `Malformed tax row (expected 9 "!"-delimited fields, got ${parts.length}): "${trimmed}"`
      );
    }
    const [, code, assietteRaw, tauxRaw, , , , montantRaw] = parts;
    taxes.push({
      code,
      assiette: parseFrenchNumber(assietteRaw),
      taux: parseFrenchNumber(tauxRaw),
      montant: parseFrenchNumber(montantRaw),
    });
  }
  return taxes;
}

export function parseLiquidation(text: string): LiquidationResult {
  const code = extractFirst(text, /CODE\s*:\s*(\d+)/);
  const redevable = extractFirst(text, /REDEVABLE\s*:\s*(.+)/);
  const benNumero = extractFirst(text, /B E N°\s*:\s*(\S+)/);

  if (!code || !redevable || !benNumero) {
    throw new Error('Liquidation header fields (CODE, REDEVABLE, B E N°) not found');
  }

  const articles: LiquidationArticleResult[] = [];

  for (const match of text.matchAll(ARTICLE_BLOCK_PATTERN)) {
    const numero = Number.parseInt(match[1], 10);
    const block = match[2];

    const hsCode = extractFirst(block, /NUMERO SH\s*:\s*(\d+)/);
    const valeurRaw = extractFirst(block, /VALEUR\s*:\s*([\d\s.,]+)/);
    const quantiteRaw = extractFirst(block, /QUANTITE\s*:\s*([\d.]+)/);
    const unite = extractFirst(block, /UNITE\s*:\s*(\S+)/);
    const totalRaw = extractFirst(block, /TOTAL ARTICLE\s*:\s*([\d\s.,]+)/);

    if (!hsCode || !valeurRaw || !quantiteRaw || !unite || !totalRaw) {
      throw new Error(
        `Article ${numero}: missing one of HS code / valeur / quantite / unite / total`
      );
    }

    const taxes = parseTaxRows(block);
    if (taxes.length === 0) {
      throw new Error(`Article ${numero}: no tax rows found`);
    }

    articles.push({
      numero,
      hsCode,
      valeurDeclaree: parseFrenchNumber(valeurRaw),
      quantite: parseFrenchNumber(quantiteRaw),
      unite,
      taxes,
      totalArticle: parseFrenchNumber(totalRaw),
    });
  }

  if (articles.length === 0) {
    throw new Error('No articles found in Liquidation document');
  }

  return { header: { code, redevable, benNumero }, articles };
}

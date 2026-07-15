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

// A genuine tax row has 8 "!" delimiters (9 fields). Real-world OCR frequently
// misreads a table's box-drawing border as a stray "!" at the start/end of
// unrelated label lines (e.g. "! QUANTITE : 354.000 UNITE : NOMBRE !", "!
// TOTAL ARTICLE : 5 511,00 |") — those aren't tax-row data and must be
// skipped as noise, not treated as malformed tax rows. A line with too few
// "!" characters to plausibly be a multi-column data row is assumed to be
// such noise; a line with enough "!" characters to look like a real attempted
// data row, but that doesn't parse cleanly, is still treated as malformed
// (likely-corrupted) tax data and must not be silently dropped.
const MIN_DELIMITERS_FOR_TAX_ROW_CANDIDATE = 3;

function countChar(text: string, char: string): number {
  return text.split(char).length - 1;
}

// The last article's block has no following "ARTICLE :" marker to stop at,
// so ARTICLE_BLOCK_PATTERN's lazy match otherwise runs all the way to the
// end of the document — swallowing whatever comes after it on later pages
// (e.g. the "RECAPITULATION" summary table, which has its own "!"-delimited
// rows and a 6-digit "RUBRIQUE" code that looks just like a tax row, or the
// "LISTE REDEVABLES SOLIDAIRES" table). Every genuine article block ends
// with its own "TOTAL ARTICLE :" line, so truncating there bounds the block
// to just that article regardless of what follows in the source document.
function truncateAtTotalArticleLine(block: string): string {
  const totalIndex = block.search(/TOTAL ARTICLE\s*:/);
  if (totalIndex === -1) return block;
  const newlineAfter = block.indexOf('\n', totalIndex);
  return newlineAfter === -1 ? block : block.slice(0, newlineAfter);
}

function parseTaxRows(block: string): TaxLine[] {
  const taxes: TaxLine[] = [];
  for (const line of block.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed.startsWith('!')) continue;

    if (!TAX_ROW_PATTERN.test(trimmed)) {
      // The header row (e.g. "! TAXE ! ASSIETTE ! ...") is expected and skipped.
      if (trimmed.includes('TAXE')) continue;
      // Not enough "!" delimiters to plausibly be a data row — OCR border noise on an
      // unrelated label line, not tax data. Skip rather than throw.
      if (countChar(trimmed, '!') < MIN_DELIMITERS_FOR_TAX_ROW_CANDIDATE) continue;
      // Looks like an attempted data row (enough delimiters) but doesn't match the
      // expected shape — genuinely malformed tax data and must not be silently dropped.
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
  // Stop the capture at the next label on the same line (e.g. OCR often merges
  // "REDEVABLE : X" and "CODE : Y" onto one line since they sit side-by-side
  // on the source document) or at a newline, whichever comes first.
  const redevable = extractFirst(text, /REDEVABLE\s*:\s*(.+?)(?=\s+CODE\s*:|\n|$)/);
  // Tolerate OCR spacing variance around "B E N°" (e.g. "BE N°", "B.E.N°").
  const benNumero = extractFirst(text, /B\s*\.?\s*E\s*\.?\s*N[°ºo]?\s*:\s*(\S+)/i);

  if (!code || !redevable || !benNumero) {
    throw new Error('Liquidation header fields (CODE, REDEVABLE, B E N°) not found');
  }

  const articles: LiquidationArticleResult[] = [];

  for (const match of text.matchAll(ARTICLE_BLOCK_PATTERN)) {
    const numero = Number.parseInt(match[1], 10);
    const block = truncateAtTotalArticleLine(match[2]);

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

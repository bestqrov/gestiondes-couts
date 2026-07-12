import type { Article, Declaration, TaxLine } from '../domain/types.js';
import type { LiquidationResult } from '../parser/liquidation/liquidationParser.js';
import type { DumResult } from '../parser/dum/dumParser.js';

const VALEUR_TOLERANCE = 0.5;
const QUANTITE_TOLERANCE = 0.5;

export class MergeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'MergeError';
  }
}

// Provisional matching rule based on the single real sample pair we have
// (Liquidation code "500001" -> DUM crédit d'enlèvement code "500001099"):
// the DUM code must be exactly the Liquidation code followed by exactly
// 3 additional digits. This is deliberately stricter than an unbounded
// prefix match so an unrelated/truncated code cannot accidentally collide.
// See docs/superpowers/specs/2026-07-11-customs-declaration-excel-design.md
// §8 "Open Items" — this exact matching rule is flagged there as something
// to finalize once more real sample pairs are available.
function isMatchingCreditEnlevementCode(liquidationCode: string, dumCode: string): boolean {
  return dumCode.length === liquidationCode.length + 3 && dumCode.startsWith(liquidationCode);
}

export function mergeDeclaration(liquidation: LiquidationResult, dum: DumResult): Declaration {
  if (!isMatchingCreditEnlevementCode(liquidation.header.code, dum.creditEnlevementCode)) {
    throw new MergeError(
      `Liquidation code "${liquidation.header.code}" does not match DUM crédit d'enlèvement code "${dum.creditEnlevementCode}"`
    );
  }

  const dumByOrdre = new Map(dum.articles.map((article) => [article.ordre, article]));

  const articles: Article[] = liquidation.articles.map((liqArticle) => {
    const dumArticle = dumByOrdre.get(liqArticle.numero);
    if (!dumArticle) {
      throw new MergeError(`Article ${liqArticle.numero} present in Liquidation but not found in DUM`);
    }

    if (dumArticle.hsCode !== liqArticle.hsCode) {
      throw new MergeError(
        `Article ${liqArticle.numero}: HS code mismatch (Liquidation "${liqArticle.hsCode}" vs DUM "${dumArticle.hsCode}")`
      );
    }

    if (Math.abs(dumArticle.valeurDeclaree - liqArticle.valeurDeclaree) > VALEUR_TOLERANCE) {
      throw new MergeError(
        `Article ${liqArticle.numero}: valeur déclarée mismatch (Liquidation ${liqArticle.valeurDeclaree} vs DUM ${dumArticle.valeurDeclaree})`
      );
    }

    if (Math.abs(dumArticle.quantite - liqArticle.quantite) > QUANTITE_TOLERANCE) {
      throw new MergeError(
        `Article ${liqArticle.numero}: quantité mismatch (Liquidation ${liqArticle.quantite} vs DUM ${dumArticle.quantite})`
      );
    }

    const taxes: TaxLine[] = liqArticle.taxes.map((tax) => ({ ...tax }));

    return {
      numero: liqArticle.numero,
      hsCode: liqArticle.hsCode,
      nomArticle: dumArticle.nomArticle,
      pays: dumArticle.paysNom,
      paysCode: dumArticle.paysCode,
      valeurDeclaree: dumArticle.valeurDeclaree,
      quantite: dumArticle.quantite,
      unite: dumArticle.unite,
      taxes,
      totalArticle: liqArticle.totalArticle,
    };
  });

  const dumOnlyOrdres = dum.articles
    .map((article) => article.ordre)
    .filter((ordre) => !liquidation.articles.some((a) => a.numero === ordre));

  if (dumOnlyOrdres.length > 0) {
    throw new MergeError(`Article(s) ${dumOnlyOrdres.join(', ')} present in DUM but not found in Liquidation`);
  }

  return {
    code: liquidation.header.code,
    redevable: liquidation.header.redevable,
    benNumero: liquidation.header.benNumero,
    articles,
  };
}

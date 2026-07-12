import type { Article } from './types.js';

export class ValidationError extends Error {
  constructor(message: string, public readonly field: string) {
    super(message);
    this.name = 'ValidationError';
  }
}

const ROUNDING_TOLERANCE = 0.02;

export function validateArticle(article: Article): void {
  if (article.quantite <= 0) {
    throw new ValidationError(
      `Article ${article.numero}: quantite must be > 0, got ${article.quantite}`,
      'quantite'
    );
  }

  const taxSum = article.taxes.reduce((sum, tax) => sum + tax.montant, 0);
  const diff = Math.abs(taxSum - article.totalArticle);
  if (diff > ROUNDING_TOLERANCE) {
    throw new ValidationError(
      `Article ${article.numero}: sum of tax montants (${taxSum.toFixed(2)}) ` +
        `does not match totalArticle (${article.totalArticle.toFixed(2)})`,
      'totalArticle'
    );
  }
}

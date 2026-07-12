import { describe, it, expect } from 'vitest';
import { validateArticle, ValidationError } from '../../src/domain/validators.js';
import type { Article } from '../../src/domain/types.js';

function makeArticle(overrides: Partial<Article> = {}): Article {
  return {
    numero: 1,
    hsCode: '6109100010',
    nomArticle: 'T-SHIRT',
    pays: 'ITALIE',
    paysCode: 'IT',
    valeurDeclaree: 27147,
    quantite: 354,
    unite: 'NOMBRE',
    taxes: [
      { code: '000110', assiette: 27147, taux: 0, montant: 0 },
      { code: '007217', assiette: 27147, taux: 0.25, montant: 68 },
      { code: '002109', assiette: 27215, taux: 20, montant: 5443 },
    ],
    totalArticle: 5511,
    ...overrides,
  };
}

describe('validateArticle', () => {
  it('passes for a well-formed article whose tax sum matches totalArticle', () => {
    expect(() => validateArticle(makeArticle())).not.toThrow();
  });

  it('throws when quantite is zero', () => {
    expect(() => validateArticle(makeArticle({ quantite: 0 }))).toThrow(ValidationError);
  });

  it('throws when tax montants do not sum to totalArticle', () => {
    expect(() => validateArticle(makeArticle({ totalArticle: 9999 }))).toThrow(ValidationError);
  });
});

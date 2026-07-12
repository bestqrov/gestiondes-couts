import { describe, it, expect } from 'vitest';
import { allocateTaxAcrossUnits, unionTaxCodes } from '../../src/excel/unitLevelTaxHelpers.js';
import type { Article } from '../../src/domain/types.js';

describe('allocateTaxAcrossUnits', () => {
  it('splits an evenly-divisible amount equally across all units', () => {
    const amounts = allocateTaxAcrossUnits(70, 350);
    expect(amounts).toHaveLength(350);
    for (const amount of amounts) {
      expect(amount).toBeCloseTo(0.2);
    }
    expect(amounts.reduce((sum, a) => sum + a, 0)).toBeCloseTo(70, 2);
  });

  it('spreads the rounding remainder cent-by-cent across the first rows so the column sum matches exactly', () => {
    const amounts = allocateTaxAcrossUnits(68, 354);
    expect(amounts).toHaveLength(354);
    // montantCents = 6800, baseCents = trunc(6800 / 354) = 19,
    // remainderCents = 6800 - 19*354 = 74 -> first 74 rows get an extra cent.
    for (let i = 0; i < 74; i++) {
      expect(amounts[i]).toBeCloseTo(0.2, 2);
    }
    for (let i = 74; i < 354; i++) {
      expect(amounts[i]).toBeCloseTo(0.19, 2);
    }
    const total = amounts.reduce((sum, a) => sum + a, 0);
    expect(total).toBeCloseTo(68, 2);
  });

  it('never produces a negative amount even when montant is small relative to quantite', () => {
    const amounts = allocateTaxAcrossUnits(0.37, 50);
    expect(amounts).toHaveLength(50);
    // montantCents = 37, baseCents = trunc(37 / 50) = 0, remainderCents = 37
    // -> first 37 rows get 0.01, remaining 13 rows get 0.00.
    for (let i = 0; i < 37; i++) {
      expect(amounts[i]).toBeCloseTo(0.01, 2);
    }
    for (let i = 37; i < 50; i++) {
      expect(amounts[i]).toBeCloseTo(0, 2);
    }
    expect(amounts.every((a) => a >= 0)).toBe(true);
    expect(amounts.reduce((sum, a) => sum + a, 0)).toBeCloseTo(0.37, 2);
  });

  it('handles a zero montant by allocating zero to every unit', () => {
    const amounts = allocateTaxAcrossUnits(0, 200);
    expect(amounts).toHaveLength(200);
    expect(amounts.every((a) => a === 0)).toBe(true);
  });

  it('throws when quantite is not a positive integer', () => {
    expect(() => allocateTaxAcrossUnits(100, 0)).toThrow();
    expect(() => allocateTaxAcrossUnits(100, -5)).toThrow();
    expect(() => allocateTaxAcrossUnits(100, 3.5)).toThrow();
  });
});

describe('unionTaxCodes', () => {
  function makeArticle(taxCodes: string[]): Article {
    return {
      numero: 1,
      hsCode: '6109100010',
      nomArticle: 'T-SHIRT',
      pays: 'ITALIE',
      paysCode: 'IT',
      valeurDeclaree: 100,
      quantite: 10,
      unite: 'U',
      taxes: taxCodes.map((code) => ({ code, assiette: 100, taux: 1, montant: 1 })),
      totalArticle: taxCodes.length,
    };
  }

  it('returns the sorted union of tax codes across all articles, deduplicated', () => {
    const articles = [makeArticle(['007217', '000110']), makeArticle(['002109', '000110'])];
    expect(unionTaxCodes(articles)).toEqual(['000110', '002109', '007217']);
  });

  it('returns an empty array for articles with no taxes', () => {
    expect(unionTaxCodes([makeArticle([])])).toEqual([]);
  });
});

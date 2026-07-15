import { describe, it, expect } from 'vitest';
import {
  aggregateCountryProductCounts,
  type TransactionDocument,
} from '../../src/db/transactionsRepository.js';

function makeTransaction(articles: TransactionDocument['articles']): TransactionDocument {
  return {
    ownerUserId: '1',
    code: '500001',
    redevable: 'Test SARL',
    valeurTotaleDeclaree: null,
    totalTaxes: 0,
    totalLandedCost: 0,
    costEstimatePartial: true,
    colisCount: null,
    referenceInterne: null,
    articles,
    createdAt: new Date().toISOString(),
  };
}

describe('aggregateCountryProductCounts', () => {
  it('returns an empty array when there are no transactions', () => {
    expect(aggregateCountryProductCounts([])).toEqual([]);
  });

  it('counts one article as one product for its country', () => {
    const transactions = [
      makeTransaction([
        { numero: 1, hsCode: '6109100010', nomArticle: 'T-SHIRT', pays: 'ITALIE', quantite: 354, costPerUnit: 42.5 },
      ]),
    ];
    expect(aggregateCountryProductCounts(transactions)).toEqual([
      { pays: 'ITALIE', productCount: 1, totalQuantite: 354 },
    ]);
  });

  it('sums productCount and totalQuantite across multiple articles/transactions for the same country', () => {
    const transactions = [
      makeTransaction([
        { numero: 1, hsCode: 'A', nomArticle: 'A', pays: 'ITALIE', quantite: 100, costPerUnit: 1 },
        { numero: 2, hsCode: 'B', nomArticle: 'B', pays: 'ITALIE', quantite: 50, costPerUnit: 1 },
      ]),
      makeTransaction([{ numero: 1, hsCode: 'C', nomArticle: 'C', pays: 'ITALIE', quantite: 25, costPerUnit: 1 }]),
    ];
    expect(aggregateCountryProductCounts(transactions)).toEqual([
      { pays: 'ITALIE', productCount: 3, totalQuantite: 175 },
    ]);
  });

  it('keeps separate countries separate, sorted by productCount descending', () => {
    const transactions = [
      makeTransaction([
        { numero: 1, hsCode: 'A', nomArticle: 'A', pays: 'BANGLADESH', quantite: 200, costPerUnit: 1 },
        { numero: 2, hsCode: 'B', nomArticle: 'B', pays: 'ITALIE', quantite: 354, costPerUnit: 1 },
        { numero: 3, hsCode: 'C', nomArticle: 'C', pays: 'ITALIE', quantite: 10, costPerUnit: 1 },
      ]),
    ];
    expect(aggregateCountryProductCounts(transactions)).toEqual([
      { pays: 'ITALIE', productCount: 2, totalQuantite: 364 },
      { pays: 'BANGLADESH', productCount: 1, totalQuantite: 200 },
    ]);
  });
});

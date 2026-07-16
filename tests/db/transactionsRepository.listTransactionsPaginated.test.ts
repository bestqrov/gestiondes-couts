import { describe, it, expect } from 'vitest';
import { createFakeCollection } from '../helpers/fakeMongoCollection.js';
import {
  listTransactionsPaginated,
  type TransactionDocument,
} from '../../src/db/transactionsRepository.js';

function makeCollection() {
  return createFakeCollection<TransactionDocument>();
}

describe('listTransactionsPaginated', () => {
  it('sorts newest first and paginates with the given page size', async () => {
    const collection = makeCollection();
    for (let i = 0; i < 5; i++) {
      await collection.insertOne({
        ownerUserId: 'u1',
        code: `50000${i}`,
        redevable: `SOCIETE ${i}`,
        valeurTotaleDeclaree: null,
        totalTaxes: 0,
        totalLandedCost: 0,
        costEstimatePartial: false,
        colisCount: null,
        referenceInterne: null,
        articles: [],
        createdAt: `2026-01-0${i + 1}T00:00:00.000Z`,
      });
    }

    const page1 = await listTransactionsPaginated(collection, { page: 1, pageSize: 2 });
    expect(page1.items.map((t) => t.code)).toEqual(['500004', '500003']);
    expect(page1.total).toBe(5);

    const page2 = await listTransactionsPaginated(collection, { page: 2, pageSize: 2 });
    expect(page2.items.map((t) => t.code)).toEqual(['500002', '500001']);

    const page3 = await listTransactionsPaginated(collection, { page: 3, pageSize: 2 });
    expect(page3.items.map((t) => t.code)).toEqual(['500000']);
  });

  it('filters by redevable (case-insensitive substring)', async () => {
    const collection = makeCollection();
    await collection.insertOne({
      ownerUserId: 'u1', code: 'A', redevable: 'Global Trade SARL', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-01-01T00:00:00.000Z',
    });
    await collection.insertOne({
      ownerUserId: 'u1', code: 'B', redevable: 'Other Company', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-01-02T00:00:00.000Z',
    });

    const result = await listTransactionsPaginated(collection, {
      page: 1,
      pageSize: 20,
      redevable: 'global',
    });
    expect(result.items.map((t) => t.code)).toEqual(['A']);
    expect(result.total).toBe(1);
  });

  it('filters by date range (dateFrom/dateTo, inclusive)', async () => {
    const collection = makeCollection();
    await collection.insertOne({
      ownerUserId: 'u1', code: 'A', redevable: 'X', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-01-01T00:00:00.000Z',
    });
    await collection.insertOne({
      ownerUserId: 'u1', code: 'B', redevable: 'X', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-01-15T00:00:00.000Z',
    });
    await collection.insertOne({
      ownerUserId: 'u1', code: 'C', redevable: 'X', valeurTotaleDeclaree: null,
      totalTaxes: 0, totalLandedCost: 0, costEstimatePartial: false, colisCount: null,
      referenceInterne: null, articles: [], createdAt: '2026-02-01T00:00:00.000Z',
    });

    const result = await listTransactionsPaginated(collection, {
      page: 1,
      pageSize: 20,
      dateFrom: '2026-01-10',
      dateTo: '2026-01-31',
    });
    expect(result.items.map((t) => t.code)).toEqual(['B']);
  });

  it('returns an empty list with total 0 when nothing matches', async () => {
    const collection = makeCollection();
    const result = await listTransactionsPaginated(collection, { page: 1, pageSize: 20 });
    expect(result.items).toEqual([]);
    expect(result.total).toBe(0);
  });
});

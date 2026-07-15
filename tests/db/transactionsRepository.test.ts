import { describe, it, expect } from 'vitest';
import {
  saveTransaction,
  listAllTransactions,
  getMostRecentTransaction,
  countTransactions,
  searchTransactionsByRedevable,
  type TransactionDocument,
} from '../../src/db/transactionsRepository.js';

// A minimal in-memory stand-in for the MongoDB driver's Collection<T>,
// implementing only the chained calls (find().sort().limit().toArray(),
// insertOne, countDocuments) this repository actually uses. mongodb-memory-
// server (a real embedded Mongo for tests) doesn't run reliably in this
// sandbox (crashes on startup), so this fake exercises the repository's
// query-building and mapping logic without a real Mongo connection.
function makeFakeCollection() {
  const docs: TransactionDocument[] = [];
  let nextId = 1;

  function matchesRegex(value: string, pattern: string): boolean {
    return new RegExp(pattern, 'i').test(value);
  }

  const collection = {
    docs,
    insertOne: async (doc: TransactionDocument) => {
      docs.push(doc);
      return { insertedId: `fake-id-${nextId++}` };
    },
    find: (filter?: { redevable?: { $regex: string; $options: string } }) => {
      let results = filter?.redevable
        ? docs.filter((d) => matchesRegex(d.redevable, filter.redevable!.$regex))
        : [...docs];
      let limitCount: number | undefined;
      const cursor = {
        sort: (spec: Record<string, 1 | -1>) => {
          const [key, dir] = Object.entries(spec)[0] as [keyof TransactionDocument, 1 | -1];
          results = [...results].sort((a, b) => {
            const av = String(a[key]);
            const bv = String(b[key]);
            return av < bv ? -dir : av > bv ? dir : 0;
          });
          return cursor;
        },
        limit: (n: number) => {
          limitCount = n;
          return cursor;
        },
        toArray: async () => (limitCount !== undefined ? results.slice(0, limitCount) : results),
      };
      return cursor;
    },
    countDocuments: async () => docs.length,
  };

  return collection as unknown as Parameters<typeof saveTransaction>[0];
}

function makeInput(overrides: Partial<Parameters<typeof saveTransaction>[1]> = {}) {
  return {
    ownerUserId: '1',
    code: '500001',
    redevable: 'Global Trade Logistics SARL',
    valeurTotaleDeclaree: 40039.992,
    totalTaxes: 5511,
    totalLandedCost: 15045.0,
    costEstimatePartial: false,
    articles: [
      {
        numero: 1,
        hsCode: '6109100010',
        nomArticle: 'T-SHIRT',
        pays: 'ITALIE',
        quantite: 354,
        costPerUnit: 42.5,
      },
    ],
    ...overrides,
  };
}

describe('transactionsRepository', () => {
  it('saves a transaction and reads it back via listAllTransactions', async () => {
    const collection = makeFakeCollection();

    const id = await saveTransaction(collection, makeInput());

    expect(id).toBeTruthy();
    const all = await listAllTransactions(collection);
    expect(all).toHaveLength(1);
    expect(all[0]).toMatchObject({
      code: '500001',
      redevable: 'Global Trade Logistics SARL',
      valeurTotaleDeclaree: 40039.992,
      totalTaxes: 5511,
      totalLandedCost: 15045.0,
      costEstimatePartial: false,
      colisCount: null,
      referenceInterne: null,
    });
    expect(all[0].articles).toEqual(makeInput().articles);
    expect(typeof all[0].createdAt).toBe('string');
  });

  it('defaults colisCount/referenceInterne to null when not provided, and stores them when given', async () => {
    const collection = makeFakeCollection();
    await saveTransaction(collection, makeInput({ colisCount: 6, referenceInterne: 'PO-1234' }));

    const [doc] = await listAllTransactions(collection);
    expect(doc.colisCount).toBe(6);
    expect(doc.referenceInterne).toBe('PO-1234');
  });

  it('getMostRecentTransaction returns the newest saved transaction', async () => {
    const collection = makeFakeCollection();
    await saveTransaction(collection, makeInput({ code: '111111' }));
    // Force a distinguishable, later createdAt for the second save.
    await new Promise((resolve) => setTimeout(resolve, 2));
    await saveTransaction(collection, makeInput({ code: '222222' }));

    const mostRecent = await getMostRecentTransaction(collection);
    expect(mostRecent?.code).toBe('222222');
  });

  it('getMostRecentTransaction returns null when nothing has been saved', async () => {
    const collection = makeFakeCollection();
    expect(await getMostRecentTransaction(collection)).toBeNull();
  });

  it('countTransactions reflects the number of saved transactions', async () => {
    const collection = makeFakeCollection();
    expect(await countTransactions(collection)).toBe(0);
    await saveTransaction(collection, makeInput());
    await saveTransaction(collection, makeInput({ code: '222222' }));
    expect(await countTransactions(collection)).toBe(2);
  });

  it('searchTransactionsByRedevable matches case-insensitively on a substring', async () => {
    const collection = makeFakeCollection();
    await saveTransaction(collection, makeInput({ redevable: 'Global Trade Logistics SARL' }));
    await saveTransaction(collection, makeInput({ code: '222222', redevable: 'Another Company SARL' }));

    const results = await searchTransactionsByRedevable(collection, 'global trade');
    expect(results).toHaveLength(1);
    expect(results[0].redevable).toBe('Global Trade Logistics SARL');
  });

  it('searchTransactionsByRedevable returns an empty array when nothing matches', async () => {
    const collection = makeFakeCollection();
    await saveTransaction(collection, makeInput());
    expect(await searchTransactionsByRedevable(collection, 'nonexistent')).toEqual([]);
  });

  it('escapes regex special characters in the search query instead of treating them as regex syntax', async () => {
    const collection = makeFakeCollection();
    await saveTransaction(collection, makeInput({ redevable: 'A+B Trading (Import)' }));

    // A literal search for "A+B" should match, not be interpreted as the regex quantifier "one or more A".
    const results = await searchTransactionsByRedevable(collection, 'A+B');
    expect(results).toHaveLength(1);
  });
});

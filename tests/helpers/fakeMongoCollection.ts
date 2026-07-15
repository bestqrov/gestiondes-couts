import { ObjectId, type Collection } from 'mongodb';

// A hand-built in-memory stand-in for a MongoDB Collection, used across
// this project's repository tests (mongodb-memory-server doesn't run in
// this sandbox — it crashes with SIGABRT). Implements only what the
// repositories under test actually call: insertOne, findOne,
// find().sort().toArray(), countDocuments, updateOne ($set + upsert), and
// createIndex (a no-op that also makes the given fields unique, mirroring
// a real unique index well enough to test duplicate-key handling).
export function createFakeCollection<T extends { _id?: ObjectId | string }>(): Collection<T> {
  const docs: T[] = [];
  const uniqueFields = new Set<string>();

  function idsEqual(a: unknown, b: unknown): boolean {
    if (a instanceof ObjectId && b instanceof ObjectId) return a.equals(b);
    return a === b;
  }

  function matches(doc: T, filter: Record<string, unknown>): boolean {
    return Object.entries(filter).every(([key, value]) => {
      if (key === '_id') return idsEqual(doc._id, value);
      return (doc as Record<string, unknown>)[key] === value;
    });
  }

  function assertUnique(candidate: Record<string, unknown>, excludeId?: ObjectId | string): void {
    for (const field of uniqueFields) {
      const value = candidate[field];
      if (value === undefined) continue;
      const clash = docs.find(
        (d) =>
          (d as Record<string, unknown>)[field] === value &&
          !(excludeId !== undefined && idsEqual(d._id, excludeId))
      );
      if (clash) {
        const error = new Error(`E11000 duplicate key error collection on field "${field}"`);
        (error as Error & { code: number }).code = 11000;
        throw error;
      }
    }
  }

  const fake = {
    createIndex: async (spec: Record<string, unknown>) => {
      for (const field of Object.keys(spec)) uniqueFields.add(field);
      return Object.keys(spec).join('_');
    },
    insertOne: async (doc: T) => {
      assertUnique(doc as Record<string, unknown>);
      const _id = new ObjectId();
      docs.push({ ...doc, _id });
      return { acknowledged: true, insertedId: _id };
    },
    findOne: async (filter: Record<string, unknown> = {}) => {
      const found = docs.find((d) => matches(d, filter));
      return found ? { ...found } : null;
    },
    find: (filter: Record<string, unknown> = {}) => {
      let results = docs.filter((d) => matches(d, filter));
      const cursor = {
        sort(sortSpec: Record<string, number>) {
          const [field, dir] = Object.entries(sortSpec)[0];
          results = [...results].sort((a, b) => {
            const av = (a as Record<string, unknown>)[field] as string | number;
            const bv = (b as Record<string, unknown>)[field] as string | number;
            if (av < bv) return -dir;
            if (av > bv) return dir;
            return 0;
          });
          return cursor;
        },
        toArray: async () => results.map((d) => ({ ...d })),
      };
      return cursor;
    },
    countDocuments: async (filter: Record<string, unknown> = {}) =>
      docs.filter((d) => matches(d, filter)).length,
    updateOne: async (
      filter: Record<string, unknown>,
      update: { $set?: Record<string, unknown> },
      options: { upsert?: boolean } = {}
    ) => {
      const idx = docs.findIndex((d) => matches(d, filter));
      const setFields = update.$set ?? {};
      if (idx === -1) {
        if (!options.upsert) {
          return { acknowledged: true, matchedCount: 0, modifiedCount: 0, upsertedId: null };
        }
        assertUnique(setFields);
        const _id = filter._id !== undefined ? filter._id : new ObjectId();
        docs.push({ _id, ...setFields } as T);
        return { acknowledged: true, matchedCount: 0, modifiedCount: 0, upsertedId: _id };
      }
      assertUnique(setFields, docs[idx]._id);
      docs[idx] = { ...docs[idx], ...setFields };
      return { acknowledged: true, matchedCount: 1, modifiedCount: 1, upsertedId: null };
    },
  };

  return fake as unknown as Collection<T>;
}

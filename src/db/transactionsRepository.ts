import type { Collection } from 'mongodb';

export interface TransactionArticle {
  numero: number;
  hsCode: string;
  nomArticle: string;
  pays: string;
  quantite: number;
  costPerUnit: number;
}

// Deliberately does NOT store the generated .xlsx file (or a path to it) —
// only the declaration's identifying/financial totals, per the decision to
// stop persisting Excel files long-term. The file itself stays a
// short-lived, per-process artifact (see server.ts's lastGeneratedFilePath)
// for the immediate re-download/PDF-export flow only.
export interface TransactionDocument {
  ownerUserId: number;
  code: string;
  redevable: string;
  valeurTotaleDeclaree: number | null;
  totalTaxes: number;
  totalLandedCost: number;
  costEstimatePartial: boolean;
  colisCount: number | null;
  referenceInterne: string | null;
  articles: TransactionArticle[];
  createdAt: string;
}

export interface SaveTransactionInput {
  ownerUserId: number;
  code: string;
  redevable: string;
  valeurTotaleDeclaree: number | null;
  totalTaxes: number;
  totalLandedCost: number;
  costEstimatePartial: boolean;
  colisCount?: number;
  referenceInterne?: string;
  articles: TransactionArticle[];
}

export const TRANSACTIONS_COLLECTION = 'transactions';

export async function saveTransaction(
  collection: Collection<TransactionDocument>,
  input: SaveTransactionInput
): Promise<string> {
  const doc: TransactionDocument = {
    ownerUserId: input.ownerUserId,
    code: input.code,
    redevable: input.redevable,
    valeurTotaleDeclaree: input.valeurTotaleDeclaree,
    totalTaxes: input.totalTaxes,
    totalLandedCost: input.totalLandedCost,
    costEstimatePartial: input.costEstimatePartial,
    colisCount: input.colisCount ?? null,
    referenceInterne: input.referenceInterne ?? null,
    articles: input.articles,
    createdAt: new Date().toISOString(),
  };
  const result = await collection.insertOne(doc);
  return String(result.insertedId);
}

export async function listAllTransactions(
  collection: Collection<TransactionDocument>
): Promise<TransactionDocument[]> {
  return collection.find().sort({ createdAt: -1 }).toArray();
}

export async function getMostRecentTransaction(
  collection: Collection<TransactionDocument>
): Promise<TransactionDocument | null> {
  const [mostRecent] = await collection.find().sort({ createdAt: -1 }).limit(1).toArray();
  return mostRecent ?? null;
}

export async function countTransactions(collection: Collection<TransactionDocument>): Promise<number> {
  return collection.countDocuments();
}

// Case-insensitive substring match on the company/redevable name, across
// every saved transaction (all admins) — matches the superadmin's "sees
// everything" role. $regex with a literal (non-anchored) pattern is fine
// at this data scale; query is not user-controlled beyond being embedded
// as a regex fragment, so a pathological input could only ever slow down
// this one query, not escape into other Mongo operations.
export async function searchTransactionsByRedevable(
  collection: Collection<TransactionDocument>,
  query: string
): Promise<TransactionDocument[]> {
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return collection
    .find({ redevable: { $regex: escaped, $options: 'i' } })
    .sort({ createdAt: -1 })
    .toArray();
}

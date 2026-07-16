import { ObjectId, type Collection } from 'mongodb';

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
  _id?: ObjectId;
  ownerUserId: string;
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
  ownerUserId: string;
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

export interface ListTransactionsOptions {
  page: number; // 1-indexed
  pageSize: number;
  redevable?: string;
  dateFrom?: string; // ISO date (yyyy-mm-dd), inclusive
  dateTo?: string; // ISO date (yyyy-mm-dd), inclusive
}

export interface ListTransactionsResult {
  items: TransactionDocument[];
  total: number;
  page: number;
  pageSize: number;
}

// Powers the superadmin "Historique" page — every saved transaction across
// all admins, newest first, with an optional redevable substring filter and
// an optional createdAt date range, paginated. dateTo is widened to the end
// of that calendar day (23:59:59.999) since createdAt is a precise
// timestamp but the filter inputs are plain dates.
export async function listTransactionsPaginated(
  collection: Collection<TransactionDocument>,
  options: ListTransactionsOptions
): Promise<ListTransactionsResult> {
  const filter: Record<string, unknown> = {};
  if (options.redevable) {
    const escaped = options.redevable.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    filter.redevable = { $regex: escaped, $options: 'i' };
  }
  if (options.dateFrom || options.dateTo) {
    const range: Record<string, string> = {};
    if (options.dateFrom) range.$gte = `${options.dateFrom}T00:00:00.000Z`;
    if (options.dateTo) range.$lte = `${options.dateTo}T23:59:59.999Z`;
    filter.createdAt = range;
  }

  const [items, total] = await Promise.all([
    collection
      .find(filter)
      .sort({ createdAt: -1 })
      .skip((options.page - 1) * options.pageSize)
      .limit(options.pageSize)
      .toArray(),
    collection.countDocuments(filter),
  ]);

  return { items, total, page: options.page, pageSize: options.pageSize };
}

export async function getTransactionById(
  collection: Collection<TransactionDocument>,
  id: string
): Promise<TransactionDocument | null> {
  return collection.findOne({ _id: new ObjectId(id) });
}

export interface CountryProductCount {
  pays: string;
  productCount: number;
  totalQuantite: number;
}

// Pure aggregation (no Mongo-specific query) so it's trivially unit-tested
// against plain objects — counts each article line-item once per country
// it belongs to (productCount), plus the summed physical unit quantity
// (totalQuantite), across every saved transaction (all admins). Sorted by
// totalQuantite descending — the dashboard legend displays actual unit
// quantity per country (not the article-line count, which under-counts:
// one declaration line for 354 T-shirts is 1 "product" but 354 units),
// so ranking matches what's actually shown.
export function aggregateCountryProductCounts(
  transactions: TransactionDocument[]
): CountryProductCount[] {
  const byCountry = new Map<string, { productCount: number; totalQuantite: number }>();
  for (const transaction of transactions) {
    for (const article of transaction.articles) {
      const existing = byCountry.get(article.pays) ?? { productCount: 0, totalQuantite: 0 };
      existing.productCount += 1;
      existing.totalQuantite += article.quantite;
      byCountry.set(article.pays, existing);
    }
  }
  return Array.from(byCountry.entries())
    .map(([pays, counts]) => ({ pays, ...counts }))
    .sort((a, b) => b.totalQuantite - a.totalQuantite);
}

export async function getCountryProductCounts(
  collection: Collection<TransactionDocument>
): Promise<CountryProductCount[]> {
  const transactions = await listAllTransactions(collection);
  return aggregateCountryProductCounts(transactions);
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

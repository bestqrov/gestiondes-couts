import { MongoClient, type Db } from 'mongodb';

// Lazy, memoized connection — the rest of the app (SQLite via better-sqlite3)
// boots synchronously, so Mongo is connected on first actual use rather
// than blocking startup. Memoizing the in-flight *promise* (not just the
// resolved value) avoids opening a second connection if multiple requests
// race to connect before the first one finishes.
let dbPromise: Promise<Db> | undefined;

export function getMongoDb(): Promise<Db> {
  if (!dbPromise) {
    const uri = process.env.MONGODB_URI;
    if (!uri) {
      return Promise.reject(new Error('MONGODB_URI is not set'));
    }
    const client = new MongoClient(uri);
    dbPromise = client.connect().then(
      (connectedClient) => connectedClient.db(),
      (error: unknown) => {
        // Don't memoize a failed connection attempt — a transient outage at
        // boot/first-request would otherwise permanently break Mongo access
        // until the process restarts. Clear so the next call retries.
        dbPromise = undefined;
        throw error;
      }
    );
  }
  return dbPromise;
}

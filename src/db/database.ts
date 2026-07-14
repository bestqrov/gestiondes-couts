import Database from 'better-sqlite3';
import { mkdirSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export function createDatabase(dbPath: string): Database.Database {
  if (dbPath !== ':memory:') {
    mkdirSync(path.dirname(dbPath), { recursive: true });
  }
  const db = new Database(dbPath);
  if (dbPath !== ':memory:') {
    db.pragma('journal_mode = WAL');
  }
  runMigrations(db);
  return db;
}

// SQLite now only holds users (auth) and app_settings (branding) — saved
// declarations ("transactions") moved to MongoDB; see src/db/mongoClient.ts
// and src/db/transactionsRepository.ts.
function runMigrations(db: Database.Database): void {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY,
      username TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL CHECK (role IN ('admin', 'superadmin')),
      created_at TEXT NOT NULL,
      disabled_at TEXT
    );

    -- Single-row table (id is always 1) holding app-wide branding: company
    -- name, logo (stored as a data: URI directly in the row — small enough
    -- in practice, and guarantees it persists in the same place/volume as
    -- everything else, no separate file storage/volume to configure), an
    -- accent color, and a font choice.
    CREATE TABLE IF NOT EXISTS app_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      company_name TEXT,
      logo_data_uri TEXT,
      brand_color TEXT,
      font_family TEXT,
      updated_at TEXT NOT NULL
    );
  `);
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_DB_PATH = path.resolve(__dirname, '../../data.db');

let dbInstance: Database.Database | undefined;

export function getDatabase(): Database.Database {
  if (!dbInstance) {
    dbInstance = createDatabase(process.env.DATABASE_PATH ?? DEFAULT_DB_PATH);
  }
  return dbInstance;
}

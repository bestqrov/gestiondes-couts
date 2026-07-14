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
    -- accent color, a font choice, and login-page contact details.
    CREATE TABLE IF NOT EXISTS app_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      company_name TEXT,
      logo_data_uri TEXT,
      brand_color TEXT,
      font_family TEXT,
      updated_at TEXT NOT NULL
    );
  `);

  // CREATE TABLE IF NOT EXISTS only affects a brand-new database — an
  // already-deployed one keeps its original column set. SQLite has no
  // "ADD COLUMN IF NOT EXISTS", so add any columns introduced after the
  // table's first release individually, guarded by a table_info check,
  // to keep this migration safe to run repeatedly against an existing file.
  const existingColumns = new Set(
    (db.prepare('PRAGMA table_info(app_settings)').all() as Array<{ name: string }>).map(
      (col) => col.name
    )
  );
  if (!existingColumns.has('contact_email')) {
    db.exec('ALTER TABLE app_settings ADD COLUMN contact_email TEXT');
  }
  if (!existingColumns.has('contact_whatsapp')) {
    db.exec('ALTER TABLE app_settings ADD COLUMN contact_whatsapp TEXT');
  }
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

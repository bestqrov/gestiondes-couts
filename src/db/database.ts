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

// Foreign keys (REFERENCES) are declared for documentation/future use but not enforced —
// PRAGMA foreign_keys is not set. Revisit if orphaned-row bugs surface.
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

    CREATE TABLE IF NOT EXISTS declarations (
      id INTEGER PRIMARY KEY,
      owner_user_id INTEGER NOT NULL REFERENCES users(id),
      code TEXT NOT NULL,
      redevable TEXT NOT NULL,
      ben_numero TEXT NOT NULL,
      devise TEXT,
      montant_facture REAL,
      taux_change REAL,
      fret REAL,
      assurance REAL,
      valeur_totale_declaree REAL,
      colis_count INTEGER,
      reference_interne TEXT,
      total_landed_cost REAL NOT NULL,
      cost_estimate_partial INTEGER NOT NULL DEFAULT 0,
      excel_file_path TEXT NOT NULL,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS declaration_articles (
      id INTEGER PRIMARY KEY,
      declaration_id INTEGER NOT NULL REFERENCES declarations(id),
      numero INTEGER NOT NULL,
      hs_code TEXT NOT NULL,
      nom_article TEXT NOT NULL,
      pays TEXT NOT NULL,
      valeur_declaree REAL NOT NULL,
      quantite REAL NOT NULL,
      total_article REAL NOT NULL,
      cost_per_unit REAL NOT NULL,
      taxes_json TEXT NOT NULL
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

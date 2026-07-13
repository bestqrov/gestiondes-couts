import Database from 'better-sqlite3';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export function createDatabase(dbPath: string): Database.Database {
  const db = new Database(dbPath);
  if (dbPath !== ':memory:') {
    db.pragma('journal_mode = WAL');
  }
  runMigrations(db);
  return db;
}

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

import { rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import Database from 'better-sqlite3';
import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';

describe('createDatabase', () => {
  it('creates the users and app_settings tables', () => {
    const db = createDatabase(':memory:');

    const tables = db
      .prepare("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
      .all()
      .map((row) => (row as { name: string }).name);

    expect(tables).toEqual(['app_settings', 'users']);
    db.close();
  });

  it('does not throw when migrations run twice against the same database file', () => {
    const dbPath = path.join(tmpdir(), `customs-app-test-db-${Date.now()}.sqlite`);
    const db1 = createDatabase(dbPath);
    db1.close();

    expect(() => {
      const db2 = createDatabase(dbPath);
      db2.close();
    }).not.toThrow();

    rmSync(dbPath, { force: true });
    rmSync(`${dbPath}-wal`, { force: true });
    rmSync(`${dbPath}-shm`, { force: true });
  });

  it('adds contact_email/contact_whatsapp via ALTER TABLE to an existing app_settings table that predates them', () => {
    const dbPath = path.join(tmpdir(), `customs-app-test-db-alter-${Date.now()}.sqlite`);

    // Simulate an already-deployed database created before contact_email/
    // contact_whatsapp existed, by building the table with the old column
    // set directly (bypassing createDatabase's current migration).
    const oldDb = new Database(dbPath);
    oldDb.exec(`
      CREATE TABLE app_settings (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        company_name TEXT,
        logo_data_uri TEXT,
        brand_color TEXT,
        font_family TEXT,
        updated_at TEXT NOT NULL
      );
    `);
    oldDb.close();

    expect(() => {
      const migratedDb = createDatabase(dbPath);
      const columns = (
        migratedDb.prepare('PRAGMA table_info(app_settings)').all() as Array<{ name: string }>
      ).map((col) => col.name);
      expect(columns).toContain('contact_email');
      expect(columns).toContain('contact_whatsapp');
      migratedDb.close();
    }).not.toThrow();

    rmSync(dbPath, { force: true });
    rmSync(`${dbPath}-wal`, { force: true });
    rmSync(`${dbPath}-shm`, { force: true });
  });
});

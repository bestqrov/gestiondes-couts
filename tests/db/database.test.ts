import { rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';

describe('createDatabase', () => {
  it('creates the users, declarations, declaration_articles, and app_settings tables', () => {
    const db = createDatabase(':memory:');

    const tables = db
      .prepare("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
      .all()
      .map((row) => (row as { name: string }).name);

    expect(tables).toEqual(['app_settings', 'declaration_articles', 'declarations', 'users']);
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
});

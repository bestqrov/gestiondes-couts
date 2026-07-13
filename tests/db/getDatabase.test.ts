import { rmSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { describe, it, expect, afterEach } from 'vitest';
import { getDatabase } from '../../src/db/database.js';

describe('getDatabase', () => {
  const dbPath = path.join(tmpdir(), `customs-app-getdb-test-${Date.now()}.sqlite`);

  afterEach(() => {
    rmSync(dbPath, { force: true });
    rmSync(`${dbPath}-wal`, { force: true });
    rmSync(`${dbPath}-shm`, { force: true });
  });

  it('creates the database file at the path from DATABASE_PATH and returns the same cached instance on repeated calls', () => {
    process.env.DATABASE_PATH = dbPath;

    const db1 = getDatabase();
    expect(existsSync(dbPath)).toBe(true);

    const db2 = getDatabase();
    expect(db2).toBe(db1); // same cached singleton instance

    delete process.env.DATABASE_PATH;
  });
});

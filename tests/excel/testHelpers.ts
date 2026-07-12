import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';

export function createTempXlsxPath(basename: string): { filePath: string; dir: string } {
  const dir = mkdtempSync(path.join(tmpdir(), 'customs-excel-test-'));
  return { filePath: path.join(dir, `${basename}.xlsx`), dir };
}

export function cleanupTempDir(dir: string): void {
  rmSync(dir, { recursive: true, force: true });
}

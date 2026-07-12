import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { detectAndParsePair } from '../../src/parser/detectAndParsePair.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures');

function loadFixtures() {
  const liquidationText = readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8');
  const dumText = readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8');
  return { liquidationText, dumText };
}

describe('detectAndParsePair', () => {
  it('parses correctly when files are in the expected order', () => {
    const { liquidationText, dumText } = loadFixtures();
    const result = detectAndParsePair(liquidationText, dumText);

    expect(result.swapped).toBe(false);
    expect(result.liquidation.header.code).toBe(result.dum.creditEnlevementCode.slice(0, 6));
    expect(result.dum.articles).toHaveLength(2);
  });

  it('parses correctly and reports swapped=true when files are in the reversed order', () => {
    const { liquidationText, dumText } = loadFixtures();
    const result = detectAndParsePair(dumText, liquidationText);

    expect(result.swapped).toBe(true);
    expect(result.dum.articles).toHaveLength(2);
    expect(result.liquidation.articles).toHaveLength(2);
  });

  it('throws a combined error when neither assignment parses successfully', () => {
    expect(() => detectAndParsePair('garbage text one', 'garbage text two')).toThrow(
      /Could not identify which uploaded file/
    );
  });
});

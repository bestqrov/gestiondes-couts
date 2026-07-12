import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { validateArticle } from '../../src/domain/validators.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

describe('parse + merge + validate pipeline', () => {
  it('produces a fully valid Declaration from the two real sample documents', () => {
    const liquidationText = readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8');
    const dumText = readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8');

    const liquidation = parseLiquidation(liquidationText);
    const dum = parseDum(dumText);
    const declaration = mergeDeclaration(liquidation, dum);

    expect(declaration.articles).toHaveLength(2);

    for (const article of declaration.articles) {
      expect(() => validateArticle(article)).not.toThrow();
    }

    const totalDeclaredValue = declaration.articles.reduce((sum, a) => sum + a.valeurDeclaree, 0);
    expect(totalDeclaredValue).toBeCloseTo(27147.0 + 12892.992, 1);
  });
});

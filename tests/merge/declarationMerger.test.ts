import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration, MergeError } from '../../src/merge/declarationMerger.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dum = parseDum(readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8'));
  return { liquidation, dum };
}

describe('mergeDeclaration', () => {
  it('merges the real Liquidation and DUM samples into a single Declaration', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const declaration = mergeDeclaration(liquidation, dum);

    expect(declaration.code).toBe('500001');
    expect(declaration.redevable).toBe('GLOBAL TRADE LOGISTICS SARL');
    expect(declaration.articles).toHaveLength(2);

    const [article1, article2] = declaration.articles;

    expect(article1.nomArticle).toBe('T-SHIRT');
    expect(article1.pays).toBe('ITALIE');
    expect(article1.paysCode).toBe('IT');
    expect(article1.hsCode).toBe('6109100010');
    expect(article1.quantite).toBeCloseTo(354.0);
    expect(article1.taxes.map((t) => t.code)).toEqual(['000110', '007217', '002109']);
    expect(article1.totalArticle).toBeCloseTo(5511.0);

    expect(article2.pays).toBe('BANGLADESH');
    expect(article2.paysCode).toBe('BD');
    expect(article2.quantite).toBeCloseTo(200.0);
  });

  it('throws when the DUM crédit d\'enlèvement code does not match the Liquidation code', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const mismatchedDum = { ...dum, creditEnlevementCode: '999999999' };

    expect(() => mergeDeclaration(liquidation, mismatchedDum)).toThrow(MergeError);
  });

  it('throws when an article present in Liquidation is missing from DUM', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const dumMissingArticle = { ...dum, articles: dum.articles.slice(0, 1) };

    expect(() => mergeDeclaration(liquidation, dumMissingArticle)).toThrow(MergeError);
  });

  it('throws when declared value differs beyond tolerance between the two documents', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const skewedDum = {
      ...dum,
      articles: dum.articles.map((a) => (a.ordre === 1 ? { ...a, valeurDeclaree: 999999 } : a)),
    };

    expect(() => mergeDeclaration(liquidation, skewedDum)).toThrow(MergeError);
  });

  it('throws when the HS code differs between Liquidation and DUM for the same article', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const mismatchedHsCodeDum = {
      ...dum,
      articles: dum.articles.map((a) => (a.ordre === 1 ? { ...a, hsCode: '9999999999' } : a)),
    };

    expect(() => mergeDeclaration(liquidation, mismatchedHsCodeDum)).toThrow(MergeError);
  });

  it('throws when quantité differs beyond tolerance between the two documents', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const skewedQuantiteDum = {
      ...dum,
      articles: dum.articles.map((a) => (a.ordre === 1 ? { ...a, quantite: a.quantite + 999 } : a)),
    };

    expect(() => mergeDeclaration(liquidation, skewedQuantiteDum)).toThrow(MergeError);
  });

  it('throws when an article present in DUM is missing from Liquidation', () => {
    const { liquidation, dum } = loadRealDeclaration();
    const liquidationMissingArticle = { ...liquidation, articles: liquidation.articles.slice(0, 1) };

    expect(() => mergeDeclaration(liquidationMissingArticle, dum)).toThrow(MergeError);
  });

  it('still matches the real fixture pair (Liquidation "500001" vs DUM "500001099")', () => {
    const { liquidation, dum } = loadRealDeclaration();

    expect(() => mergeDeclaration(liquidation, dum)).not.toThrow();
  });

  it('throws when the DUM code has the right prefix but the wrong suffix length', () => {
    const { liquidation, dum } = loadRealDeclaration();

    const tooFewDigits = { ...dum, creditEnlevementCode: '50000113' }; // only 2 extra digits
    const tooManyDigits = { ...dum, creditEnlevementCode: '5000010990' }; // 4 extra digits

    expect(() => mergeDeclaration(liquidation, tooFewDigits)).toThrow(MergeError);
    expect(() => mergeDeclaration(liquidation, tooManyDigits)).toThrow(MergeError);
  });
});

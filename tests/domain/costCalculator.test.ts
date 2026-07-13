import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';
import { parseDum } from '../../src/parser/dum/dumParser.js';
import { mergeDeclaration } from '../../src/merge/declarationMerger.js';
import { calculateLandedCost } from '../../src/domain/costCalculator.js';

const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '../parser/fixtures');

function loadRealDeclaration() {
  const liquidation = parseLiquidation(
    readFileSync(path.join(fixturesDir, 'liquidation-sample-1.txt'), 'utf-8')
  );
  const dumText = readFileSync(path.join(fixturesDir, 'dum-sample-1.txt'), 'utf-8');
  const dum = parseDum(dumText);
  return { declaration: mergeDeclaration(liquidation, dum), shipmentCost: dum.shipmentCost! };
}

describe('calculateLandedCost', () => {
  it('computes a full (non-partial) landed cost whose total equals shipment-cost + total taxes, regardless of per-article split', () => {
    const { declaration, shipmentCost } = loadRealDeclaration();

    const result = calculateLandedCost(declaration, shipmentCost);

    expect(result.partial).toBe(false);
    expect(result.articleCosts).toHaveLength(2);

    const totalTaxes = declaration.articles.reduce(
      (sum, a) => sum + a.taxes.reduce((s, t) => s + t.montant, 0),
      0
    );
    const expectedTotal =
      (shipmentCost.montantFacture + shipmentCost.fret + shipmentCost.assurance) *
        shipmentCost.tauxChange +
      totalTaxes;

    expect(result.totalLandedCost).toBeCloseTo(expectedTotal, 2);

    // Reconciliation: summing (costPerUnit * quantite) per article must equal the total.
    const reconciledTotal = result.articleCosts.reduce((sum, c) => {
      const article = declaration.articles.find((a) => a.numero === c.numero)!;
      return sum + c.costPerUnit * article.quantite;
    }, 0);
    expect(reconciledTotal).toBeCloseTo(expectedTotal, 2);

    // Every article's cost per unit must be positive and finite.
    for (const cost of result.articleCosts) {
      expect(cost.costPerUnit).toBeGreaterThan(0);
      expect(Number.isFinite(cost.costPerUnit)).toBe(true);
    }
  });

  it('falls back to duty-only-per-unit and marks the result partial when shipment cost fields are unavailable', () => {
    const { declaration } = loadRealDeclaration();

    const result = calculateLandedCost(declaration, {});

    expect(result.partial).toBe(true);

    const article1 = declaration.articles.find((a) => a.numero === 1)!;
    const article1TaxTotal = article1.taxes.reduce((sum, t) => sum + t.montant, 0);
    const cost1 = result.articleCosts.find((c) => c.numero === 1)!;
    expect(cost1.costPerUnit).toBeCloseTo(article1TaxTotal / article1.quantite, 4);
  });

  it('treats a partially-populated shipment cost (e.g. missing tauxChange) as partial, not a crash', () => {
    const { declaration, shipmentCost } = loadRealDeclaration();
    const { tauxChange, ...incomplete } = shipmentCost;

    const result = calculateLandedCost(declaration, incomplete);

    expect(result.partial).toBe(true);
    for (const cost of result.articleCosts) {
      expect(Number.isFinite(cost.costPerUnit)).toBe(true);
    }
  });

  it('marks the result partial (and omits shipment cost from the total) when every article has a zero declared value, even with full shipment data', () => {
    const { declaration, shipmentCost } = loadRealDeclaration();
    for (const article of declaration.articles) {
      article.valeurDeclaree = 0;
    }

    const result = calculateLandedCost(declaration, shipmentCost);

    expect(result.partial).toBe(true);

    const totalTaxes = declaration.articles.reduce(
      (sum, a) => sum + a.taxes.reduce((s, t) => s + t.montant, 0),
      0
    );
    expect(result.totalLandedCost).toBeCloseTo(totalTaxes, 2);
  });

  it('throws when an article has a non-positive quantite instead of silently returning a 0 cost per unit', () => {
    const { declaration, shipmentCost } = loadRealDeclaration();
    declaration.articles[0].quantite = 0;

    expect(() => calculateLandedCost(declaration, shipmentCost)).toThrow('quantite must be > 0');
  });
});

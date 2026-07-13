import type { Declaration } from './types.js';

export interface PartialShipmentCost {
  devise?: string;
  montantFacture?: number;
  tauxChange?: number;
  fret?: number;
  assurance?: number;
  valeurTotaleDeclaree?: number;
}

export interface ArticleCostResult {
  numero: number;
  costPerUnit: number;
}

export interface CostCalculationResult {
  articleCosts: ArticleCostResult[];
  totalLandedCost: number;
  partial: boolean;
}

// Per design spec §5: each article's share of shipment-level costs
// (invoiced amount + freight + insurance, converted via the DUM's own
// exchange rate) is allocated proportionally to its share of the
// shipment's total declared value — the same principle already used for
// tax-per-unit splitting, applied one level up. `partial` is true whenever
// the result omits a real cost component the caller should not present as
// complete: either a shipment-level field is missing, or every article's
// valeurDeclaree is zero/absent (so a real montantFacture/fret/assurance
// exists but has nothing to allocate against, and would otherwise be
// silently dropped from totalLandedCost).
//
// quantite must be > 0 — callers are expected to run validateArticle()
// (src/domain/validators.ts) first, which already enforces this; this
// function throws rather than silently returning a 0 cost per unit, to
// avoid presenting bad data as a valid answer.
export function calculateLandedCost(
  declaration: Declaration,
  shipment: PartialShipmentCost
): CostCalculationResult {
  const { montantFacture, fret, assurance, tauxChange } = shipment;
  const hasFullShipmentData =
    montantFacture !== undefined &&
    fret !== undefined &&
    assurance !== undefined &&
    tauxChange !== undefined;

  const totalDeclaredValue = declaration.articles.reduce((sum, a) => sum + a.valeurDeclaree, 0);
  const canAllocateShipmentCost = hasFullShipmentData && totalDeclaredValue > 0;

  let totalLandedCost = 0;
  const articleCosts: ArticleCostResult[] = declaration.articles.map((article) => {
    if (article.quantite <= 0) {
      throw new Error(
        `Article ${article.numero}: quantite must be > 0, got ${article.quantite}`
      );
    }

    const taxTotal = article.taxes.reduce((sum, tax) => sum + tax.montant, 0);

    let shipmentCostForArticle = 0;
    if (canAllocateShipmentCost) {
      const share = article.valeurDeclaree / totalDeclaredValue;
      shipmentCostForArticle = (montantFacture + fret + assurance) * tauxChange * share;
    }

    const totalArticleCost = shipmentCostForArticle + taxTotal;
    const costPerUnit = totalArticleCost / article.quantite;

    totalLandedCost += totalArticleCost;
    return { numero: article.numero, costPerUnit };
  });

  const partial = !hasFullShipmentData || totalDeclaredValue <= 0;

  return { articleCosts, totalLandedCost, partial };
}

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
// tax-per-unit splitting, applied one level up. If any of the four
// shipment-level fields is missing, shipment cost is treated as zero and
// the result is flagged `partial` so callers don't present an incomplete
// number as a complete one.
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

  const articleCosts: ArticleCostResult[] = declaration.articles.map((article) => {
    const taxTotal = article.taxes.reduce((sum, tax) => sum + tax.montant, 0);

    let shipmentCostForArticle = 0;
    if (hasFullShipmentData && totalDeclaredValue > 0) {
      const share = article.valeurDeclaree / totalDeclaredValue;
      shipmentCostForArticle = (montantFacture + fret + assurance) * tauxChange * share;
    }

    const totalArticleCost = shipmentCostForArticle + taxTotal;
    const costPerUnit = article.quantite > 0 ? totalArticleCost / article.quantite : 0;

    return { numero: article.numero, costPerUnit };
  });

  const totalLandedCost = articleCosts.reduce((sum, cost) => {
    const article = declaration.articles.find((a) => a.numero === cost.numero)!;
    return sum + cost.costPerUnit * article.quantite;
  }, 0);

  return { articleCosts, totalLandedCost, partial: !hasFullShipmentData };
}

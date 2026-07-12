export interface TaxLine {
  code: string;
  assiette: number;
  taux: number;
  montant: number;
}

export interface Article {
  numero: number;
  hsCode: string;
  nomArticle: string;
  pays: string;
  paysCode: string;
  valeurDeclaree: number;
  quantite: number;
  unite: string;
  taxes: TaxLine[];
  totalArticle: number;
}

// Declaration dates (échéance paiement, date déclaration, etc., visible in the
// Liquidation fixture) are deliberately deferred: no confirmed Excel column or
// UI requirement consumes them yet. Add here + to LiquidationParser when one does.
export interface Declaration {
  code: string;
  redevable: string;
  benNumero: string;
  articles: Article[];
}

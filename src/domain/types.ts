export interface TaxLine {
  code: string;
  assiette: number;
  taux: number;
  montant: number;
}

// A declaration-wide rubrique from the Liquidation's RECAPITULATION table
// (e.g. REDV.INF., RI SEGMA, REMISES CREDIT) — never allocated per article
// the way TaxLine's are, only known as one montant for the whole declaration.
export interface OrdonnancementTaxLine {
  code: string;
  designation: string;
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
  ordonnancementTaxes: OrdonnancementTaxLine[];
}

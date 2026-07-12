import type { Article } from '../domain/types.js';

export function allocateTaxAcrossUnits(montant: number, quantite: number): number[] {
  if (!Number.isInteger(quantite) || quantite <= 0) {
    throw new Error(
      `quantite must be a positive integer to allocate tax across units, got ${quantite}`
    );
  }

  // Work in integer cents to avoid floating-point drift, and spread the
  // remainder one cent at a time across the first `remainderCents` rows
  // (rather than dumping it all on the last row) so no row's amount can
  // go negative when montant >= 0 — a negative per-unit tax would look
  // like a data error to the accounting staff reading this sheet.
  const montantCents = Math.round(montant * 100);
  const baseCents = Math.trunc(montantCents / quantite);
  const remainderCents = montantCents - baseCents * quantite;

  const amounts = new Array<number>(quantite);
  for (let i = 0; i < quantite; i++) {
    const cents = baseCents + (i < remainderCents ? 1 : 0);
    amounts[i] = cents / 100;
  }
  return amounts;
}

export function unionTaxCodes(articles: Article[]): string[] {
  const codes = new Set<string>();
  for (const article of articles) {
    for (const tax of article.taxes) {
      codes.add(tax.code);
    }
  }
  return Array.from(codes).sort();
}

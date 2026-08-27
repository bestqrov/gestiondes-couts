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

// Rubrique code -> designation, from the customs office's own reference
// table. Codes not in this table are left with a blank header (rather than
// falling back to the raw code) until a designation is provided for them.
const TAX_CODE_DESIGNATIONS: Record<string, string> = {
  '002701': 'REDV.INF.(AVEC D et T)',
  '006901': 'RI SEGMA',
  '000110': 'DTS IMPORT NORMAL',
  '001102': 'TIC SUR LES ALCOOLS',
  '007217': 'TAXE F.P.E.I. EXP.',
  '002109': 'TVA IMPORT AUTRE PDS',
  '004801': 'FDS PROT.ENVI.DEV.DUR',
  '006000': 'REMISES CREDIT',
};

export function taxCodeDesignation(code: string): string {
  return TAX_CODE_DESIGNATIONS[code] ?? '';
}

// First letter of each word in a tax designation, e.g. "TVA IMPORT AUTRE
// PDS" -> "TIAP" — used to build a short header for that tax's "Total"
// column (montant x pieces) without hardcoding a separate abbreviation table.
export function abbreviateDesignation(designation: string): string {
  const words = designation.match(/[A-Za-z0-9]+/g) ?? [];
  return words.map((word) => word[0].toUpperCase()).join('');
}

// RI SEGMA and REMISES CREDIT both abbreviate down to "RC" via pure
// initials, so their "Total" columns were indistinguishable — spelled-out
// short names instead of computed initials for these two specifically.
const TAX_ABBREVIATION_OVERRIDES: Record<string, string> = {
  '006901': 'Risg', // RI SEGMA
  '006000': 'Rems', // REMISES CREDIT
};

export function taxColumnAbbreviation(code: string, designation: string): string {
  return TAX_ABBREVIATION_OVERRIDES[code] ?? abbreviateDesignation(designation);
}

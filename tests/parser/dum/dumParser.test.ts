import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseDum } from '../../../src/parser/dum/dumParser.js';

const fixturePath = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  '../fixtures/dum-sample-1.txt'
);

describe('parseDum', () => {
  it('parses the crédit d\'enlèvement code and both articles from the real sample document', () => {
    const text = readFileSync(fixturePath, 'utf-8');
    const result = parseDum(text);

    expect(result.creditEnlevementCode).toBe('500001099');
    expect(result.articles).toHaveLength(2);

    const [article1, article2] = result.articles;

    expect(article1).toEqual({
      ordre: 1,
      hsCode: '6109100010',
      nomArticle: 'T-SHIRT',
      paysNom: 'ITALIE',
      paysCode: 'IT',
      valeurDeclaree: 27147.0,
      quantite: 354.0,
      unite: 'U',
    });

    expect(article2).toEqual({
      ordre: 2,
      hsCode: '6109100010',
      nomArticle: 'T-SHIRT',
      paysNom: 'BANGLADESH',
      paysCode: 'BD',
      valeurDeclaree: 12892.992,
      quantite: 200.0,
      unite: 'U',
    });
  });

  it('throws when no articles are found', () => {
    expect(() => parseDum("Crédit d'enlèvement 123")).toThrow('No articles found');
  });

  it('throws when the Crédit d\'enlèvement code is missing from the document', () => {
    const text = `N° d'ordre de l'art. : 1
Colis et désignation des marchandises : T-SHIRT 354.00 NB
Code marchandises : 6109100010
Valeur déclarée : 27 147.000
Unités complémentaires : 354.0 U
Pays d'origine (Nom et code) : ITALIE IT
`;
    expect(() => parseDum(text)).toThrow("Crédit d'enlèvement code not found");
  });

  it('throws mentioning the article number when a required article field is missing', () => {
    // No "Code marchandises :" line for article 1.
    const text = `N° d'ordre de l'art. : 1
Colis et désignation des marchandises : T-SHIRT 354.00 NB
Valeur déclarée : 27 147.000
Unités complémentaires : 354.0 U
Pays d'origine (Nom et code) : ITALIE IT

Renseignements financiers et bancaires : 02 Crédit d'enlèvement 500001099
`;
    expect(() => parseDum(text)).toThrow(/DUM article 1/);
  });

  it('throws when the designation line does not match the expected "NAME QTY UNIT" format', () => {
    const text = `N° d'ordre de l'art. : 1
Colis et désignation des marchandises : SOMETHING WEIRD
Code marchandises : 6109100010
Valeur déclarée : 27 147.000
Unités complémentaires : 354.0 U
Pays d'origine (Nom et code) : ITALIE IT

Renseignements financiers et bancaires : 02 Crédit d'enlèvement 500001099
`;
    expect(() => parseDum(text)).toThrow(/designation line does not match/);
  });

  it('parses a document with a single article, correctly terminating the block at end-of-string', () => {
    // The credit code appears before the article block (not after) so that the article block's
    // lookahead has no "Renseignements financiers" or next-article marker to stop at — it must
    // rely on the "$" end-of-string alternative to terminate correctly.
    const text = `Renseignements financiers et bancaires : 02 Crédit d'enlèvement 500001099

N° d'ordre de l'art. : 1
Colis et désignation des marchandises : T-SHIRT 354.00 NB
Code marchandises : 6109100010
Valeur déclarée : 27 147.000
Unités complémentaires : 354.0 U
Pays d'origine (Nom et code) : ITALIE IT
`;
    const result = parseDum(text);

    expect(result.creditEnlevementCode).toBe('500001099');
    expect(result.articles).toHaveLength(1);
    expect(result.articles[0]).toEqual({
      ordre: 1,
      hsCode: '6109100010',
      nomArticle: 'T-SHIRT',
      paysNom: 'ITALIE',
      paysCode: 'IT',
      valeurDeclaree: 27147.0,
      quantite: 354.0,
      unite: 'U',
    });
  });
});

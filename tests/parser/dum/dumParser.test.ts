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
  it('parses the crédit d\'enlèvement code and both articles from the real, jumbled pdfjs-extracted text', () => {
    // This fixture is the actual raw text pdfjs-dist extracts from a real DUM
    // PDF — labels and values are scattered non-adjacently (extraction
    // follows the PDF's content-stream order, not visual reading order), so
    // this is what the parser must handle in production, not a hand-cleaned
    // approximation.
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
    const text =
      '6109100010   27 147.000 43.69   AP 354.0 U ITALIE   IT  COLIS  T-SHIRT 354.00 NB 1';
    expect(() => parseDum(text)).toThrow("Crédit d'enlèvement code not found");
  });

  it('parses a single-article document with a different HS code, country, and product name (not overfit to the one real sample)', () => {
    const text = `Crédit d'enlèvement 700002123
8471300000   9 500.000 12.00   AP 40.0 U ESPAGNE   ES  MARCHANDISES EMBALLEES  ORDINATEUR PORTABLE 40.00 NB 1`;
    const result = parseDum(text);

    expect(result.creditEnlevementCode).toBe('700002123');
    expect(result.articles).toEqual([
      {
        ordre: 1,
        hsCode: '8471300000',
        nomArticle: 'ORDINATEUR PORTABLE',
        paysNom: 'ESPAGNE',
        paysCode: 'ES',
        valeurDeclaree: 9500.0,
        quantite: 40.0,
        unite: 'U',
      },
    ]);
  });

  it('does not sweep preceding boilerplate words (e.g. "COLIS") into the product name', () => {
    // The désignation capture requires a 2+ space gap immediately before it —
    // this test pins that a single-word boilerplate token right before the
    // real name is correctly excluded.
    const text = `Crédit d'enlèvement 700002123
6109100010   1 000.000 5.00   AP 10.0 U MAROC   MA  COLIS  CHEMISE 10.00 NB 1`;
    const result = parseDum(text);

    expect(result.articles[0].nomArticle).toBe('CHEMISE');
  });

  it('extracts shipment-level cost fields (devise, montant facturé, taux de change, fret, assurance, valeur totale déclarée) from the real sample', () => {
    const text = readFileSync(fixturePath, 'utf-8');
    const result = parseDum(text);

    expect(result.shipmentCost).toEqual({
      devise: 'EUR',
      montantFacture: 2981.34,
      tauxChange: 10.6675,
      fret: 7467.0,
      assurance: 118.0,
      valeurTotaleDeclaree: 40039.992,
    });
  });

  it('leaves shipmentCost undefined (not a hard failure) when the cluster is not found', () => {
    const text = `Crédit d'enlèvement 700002123
6109100010   1 000.000 5.00   AP 10.0 U MAROC   MA  COLIS  CHEMISE 10.00 NB 1`;
    const result = parseDum(text);

    expect(result.shipmentCost).toBeUndefined();
  });
});
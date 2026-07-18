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
      unite: 'NB',
    });

    expect(article2).toEqual({
      ordre: 2,
      hsCode: '6109100010',
      nomArticle: 'T-SHIRT',
      paysNom: 'BANGLADESH',
      paysCode: 'BD',
      valeurDeclaree: 12892.992,
      quantite: 200.0,
      unite: 'NB',
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
        unite: 'NB',
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

  it('handles a "(PTI:OUI)" suffix directly after the HS code, with no whitespace before it', () => {
    // Real-world regression: a code like "3303000020(PTI:OUI)" used to break
    // the match entirely (the pattern required whitespace right after the
    // 10-digit code), which was reported as "Article 1 present in
    // Liquidation but not found in DUM" — the DUM parser had silently
    // dropped the article rather than mismatching a field.
    const text = `Crédit d'enlèvement 700002123
3303000020(PTI:OUI)   39 518.611 195.633   AP 195.633 KG ITALIE   IT  COLIS  EAU DE TOILETTE 480.00 NB 1`;
    const result = parseDum(text);

    expect(result.articles).toEqual([
      {
        ordre: 1,
        hsCode: '3303000020',
        nomArticle: 'EAU DE TOILETTE',
        paysNom: 'ITALIE',
        paysCode: 'IT',
        valeurDeclaree: 39518.611,
        quantite: 480.0,
        unite: 'NB',
      },
    ]);
  });

  it('captures the true piece quantity from the désignation line, not the "unités complémentaires" field, when they differ (e.g. a KG-based complementary unit for a good classified by weight)', () => {
    // Same real-world case as above, phrased as its own assertion: 480 (the
    // Liquidation's actual quantity) must win over 195.633 (the unrelated
    // complementary-unit weight in KG) — using the latter would silently
    // corrupt the merged declaration instead of failing loudly, since it's
    // a plausible-looking number, not an obviously wrong one.
    const text = `Crédit d'enlèvement 700002123
3303000020(PTI:OUI)   39 518.611 195.633   AP 195.633 KG ITALIE   IT  COLIS  EAU DE TOILETTE 480.00 NB 1`;
    const result = parseDum(text);

    expect(result.articles[0].quantite).toBe(480.0);
  });

  it('recognizes "PAIRE" as a valid désignation-line unit (e.g. footwear), not just "NB"/"U"', () => {
    const text = `Crédit d'enlèvement 700002123
6402999093(PTI:OUI)   346.186 1.10   AP 4.0 U CHINE   CN  MARCHANDISES NON EMBALLEE  AUTRE CHAUSSURES 4.00 PAIRE 31`;
    const result = parseDum(text);

    expect(result.articles).toEqual([
      {
        ordre: 31,
        hsCode: '6402999093',
        nomArticle: 'AUTRE CHAUSSURES',
        paysNom: 'CHINE',
        paysCode: 'CN',
        valeurDeclaree: 346.186,
        quantite: 4.0,
        unite: 'PAIRE',
      },
    ]);
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

  it('maps each capture group to the correct field, independent of the real fixture\'s incidental value ordering', () => {
    // Synthetic values are chosen to be pairwise distinct and not sorted by
    // magnitude, so a transposed capture group (e.g. fret/assurance swapped)
    // would fail this test even though it might not fail against the real
    // fixture, where the values happen to differ enough in magnitude to mask
    // such a bug.
    const text = `Crédit d'enlèvement 700002123
6109100010   1 000.000 5.00   AP 10.0 U MAROC   MA  COLIS  CHEMISE 10.00 NB 1
USD 1111.11 22.2 3333.33 99 44.4 5555.55 01 02 2026`;
    const result = parseDum(text);

    expect(result.shipmentCost).toEqual({
      devise: 'USD',
      montantFacture: 1111.11,
      tauxChange: 22.2,
      fret: 3333.33,
      assurance: 44.4,
      valeurTotaleDeclaree: 5555.55,
    });
  });

  it('leaves shipmentCost undefined (not a hard failure) when the cluster is not found', () => {
    const text = `Crédit d'enlèvement 700002123
6109100010   1 000.000 5.00   AP 10.0 U MAROC   MA  COLIS  CHEMISE 10.00 NB 1`;
    const result = parseDum(text);

    expect(result.shipmentCost).toBeUndefined();
  });
});
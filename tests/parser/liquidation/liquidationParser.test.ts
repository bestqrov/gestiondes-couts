import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { parseLiquidation } from '../../../src/parser/liquidation/liquidationParser.js';

const fixturePath = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  '../fixtures/liquidation-sample-1.txt'
);

describe('parseLiquidation', () => {
  it('parses header and both articles from the real sample document', () => {
    const text = readFileSync(fixturePath, 'utf-8');
    const result = parseLiquidation(text);

    expect(result.header).toEqual({
      code: '500001',
      redevable: 'GLOBAL TRADE LOGISTICS SARL',
      benNumero: '501',
    });

    expect(result.articles).toHaveLength(2);

    const [article1, article2] = result.articles;

    expect(article1.numero).toBe(1);
    expect(article1.hsCode).toBe('6109100010');
    expect(article1.valeurDeclaree).toBeCloseTo(27147.0);
    expect(article1.quantite).toBeCloseTo(354.0);
    expect(article1.unite).toBe('NOMBRE');
    expect(article1.totalArticle).toBeCloseTo(5511.0);
    expect(article1.taxes).toEqual([
      { code: '000110', assiette: 27147.0, taux: 0.0, montant: 0.0 },
      { code: '007217', assiette: 27147.0, taux: 0.25, montant: 68.0 },
      { code: '002109', assiette: 27215.0, taux: 20.0, montant: 5443.0 },
    ]);

    expect(article2.numero).toBe(2);
    expect(article2.valeurDeclaree).toBeCloseTo(12892.99);
    expect(article2.quantite).toBeCloseTo(200.0);
    expect(article2.totalArticle).toBeCloseTo(7260.0);
    expect(article2.taxes.map((t) => t.code)).toEqual(['000110', '007217', '002109']);
  });

  it('tolerates real-world OCR spacing noise in header labels ("BE N°" instead of "B E N°", REDEVABLE/CODE merged on one line)', () => {
    // Mirrors actual Tesseract output observed on a real scanned document: the
    // REDEVABLE and CODE labels sit side-by-side on the source page and get
    // merged onto one OCR'd line, and "B E N°" is frequently OCR'd as "BE N°"
    // (missing the space between B and E).
    const text = `Type Intervenant : Operateur
REDEVABLE : GLOBAL TRADE LOGISTICS SARL CODE : 500001
CATEGORIE D'ORDONNANCEMENT : Crédit d'enlèvement BE N° : 501 DU : 25/06/2026

ARTICLE  : 1              NUMERO SH : 6109100010     VALEUR :   27 147,00
QUANTITE : 354.000                UNITE : NOMBRE

TAXE   ! ASSIETTE  ! TAUX ! S.TVA ! S.FR ! TAUX VIRTUEL !  MONTANT
! 000110 !  27147.00 !  0.0 !   T   !      !              !     0,00 !
TOTAL ARTICLE :          5 511,00
`;
    const result = parseLiquidation(text);
    expect(result.header).toEqual({
      code: '500001',
      redevable: 'GLOBAL TRADE LOGISTICS SARL',
      benNumero: '501',
    });
  });

  it('throws when no articles are found', () => {
    expect(() => parseLiquidation('CODE : 123\nREDEVABLE : X\nB E N° : 1')).toThrow(
      'No articles found'
    );
  });

  it('throws mentioning the article number when a required article field is missing', () => {
    // No "QUANTITE :" line for article 1.
    const text = `CODE : 123
REDEVABLE : X
B E N° : 1

ARTICLE  : 1              NUMERO SH : 6109100010     VALEUR :   27 147,00
UNITE : NOMBRE

TAXE   ! ASSIETTE  ! TAUX ! S.TVA ! S.FR ! TAUX VIRTUEL !  MONTANT
! 000110 !  27147.00 !  0.0 !   T   !      !              !     0,00 !
TOTAL ARTICLE :          5 511,00
`;
    expect(() => parseLiquidation(text)).toThrow(/Article 1/);
  });

  it('throws "no tax rows found" when an article has zero tax rows', () => {
    const text = `CODE : 123
REDEVABLE : X
B E N° : 1

ARTICLE  : 1              NUMERO SH : 6109100010     VALEUR :   27 147,00
QUANTITE : 354.000                UNITE : NOMBRE

TAXE   ! ASSIETTE  ! TAUX ! S.TVA ! S.FR ! TAUX VIRTUEL !  MONTANT
TOTAL ARTICLE :          5 511,00
`;
    expect(() => parseLiquidation(text)).toThrow(/no tax rows found/i);
  });

  it('throws rather than silently dropping a malformed tax row (wrong column count)', () => {
    const text = `CODE : 123
REDEVABLE : X
B E N° : 1

ARTICLE  : 1              NUMERO SH : 6109100010     VALEUR :   27 147,00
QUANTITE : 354.000                UNITE : NOMBRE

TAXE   ! ASSIETTE  ! TAUX ! S.TVA ! S.FR ! TAUX VIRTUEL !  MONTANT
! 000110 !  27147.00 !  0.0 !     0,00 !
TOTAL ARTICLE :          5 511,00
`;
    expect(() => parseLiquidation(text)).toThrow(/Malformed tax row/i);
  });

  it('throws rather than silently dropping a line starting with "!" that is not a recognizable tax row or header', () => {
    const text = `CODE : 123
REDEVABLE : X
B E N° : 1

ARTICLE  : 1              NUMERO SH : 6109100010     VALEUR :   27 147,00
QUANTITE : 354.000                UNITE : NOMBRE

! 0011 ! garbled OCR line !
! 000110 !  27147.00 !  0.0 !   T   !      !              !     0,00 !
TOTAL ARTICLE :          5 511,00
`;
    expect(() => parseLiquidation(text)).toThrow(/Malformed tax row/i);
  });

  it('skips OCR border noise ("!" wrapped around label lines like QUANTITE/TOTAL ARTICLE) rather than throwing', () => {
    // Mirrors real Tesseract output: the table's box-drawing border is
    // frequently misread as a stray "!" wrapping the QUANTITE and
    // TOTAL ARTICLE lines, even though those aren't tax-row data.
    const text = `CODE : 123
REDEVABLE : X
B E N° : 1

ARTICLE  : 1              NUMERO SH : 6109100010     VALEUR :   27 147,00
! QUANTITE : 354.000                UNITE : NOMBRE !

TAXE   ! ASSIETTE  ! TAUX ! S.TVA ! S.FR ! TAUX VIRTUEL !  MONTANT
! 000110 !  27147.00 !  0.0 !   T   !      !              !     0,00 !
! TOTAL ARTICLE :          5 511,00 |
`;
    const result = parseLiquidation(text);
    expect(result.articles).toHaveLength(1);
    expect(result.articles[0].quantite).toBeCloseTo(354.0);
    expect(result.articles[0].totalArticle).toBeCloseTo(5511.0);
    expect(result.articles[0].taxes).toEqual([
      { code: '000110', assiette: 27147.0, taux: 0.0, montant: 0.0 },
    ]);
  });

  // Note: the existing "parses header and both articles from the real sample document" test
  // already exercises the (?<!TOTAL ) lookbehind implicitly, since the fixture contains
  // "TOTAL ARTICLE :" lines between article blocks and correctly yields exactly 2 articles
  // (not more, which is what would happen if "TOTAL ARTICLE :" were mistaken for a new
  // article-block boundary). No separate regression test is needed for this.
});

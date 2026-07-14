import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';
import { createUser } from '../../src/db/usersRepository.js';
import {
  saveDeclaration,
  listDeclarationsForUser,
  listAllDeclarations,
  getDeclarationById,
  getArticlesForDeclaration,
  searchDeclarationsByRedevable,
} from '../../src/db/declarationsRepository.js';
import type { Declaration } from '../../src/domain/types.js';

function makeDeclaration(): Declaration {
  return {
    code: '500001',
    redevable: 'GLOBAL TRADE LOGISTICS SARL',
    benNumero: '501',
    articles: [
      {
        numero: 1,
        hsCode: '6109100010',
        nomArticle: 'T-SHIRT',
        pays: 'ITALIE',
        paysCode: 'IT',
        valeurDeclaree: 27147,
        quantite: 354,
        unite: 'NOMBRE',
        taxes: [
          { code: '000110', assiette: 27147, taux: 0, montant: 0 },
          { code: '007217', assiette: 27147, taux: 0.25, montant: 68 },
          { code: '002109', assiette: 27215, taux: 20, montant: 5443 },
        ],
        totalArticle: 5511,
      },
    ],
  };
}

describe('declarationsRepository', () => {
  it('saves a declaration with its articles and reads it back', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'admin1', 'pw', 'admin');
    const declaration = makeDeclaration();

    const id = saveDeclaration(db, {
      ownerUserId: user.id,
      declaration,
      shipmentCostFields: {
        devise: 'EUR',
        montantFacture: 2981.34,
        tauxChange: 10.6675,
        fret: 7467.0,
        assurance: 118.0,
        valeurTotaleDeclaree: 40039.992,
      },
      articleCosts: [{ numero: 1, costPerUnit: 42.5 }],
      totalLandedCost: 15045.0,
      costEstimatePartial: false,
      colisCount: 6,
      referenceInterne: 'PO-1234',
      excelFilePath: '/data/declaration-1.xlsx',
    });

    expect(id).toBeGreaterThan(0);

    const saved = getDeclarationById(db, id);
    expect(saved).toMatchObject({
      id,
      ownerUserId: user.id,
      code: '500001',
      redevable: 'GLOBAL TRADE LOGISTICS SARL',
      valeurTotaleDeclaree: 40039.992,
      totalLandedCost: 15045.0,
      costEstimatePartial: false,
      colisCount: 6,
      referenceInterne: 'PO-1234',
      excelFilePath: '/data/declaration-1.xlsx',
    });

    const articleRows = db
      .prepare('SELECT * FROM declaration_articles WHERE declaration_id = ?')
      .all(id) as Array<{
        numero: number;
        hs_code: string;
        nom_article: string;
        pays: string;
        valeur_declaree: number;
        quantite: number;
        total_article: number;
        cost_per_unit: number;
        taxes_json: string;
      }>;

    expect(articleRows).toHaveLength(1);
    expect(articleRows[0]).toMatchObject({
      numero: 1,
      hs_code: '6109100010',
      nom_article: 'T-SHIRT',
      pays: 'ITALIE',
      valeur_declaree: 27147,
      quantite: 354,
      total_article: 5511,
      cost_per_unit: 42.5,
    });
    expect(JSON.parse(articleRows[0].taxes_json)).toEqual([
      { code: '000110', assiette: 27147, taux: 0, montant: 0 },
      { code: '007217', assiette: 27147, taux: 0.25, montant: 68 },
      { code: '002109', assiette: 27215, taux: 20, montant: 5443 },
    ]);

    db.close();
  });

  it('returns undefined from getDeclarationById for a nonexistent id', () => {
    const db = createDatabase(':memory:');
    expect(getDeclarationById(db, 999)).toBeUndefined();
    db.close();
  });

  it('throws and saves nothing if an article is missing its computed cost', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'admin1', 'pw', 'admin');
    const declaration = makeDeclaration();

    expect(() =>
      saveDeclaration(db, {
        ownerUserId: user.id,
        declaration,
        shipmentCostFields: {},
        articleCosts: [], // missing cost for article 1
        totalLandedCost: 0,
        costEstimatePartial: true,
        excelFilePath: '/data/declaration-x.xlsx',
      })
    ).toThrow('Missing computed cost for article 1');

    expect(listAllDeclarations(db)).toHaveLength(0); // transaction rolled back

    db.close();
  });

  it('lists declarations scoped to their owner, and lists all declarations for the superadmin view', () => {
    const db = createDatabase(':memory:');
    const alice = createUser(db, 'alice', 'pw', 'admin');
    const bob = createUser(db, 'bob', 'pw', 'admin');

    saveDeclaration(db, {
      ownerUserId: alice.id,
      declaration: makeDeclaration(),
      shipmentCostFields: {},
      articleCosts: [{ numero: 1, costPerUnit: 10 }],
      totalLandedCost: 100,
      costEstimatePartial: true,
      excelFilePath: '/data/a.xlsx',
    });
    saveDeclaration(db, {
      ownerUserId: bob.id,
      declaration: makeDeclaration(),
      shipmentCostFields: {},
      articleCosts: [{ numero: 1, costPerUnit: 20 }],
      totalLandedCost: 200,
      costEstimatePartial: true,
      excelFilePath: '/data/b.xlsx',
    });

    expect(listDeclarationsForUser(db, alice.id)).toHaveLength(1);
    expect(listDeclarationsForUser(db, alice.id)[0].excelFilePath).toBe('/data/a.xlsx');
    expect(listDeclarationsForUser(db, bob.id)).toHaveLength(1);
    expect(listAllDeclarations(db)).toHaveLength(2);

    db.close();
  });

  it('reads back a saved declaration\'s per-article cost breakdown via getArticlesForDeclaration', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'admin1', 'pw', 'admin');
    const declaration = makeDeclaration();

    const id = saveDeclaration(db, {
      ownerUserId: user.id,
      declaration,
      shipmentCostFields: {},
      articleCosts: [{ numero: 1, costPerUnit: 42.5 }],
      totalLandedCost: 15045.0,
      costEstimatePartial: true,
      excelFilePath: '/data/declaration-1.xlsx',
    });

    const articles = getArticlesForDeclaration(db, id);
    expect(articles).toEqual([
      {
        numero: 1,
        hsCode: '6109100010',
        nomArticle: 'T-SHIRT',
        pays: 'ITALIE',
        valeurDeclaree: 27147,
        quantite: 354,
        totalArticle: 5511,
        costPerUnit: 42.5,
      },
    ]);

    db.close();
  });

  it('returns an empty array from getArticlesForDeclaration for a declaration with no articles saved', () => {
    const db = createDatabase(':memory:');
    expect(getArticlesForDeclaration(db, 999)).toEqual([]);
    db.close();
  });

  it('searches declarations by a case-insensitive substring match on redevable, including computed total taxes', () => {
    const db = createDatabase(':memory:');
    const alice = createUser(db, 'alice', 'pw', 'admin');

    saveDeclaration(db, {
      ownerUserId: alice.id,
      declaration: { ...makeDeclaration(), redevable: 'Global Trade Logistics SARL' },
      shipmentCostFields: { valeurTotaleDeclaree: 40039.992 },
      articleCosts: [{ numero: 1, costPerUnit: 42.5 }],
      totalLandedCost: 15045.0,
      costEstimatePartial: false,
      excelFilePath: '/data/a.xlsx',
    });
    saveDeclaration(db, {
      ownerUserId: alice.id,
      declaration: { ...makeDeclaration(), redevable: 'Another Company SARL' },
      shipmentCostFields: {},
      articleCosts: [{ numero: 1, costPerUnit: 10 }],
      totalLandedCost: 100,
      costEstimatePartial: true,
      excelFilePath: '/data/b.xlsx',
    });

    const results = searchDeclarationsByRedevable(db, 'global trade');

    expect(results).toHaveLength(1);
    expect(results[0]).toMatchObject({
      redevable: 'Global Trade Logistics SARL',
      valeurTotaleDeclaree: 40039.992,
      totalLandedCost: 15045.0,
      totalTaxes: 5511, // 0 + 68 + 5443, from makeDeclaration()'s article taxes
    });

    db.close();
  });

  it('returns an empty array from searchDeclarationsByRedevable when nothing matches', () => {
    const db = createDatabase(':memory:');
    const alice = createUser(db, 'alice', 'pw', 'admin');
    saveDeclaration(db, {
      ownerUserId: alice.id,
      declaration: makeDeclaration(),
      shipmentCostFields: {},
      articleCosts: [{ numero: 1, costPerUnit: 10 }],
      totalLandedCost: 100,
      costEstimatePartial: true,
      excelFilePath: '/data/a.xlsx',
    });

    expect(searchDeclarationsByRedevable(db, 'nonexistent company')).toEqual([]);
    db.close();
  });
});

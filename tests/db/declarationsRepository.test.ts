import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';
import { createUser } from '../../src/db/usersRepository.js';
import {
  saveDeclaration,
  listDeclarationsForUser,
  listAllDeclarations,
  getDeclarationById,
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
      totalLandedCost: 15045.0,
      costEstimatePartial: false,
      colisCount: 6,
      referenceInterne: 'PO-1234',
      excelFilePath: '/data/declaration-1.xlsx',
    });

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
});

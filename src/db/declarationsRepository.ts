import type Database from 'better-sqlite3';
import type { Declaration } from '../domain/types.js';

export interface ShipmentCostFields {
  devise?: string;
  montantFacture?: number;
  tauxChange?: number;
  fret?: number;
  assurance?: number;
  valeurTotaleDeclaree?: number;
}

export interface ArticleCost {
  numero: number;
  costPerUnit: number;
}

export interface SaveDeclarationInput {
  ownerUserId: number;
  declaration: Declaration;
  shipmentCostFields: ShipmentCostFields;
  articleCosts: ArticleCost[];
  totalLandedCost: number;
  costEstimatePartial: boolean;
  colisCount?: number;
  referenceInterne?: string;
  excelFilePath: string;
}

export interface SavedDeclarationSummary {
  id: number;
  ownerUserId: number;
  code: string;
  redevable: string;
  totalLandedCost: number;
  costEstimatePartial: boolean;
  colisCount: number | null;
  referenceInterne: string | null;
  excelFilePath: string;
  createdAt: string;
}

interface DeclarationRow {
  id: number;
  owner_user_id: number;
  code: string;
  redevable: string;
  total_landed_cost: number;
  cost_estimate_partial: number;
  colis_count: number | null;
  reference_interne: string | null;
  excel_file_path: string;
  created_at: string;
}

function rowToSummary(row: DeclarationRow): SavedDeclarationSummary {
  return {
    id: row.id,
    ownerUserId: row.owner_user_id,
    code: row.code,
    redevable: row.redevable,
    totalLandedCost: row.total_landed_cost,
    costEstimatePartial: row.cost_estimate_partial === 1,
    colisCount: row.colis_count,
    referenceInterne: row.reference_interne,
    excelFilePath: row.excel_file_path,
    createdAt: row.created_at,
  };
}

export function saveDeclaration(db: Database.Database, input: SaveDeclarationInput): number {
  const createdAt = new Date().toISOString();
  const insertDeclaration = db.prepare(`
    INSERT INTO declarations (
      owner_user_id, code, redevable, ben_numero, devise, montant_facture, taux_change,
      fret, assurance, valeur_totale_declaree, colis_count, reference_interne,
      total_landed_cost, cost_estimate_partial, excel_file_path, created_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const insertArticle = db.prepare(`
    INSERT INTO declaration_articles (
      declaration_id, numero, hs_code, nom_article, pays, valeur_declaree, quantite,
      total_article, cost_per_unit, taxes_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const run = db.transaction((): number => {
    const result = insertDeclaration.run(
      input.ownerUserId,
      input.declaration.code,
      input.declaration.redevable,
      input.declaration.benNumero,
      input.shipmentCostFields.devise ?? null,
      input.shipmentCostFields.montantFacture ?? null,
      input.shipmentCostFields.tauxChange ?? null,
      input.shipmentCostFields.fret ?? null,
      input.shipmentCostFields.assurance ?? null,
      input.shipmentCostFields.valeurTotaleDeclaree ?? null,
      input.colisCount ?? null,
      input.referenceInterne ?? null,
      input.totalLandedCost,
      input.costEstimatePartial ? 1 : 0,
      input.excelFilePath,
      createdAt
    );
    const declarationId = Number(result.lastInsertRowid);

    for (const article of input.declaration.articles) {
      const cost = input.articleCosts.find((c) => c.numero === article.numero);
      if (!cost) {
        throw new Error(`Missing computed cost for article ${article.numero}`);
      }
      insertArticle.run(
        declarationId,
        article.numero,
        article.hsCode,
        article.nomArticle,
        article.pays,
        article.valeurDeclaree,
        article.quantite,
        article.totalArticle,
        cost.costPerUnit,
        JSON.stringify(article.taxes)
      );
    }

    return declarationId;
  });

  return run();
}

export function listDeclarationsForUser(
  db: Database.Database,
  ownerUserId: number
): SavedDeclarationSummary[] {
  const rows = db
    .prepare('SELECT * FROM declarations WHERE owner_user_id = ? ORDER BY created_at DESC')
    .all(ownerUserId) as DeclarationRow[];
  return rows.map(rowToSummary);
}

export function listAllDeclarations(db: Database.Database): SavedDeclarationSummary[] {
  const rows = db
    .prepare('SELECT * FROM declarations ORDER BY created_at DESC')
    .all() as DeclarationRow[];
  return rows.map(rowToSummary);
}

export function getDeclarationById(
  db: Database.Database,
  id: number
): SavedDeclarationSummary | undefined {
  const row = db.prepare('SELECT * FROM declarations WHERE id = ?').get(id) as
    | DeclarationRow
    | undefined;
  return row ? rowToSummary(row) : undefined;
}

export interface SavedArticleCost {
  numero: number;
  hsCode: string;
  nomArticle: string;
  pays: string;
  valeurDeclaree: number;
  quantite: number;
  totalArticle: number;
  costPerUnit: number;
}

interface ArticleRow {
  id: number;
  declaration_id: number;
  numero: number;
  hs_code: string;
  nom_article: string;
  pays: string;
  valeur_declaree: number;
  quantite: number;
  total_article: number;
  cost_per_unit: number;
  taxes_json: string;
}

function rowToArticleCost(row: ArticleRow): SavedArticleCost {
  return {
    numero: row.numero,
    hsCode: row.hs_code,
    nomArticle: row.nom_article,
    pays: row.pays,
    valeurDeclaree: row.valeur_declaree,
    quantite: row.quantite,
    totalArticle: row.total_article,
    costPerUnit: row.cost_per_unit,
  };
}

export function getArticlesForDeclaration(
  db: Database.Database,
  declarationId: number
): SavedArticleCost[] {
  const rows = db
    .prepare('SELECT * FROM declaration_articles WHERE declaration_id = ? ORDER BY numero ASC')
    .all(declarationId) as ArticleRow[];
  return rows.map(rowToArticleCost);
}

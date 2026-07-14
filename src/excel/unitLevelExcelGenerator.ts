import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { allocateTaxAcrossUnits, unionTaxCodes } from './unitLevelTaxHelpers.js';
import { styleHeaderRow, styleDataRow } from './excelStyling.js';

// Shared by the standalone File 2 generator below and by the combined
// (single-file, multi-sheet) generator — both need to add this exact sheet
// to a workbook writer they own, so the sheet-building logic lives here once.
// In the combined workbook this sheet is named "Global": every article's
// unit rows, one after another, so an admin doesn't have to flip between
// per-article sheets to see everything at once.
export function addUnitLevelSheet(
  workbook: ExcelJS.stream.xlsx.WorkbookWriter,
  declaration: Declaration,
  sheetName = 'Unit Detail'
): void {
  const taxCodes = unionTaxCodes(declaration.articles);
  // Valeur Déclarée is the last column, after every tax code column.
  const columnCount = 3 + taxCodes.length + 1;
  const valeurDeclareeColumn = columnCount;
  const moneyColumns = new Set<number>([
    valeurDeclareeColumn,
    ...taxCodes.map((_, i) => 4 + i),
  ]);

  const sheet = workbook.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 1 }] });

  sheet.columns = [
    { key: 'nomArticle', width: 30 },
    { key: 'hsCode', width: 15 },
    { key: 'serialNumber', width: 15 },
    ...taxCodes.map((code) => ({ key: code, width: 14 })),
    { key: 'valeurDeclaree', width: 16 },
  ];

  const headerRow = sheet.addRow([
    'Nom Article',
    'HSC',
    'Serial Number',
    ...taxCodes,
    'Valeur Déclarée',
  ]);
  styleHeaderRow(headerRow, columnCount);
  headerRow.commit();

  let dataRowIndex = 0;
  for (const article of declaration.articles) {
    const quantite = Math.round(article.quantite);
    if (Math.abs(article.quantite - quantite) > 0.01) {
      throw new Error(
        `Article ${article.numero}: quantite (${article.quantite}) is not a whole number; cannot generate one row per unit`
      );
    }

    // Valeur Déclarée / Unité — the same per-unit value on every row of this
    // article, not a per-unit allocation that needs to reconcile back to a
    // total the way tax montants do (allocateTaxAcrossUnits handles that
    // reconciliation case; this is a plain division).
    const valeurDeclareePerUnit = article.valeurDeclaree / quantite;

    const perCodeAllocations = new Map<string, number[]>();
    for (const code of taxCodes) {
      const tax = article.taxes.find((t) => t.code === code);
      perCodeAllocations.set(
        code,
        tax ? allocateTaxAcrossUnits(tax.montant, quantite) : new Array(quantite).fill(0)
      );
    }

    for (let unit = 0; unit < quantite; unit++) {
      const rowValues: Record<string, string | number> = {
        nomArticle: article.nomArticle,
        hsCode: article.hsCode,
        serialNumber: unit + 1,
        valeurDeclaree: valeurDeclareePerUnit,
      };
      for (const code of taxCodes) {
        rowValues[code] = perCodeAllocations.get(code)![unit];
      }
      const row = sheet.addRow(rowValues);
      styleDataRow(row, columnCount, dataRowIndex, moneyColumns);
      row.commit();
      dataRowIndex++;
    }
  }

  sheet.commit();
}

// Excel sheet names: max 31 chars, and cannot contain \ / ? * [ ] :
const INVALID_SHEET_NAME_CHARS = /[\\/?*[\]:]/g;

function sheetNameForArticle(article: Declaration['articles'][number]): string {
  const raw = `${article.numero}-${article.nomArticle}`;
  const sanitized = raw.replace(INVALID_SHEET_NAME_CHARS, '').trim();
  return (sanitized || `Article ${article.numero}`).slice(0, 31);
}

// One sheet per product/article (instead of a single shared "Unit Detail"
// sheet mixing every article's rows together) — each sheet only has the tax
// code columns that article actually has, rather than the full
// declaration-wide union padded with zeros for codes it doesn't carry.
// Sheet names are prefixed with the article's numero, which domain
// validation guarantees is unique per declaration, so two articles can
// never produce the same sheet name even if they share a product name.
export function addPerArticleUnitSheets(
  workbook: ExcelJS.stream.xlsx.WorkbookWriter,
  declaration: Declaration
): void {
  for (const article of declaration.articles) {
    const sheetName = sheetNameForArticle(article);
    const taxCodes = article.taxes.map((tax) => tax.code).sort();
    const columnCount = 3 + taxCodes.length + 1;
    const valeurDeclareeColumn = columnCount;
    const moneyColumns = new Set<number>([
      valeurDeclareeColumn,
      ...taxCodes.map((_, i) => 4 + i),
    ]);

    const sheet = workbook.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 1 }] });
    sheet.columns = [
      { key: 'nomArticle', width: 30 },
      { key: 'hsCode', width: 15 },
      { key: 'serialNumber', width: 15 },
      ...taxCodes.map((code) => ({ key: code, width: 14 })),
      { key: 'valeurDeclaree', width: 16 },
    ];

    const headerRow = sheet.addRow([
      'Nom Article',
      'HSC',
      'Serial Number',
      ...taxCodes,
      'Valeur Déclarée',
    ]);
    styleHeaderRow(headerRow, columnCount);
    headerRow.commit();

    const quantite = Math.round(article.quantite);
    if (Math.abs(article.quantite - quantite) > 0.01) {
      throw new Error(
        `Article ${article.numero}: quantite (${article.quantite}) is not a whole number; cannot generate one row per unit`
      );
    }

    const valeurDeclareePerUnit = article.valeurDeclaree / quantite;

    const perCodeAllocations = new Map<string, number[]>();
    for (const tax of article.taxes) {
      perCodeAllocations.set(tax.code, allocateTaxAcrossUnits(tax.montant, quantite));
    }

    for (let unit = 0; unit < quantite; unit++) {
      const rowValues: Record<string, string | number> = {
        nomArticle: article.nomArticle,
        hsCode: article.hsCode,
        serialNumber: unit + 1,
        valeurDeclaree: valeurDeclareePerUnit,
      };
      for (const code of taxCodes) {
        rowValues[code] = perCodeAllocations.get(code)![unit];
      }
      const row = sheet.addRow(rowValues);
      styleDataRow(row, columnCount, unit, moneyColumns);
      row.commit();
    }

    sheet.commit();
  }
}

export async function generateUnitLevelExcel(
  declaration: Declaration,
  outputPath: string
): Promise<void> {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: true,
  });
  addUnitLevelSheet(workbook, declaration);
  await workbook.commit();
}

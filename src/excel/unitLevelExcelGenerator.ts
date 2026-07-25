import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { allocateTaxAcrossUnits, taxCodeDesignation, unionTaxCodes } from './unitLevelTaxHelpers.js';
import {
  styleDataRow,
  styleHeaderRowGrouped,
  addSheetTitleRows,
  addGroupSeparatorBorder,
  resolveBrandArgb,
  resolveBrandDarkArgb,
  resolveCompanyName,
  resolveDocumentTitle,
  type BrandingInfo,
  type ColumnGroup,
} from './excelStyling.js';

function unitSheetColumnGroups(columnCount: number, taxCodeCount: number): ColumnGroup[] {
  return [
    { kind: 'identity', from: 1, to: 5 }, // Nom Article, Date d'arrivée, N° Enregistrement, HSC, Serial Number
    { kind: 'tax', from: 6, to: 5 + taxCodeCount },
    { kind: 'value', from: columnCount - 1, to: columnCount }, // Valeur Déclarée, Prorata
  ];
}

// Shared by the standalone File 2 generator below and by the combined
// (single-file, multi-sheet) generator — both need to add this exact sheet
// to a workbook they own, so the sheet-building logic lives here once.
// In the combined workbook this sheet is named "Global": every article's
// unit rows, one after another, so an admin doesn't have to flip between
// per-article sheets to see everything at once.
export async function addUnitLevelSheet(
  workbook: ExcelJS.Workbook,
  declaration: Declaration,
  branding: BrandingInfo,
  sheetName = 'Unit Detail'
): Promise<void> {
  const taxCodes = unionTaxCodes(declaration.articles);
  // Rubriques from the Liquidation's RECAPITULATION table that never appear
  // in any article's own tax rows (e.g. REDV.INF., RI SEGMA, REMISES CREDIT)
  // — only known as one declaration-wide montant, so each unit row gets its
  // share via montant × Prorata rather than allocateTaxAcrossUnits.
  const extraOrdonnancementTaxes = declaration.ordonnancementTaxes.filter(
    (tax) => !taxCodes.includes(tax.code)
  );
  const extraCodes = extraOrdonnancementTaxes.map((tax) => tax.code);
  const allTaxCodeCount = taxCodes.length + extraCodes.length;
  // "Date d'arrivée" and "N° Enregistrement" (both from the DUM, same value
  // on every row) sit right before HSC, ahead of the tax columns. Valeur
  // Déclarée and Prorata are the last two columns, after every tax code
  // column (per-article union, then extra ordonnancement-only codes).
  const columnCount = 5 + allTaxCodeCount + 2;
  const valeurDeclareeColumn = columnCount - 1;
  const prorataColumn = columnCount;
  const moneyColumns = new Set<number>([
    valeurDeclareeColumn,
    ...taxCodes.map((_, i) => 6 + i),
    ...extraCodes.map((_, i) => 6 + taxCodes.length + i),
  ]);
  const percentColumns = new Set<number>([prorataColumn]);
  const declarationValeurDeclareeTotal = declaration.articles.reduce(
    (sum, article) => sum + article.valeurDeclaree,
    0
  );

  const sheet = workbook.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 3 }] });

  sheet.columns = [
    { key: 'nomArticle', width: 36 },
    { key: 'dateArrivee', width: 18 },
    { key: 'numeroEnregistrement', width: 26 },
    { key: 'hsCode', width: 20 },
    { key: 'serialNumber', width: 18 },
    ...taxCodes.map((code) => ({ key: code, width: 24 })),
    ...extraCodes.map((code) => ({ key: code, width: 24 })),
    { key: 'valeurDeclaree', width: 22 },
    { key: 'prorata', width: 18 },
  ];

  await addSheetTitleRows(
    workbook,
    sheet,
    columnCount,
    resolveCompanyName(branding.companyName),
    resolveDocumentTitle(declaration),
    resolveBrandArgb(branding.brandColor),
    resolveBrandDarkArgb(branding.brandColor),
    branding.logoDataUri
  );

  const headerRow = sheet.addRow([
    'Nom Article',
    "Date d'arrivée",
    'N° Enregistrement',
    'HSC',
    'Serial Number',
    ...taxCodes.map(taxCodeDesignation),
    ...extraOrdonnancementTaxes.map((tax) => tax.designation),
    'Valeur Déclarée',
    'Prorata',
  ]);
  styleHeaderRowGrouped(headerRow, columnCount, unitSheetColumnGroups(columnCount, allTaxCodeCount));

  let dataRowIndex = 0;
  declaration.articles.forEach((article, articleIndex) => {
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
    // Prorata — this unit's share of the whole declaration's total declared
    // value (montant in Valeur Déclarée / the sum of every article's Valeur
    // Déclarée across the declaration), not just its own product's total.
    const prorata = valeurDeclareePerUnit / declarationValeurDeclareeTotal;

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
        dateArrivee: declaration.dateArrivee ?? '',
        numeroEnregistrement: declaration.numeroEnregistrement ?? '',
        hsCode: article.hsCode,
        serialNumber: unit + 1,
        valeurDeclaree: valeurDeclareePerUnit,
        prorata,
      };
      for (const code of taxCodes) {
        rowValues[code] = perCodeAllocations.get(code)![unit];
      }
      for (const tax of extraOrdonnancementTaxes) {
        rowValues[tax.code] = tax.montant * prorata;
      }
      const row = sheet.addRow(rowValues);
      styleDataRow(row, columnCount, dataRowIndex, moneyColumns, percentColumns);
      // A thicker top border marks where each new product's block of unit
      // rows begins, so it's visually obvious where one product ends and
      // the next starts in this combined sheet — skipped for the very
      // first product (nothing to separate it from).
      if (unit === 0 && articleIndex > 0) {
        addGroupSeparatorBorder(row, columnCount);
      }
      dataRowIndex++;
    }
  });
}

export async function generateUnitLevelExcel(
  declaration: Declaration,
  outputPath: string,
  branding: BrandingInfo
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  await addUnitLevelSheet(workbook, declaration, branding);
  await workbook.xlsx.writeFile(outputPath);
}

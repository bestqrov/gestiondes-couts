import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import {
  allocateTaxAcrossUnits,
  hsAndOriginKey,
  taxCodeDesignation,
  unionTaxCodes,
  valeurDeclareeTotalsByHsAndOrigin,
} from './unitLevelTaxHelpers.js';
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

// Pays gets its own green highlight (header + every data cell), distinct
// from the rest of the indigo "identity" group, so it stands out as the key
// (alongside HSC) that the "HS total" sheet's tax matching is now built on.
const PAYS_COLUMN = 8;
const PAYS_HEADER_ARGB = 'FF16A34A';
const PAYS_CELL_ARGB = 'FFBBF7D0';

function unitSheetColumnGroups(columnCount: number, taxCodeCount: number): ColumnGroup[] {
  return [
    // Nom Article, DONNEES COMPTABLES, Poids net (kg), Date d'arrivée,
    // Nature et numéro du titre de transport, N° Enregistrement,
    // Date de déclaration, Pays, HSC, Serial Number
    { kind: 'identity', from: 1, to: 10 },
    { kind: 'tax', from: 11, to: 10 + taxCodeCount },
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
  generatedAt: Date,
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
  // DONNEES COMPTABLES, Poids net (kg), Date d'arrivée, Nature et numéro du
  // titre de transport, N° Enregistrement, and Date de déclaration (from the
  // DUM) sit right before Pays and HSC, ahead of the tax columns — Poids net
  // and Pays are per-article, the rest are the same value on every row.
  // Valeur Déclarée and Prorata are the last two columns, after every tax
  // code column (per-article union, then extra ordonnancement-only codes).
  const columnCount = 10 + allTaxCodeCount + 2;
  const valeurDeclareeColumn = columnCount - 1;
  const prorataColumn = columnCount;
  const poidsNetColumn = 3;
  const moneyColumns = new Set<number>([
    poidsNetColumn,
    valeurDeclareeColumn,
    ...taxCodes.map((_, i) => 11 + i),
    ...extraCodes.map((_, i) => 11 + taxCodes.length + i),
  ]);
  const percentColumns = new Set<number>([prorataColumn]);
  const valeurDeclareeTotalsByGroup = valeurDeclareeTotalsByHsAndOrigin(declaration.articles);
  // Declaration-wide-only rubriques (REDV.INF., RI SEGMA, REMISES CREDIT —
  // extraOrdonnancementTaxes) only have a single montant for the whole
  // declaration, not one per HS+origin group, so they must keep being spread
  // with a share that sums to 1 across every row of the declaration (not per
  // group) or a declaration with more than one HS+origin group would apply
  // their montant once per group instead of once, total.
  const declarationValeurDeclareeTotal = declaration.articles.reduce(
    (sum, article) => sum + article.valeurDeclaree,
    0
  );

  const sheet = workbook.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 4 }] });

  sheet.columns = [
    { key: 'nomArticle', width: 36 },
    { key: 'donneesComptables', width: 22 },
    { key: 'poidsNet', width: 18 },
    { key: 'dateArrivee', width: 18 },
    { key: 'titreTransport', width: 30 },
    { key: 'numeroEnregistrement', width: 26 },
    { key: 'dateDeclaration', width: 26 },
    { key: 'pays', width: 20 },
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
    branding.logoDataUri,
    generatedAt
  );

  const headerRow = sheet.addRow([
    'Nom Article',
    'DONNEES COMPTABLES',
    'Poids net (kg)',
    "Date d'arrivée",
    'Nature et numéro du titre de transport',
    'N° Enregistrement',
    'Date de déclaration',
    'Pays',
    'HSC',
    'Serial Number',
    ...taxCodes.map(taxCodeDesignation),
    ...extraOrdonnancementTaxes.map((tax) => tax.designation),
    'Valeur Déclarée',
    'Prorata',
  ]);
  styleHeaderRowGrouped(headerRow, columnCount, unitSheetColumnGroups(columnCount, allTaxCodeCount));
  const paysHeaderCell = headerRow.getCell(PAYS_COLUMN);
  paysHeaderCell.style = {
    ...paysHeaderCell.style,
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: PAYS_HEADER_ARGB } },
  };

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
    // Poids net (kg) / Unités complémentaires (DUM fields 33 / 32) — same
    // per-unit treatment as Valeur Déclarée, not the raw article-wide poids
    // net repeated on every row. Falls back to the raw poidsNet if
    // unitesComplementaires is 0 (guards a divide-by-zero rather than
    // producing Infinity/NaN in the sheet).
    const poidsNetPerUnit =
      article.unitesComplementaires > 0 ? article.poidsNet / article.unitesComplementaires : article.poidsNet;
    // Prorata — this unit's share of its product group's total declared
    // value (montant in Valeur Déclarée / the sum of Valeur Déclarée across
    // every article sharing the same HS code prefix + country of origin),
    // not the whole declaration's total.
    const groupTotal = valeurDeclareeTotalsByGroup.get(hsAndOriginKey(article.hsCode, article.pays)) ?? 0;
    const prorata = groupTotal > 0 ? valeurDeclareePerUnit / groupTotal : 0;
    // Share used only to spread extraOrdonnancementTaxes' declaration-wide
    // montants (see above) — sums to 1 across the whole declaration, unlike
    // the displayed `prorata`, which sums to 1 within each HS+origin group.
    const declarationShare =
      declarationValeurDeclareeTotal > 0 ? valeurDeclareePerUnit / declarationValeurDeclareeTotal : 0;

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
        donneesComptables: declaration.donneesComptables ?? '',
        poidsNet: poidsNetPerUnit,
        dateArrivee: declaration.dateArrivee ?? '',
        titreTransport: declaration.titreTransport ?? '',
        numeroEnregistrement: declaration.numeroEnregistrement ?? '',
        dateDeclaration: declaration.dateDeclaration ?? '',
        hsCode: article.hsCode,
        pays: article.pays,
        serialNumber: unit + 1,
        valeurDeclaree: valeurDeclareePerUnit,
        prorata,
      };
      for (const code of taxCodes) {
        rowValues[code] = perCodeAllocations.get(code)![unit];
      }
      for (const tax of extraOrdonnancementTaxes) {
        rowValues[tax.code] = tax.montant * declarationShare;
      }
      const row = sheet.addRow(rowValues);
      styleDataRow(row, columnCount, dataRowIndex, moneyColumns, percentColumns);
      const paysCell = row.getCell(PAYS_COLUMN);
      paysCell.style = {
        ...paysCell.style,
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: PAYS_CELL_ARGB } },
      };
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
  await addUnitLevelSheet(workbook, declaration, branding, new Date());
  await workbook.xlsx.writeFile(outputPath);
}

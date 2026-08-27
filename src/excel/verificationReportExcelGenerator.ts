import ExcelJS from 'exceljs';
import type { PackingListRow } from '../parser/packingList/packingListParser.js';
import type { Declaration } from '../domain/types.js';
import { hsAndOriginKey } from './unitLevelTaxHelpers.js';
import { createPackingListMatcher } from './packingListMatcher.js';
import {
  addSheetTitleRows,
  styleHeaderRow,
  styleDataRow,
  resolveBrandArgb,
  resolveBrandDarkArgb,
  resolveCompanyName,
  resolveDocumentTitle,
  type BrandingInfo,
} from './excelStyling.js';

const UNMATCHED_COLUMN_COUNT = 8;
const GAP_COLUMN_COUNT = 7;
const COLUMN_COUNT = Math.max(UNMATCHED_COLUMN_COUNT, GAP_COLUMN_COUNT);

// A section title row, spanning every column — same visual weight as a
// sub-header, distinct from both the letterhead above and the table
// headers below.
function addSectionTitleRow(sheet: ExcelJS.Worksheet, title: string): ExcelJS.Row {
  const row = sheet.addRow([title]);
  for (let col = 1; col <= COLUMN_COUNT; col++) {
    row.getCell(col).style = {
      font: { bold: true, size: 12, color: { argb: 'FF1E293B' } },
      alignment: { vertical: 'middle', horizontal: 'left' },
    };
  }
  row.height = 22;
  sheet.mergeCells(row.number, 1, row.number, COLUMN_COUNT);
  return row;
}

// Builds a small report of everything the HS total matching couldn't
// reconcile between the uploaded packing list and the declaration
// (DUM/Liquidation) — meant to surface likely data-entry mistakes (a wrong
// HS code, an origin that doesn't match any declaration article, a product
// missing from one side or the other) so they can be checked/corrected at
// the source, rather than silently showing up as a 0% or under-100% PRORATA
// in "HS total" with no explanation. Two independent sections, since a row
// can fail to match for reasons unrelated to a declaration group's own
// coverage gap:
//   1. Packing-list rows that matched no declaration article at all (wrong/
//      unknown HS code, or a real origin that doesn't correspond to any
//      declaration article for that HS position).
//   2. Declaration HS+origin groups whose matched packing-list rows don't
//      add up to the article's own declared quantite — some of the
//      declared units simply aren't itemized in the uploaded packing list.
export async function generateVerificationReportExcel(
  declaration: Declaration,
  packingListRows: PackingListRow[],
  outputPath: string,
  branding: BrandingInfo
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const generatedAt = new Date();
  const sheet = workbook.addWorksheet('À vérifier');

  await addSheetTitleRows(
    workbook,
    sheet,
    COLUMN_COUNT,
    resolveCompanyName(branding.companyName),
    resolveDocumentTitle(declaration),
    resolveBrandArgb(branding.brandColor),
    resolveBrandDarkArgb(branding.brandColor),
    branding.logoDataUri,
    generatedAt
  );

  const matcher = createPackingListMatcher(declaration);
  const matchedPiecesByGroup = new Map<string, number>();
  const unmatchedRows: PackingListRow[] = [];
  for (const row of packingListRows) {
    const match = matcher.resolveRowMatch(row);
    if (!match) {
      unmatchedRows.push(row);
      continue;
    }
    const key = hsAndOriginKey(match.hsCode, match.pays);
    matchedPiecesByGroup.set(key, (matchedPiecesByGroup.get(key) ?? 0) + row.pieces);
  }

  addSectionTitleRow(
    sheet,
    `Lignes de la packing list non reconnues (${unmatchedRows.length}) — code SH ou pays d'origine ne correspondant à aucun article de la déclaration`
  );
  const unmatchedHeaderRow = sheet.addRow([
    'item',
    'DESCRIPTION',
    'color',
    'pieces',
    'unit',
    'total',
    'origin',
    'HS CODE',
  ]);
  styleHeaderRow(unmatchedHeaderRow, UNMATCHED_COLUMN_COUNT);
  unmatchedRows.forEach((row, index) => {
    const excelRow = sheet.addRow([
      row.item,
      row.description,
      row.color,
      row.pieces,
      row.unit,
      row.total,
      row.origin,
      row.hsCode,
    ]);
    styleDataRow(excelRow, UNMATCHED_COLUMN_COUNT, index, new Set([5, 6]));
  });
  if (unmatchedRows.length === 0) {
    const emptyRow = sheet.addRow(['Aucune — toutes les lignes de la packing list ont été reconnues.']);
    emptyRow.getCell(1).font = { italic: true, color: { argb: 'FF64748B' } };
    sheet.mergeCells(emptyRow.number, 1, emptyRow.number, UNMATCHED_COLUMN_COUNT);
  }

  sheet.addRow([]);

  const underCoveredGroups: {
    hsCode: string;
    pays: string;
    nomArticle: string;
    quantiteDeclaree: number;
    quantiteCouverte: number;
    ecart: number;
    valeurManquante: number;
  }[] = [];
  const seenGroups = new Set<string>();
  for (const article of declaration.articles) {
    const key = hsAndOriginKey(article.hsCode, article.pays);
    if (seenGroups.has(key)) continue;
    seenGroups.add(key);
    const quantiteDeclaree = matcher.quantiteTotalsByGroup.get(key) ?? 0;
    const quantiteCouverte = matchedPiecesByGroup.get(key) ?? 0;
    const ecart = quantiteDeclaree - quantiteCouverte;
    if (ecart <= 0) continue;
    const groupTotal = matcher.valeurDeclareeTotalsByGroup.get(key) ?? 0;
    const valeurManquante = quantiteDeclaree > 0 ? (ecart / quantiteDeclaree) * groupTotal : 0;
    underCoveredGroups.push({
      hsCode: article.hsCode,
      pays: article.pays,
      nomArticle: article.nomArticle,
      quantiteDeclaree,
      quantiteCouverte,
      ecart,
      valeurManquante,
    });
  }

  addSectionTitleRow(
    sheet,
    `Produits de la déclaration partiellement ou pas du tout couverts par la packing list (${underCoveredGroups.length})`
  );
  const gapHeaderRow = sheet.addRow([
    'HSC',
    'Pays',
    'Nom Article',
    'Quantité déclarée',
    'Quantité trouvée dans la packing list',
    'Écart (manquant)',
    'Valeur déclarée manquante',
  ]);
  styleHeaderRow(gapHeaderRow, GAP_COLUMN_COUNT);
  underCoveredGroups.forEach((group, index) => {
    const excelRow = sheet.addRow([
      group.hsCode,
      group.pays,
      group.nomArticle,
      group.quantiteDeclaree,
      group.quantiteCouverte,
      group.ecart,
      group.valeurManquante,
    ]);
    styleDataRow(excelRow, GAP_COLUMN_COUNT, index, new Set([7]));
  });
  if (underCoveredGroups.length === 0) {
    const emptyRow = sheet.addRow(['Aucun — chaque produit de la déclaration est entièrement couvert par la packing list.']);
    emptyRow.getCell(1).font = { italic: true, color: { argb: 'FF64748B' } };
    sheet.mergeCells(emptyRow.number, 1, emptyRow.number, GAP_COLUMN_COUNT);
  }

  sheet.columns = [
    { width: 24 },
    { width: 24 },
    { width: 16 },
    { width: 12 },
    { width: 12 },
    { width: 14 },
    { width: 16 },
    { width: 16 },
  ];

  await workbook.xlsx.writeFile(outputPath);
}

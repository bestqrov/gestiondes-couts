import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { normalizeCountryName } from '../domain/countryNames.js';
import {
  styleDataRow,
  styleHeaderRowGrouped,
  addSheetTitleRows,
  resolveBrandArgb,
  resolveBrandDarkArgb,
  resolveCompanyName,
  resolveDocumentTitle,
  type BrandingInfo,
  type ColumnGroup,
} from './excelStyling.js';

const COLUMN_COUNT = 4;
const PRORATA_COLUMN = 4;
const COLUMN_GROUPS: ColumnGroup[] = [
  { kind: 'identity', from: 1, to: 3 }, // product, origin, HSC
  { kind: 'value', from: 4, to: 4 }, // PRORATA
];

// Same HS "position" convention used everywhere else articles get grouped
// (declarationMerger's HS_CODE_COMPARISON_LENGTH, packingListExcelGenerator's
// HS_CODE_MATCH_LENGTH) — only the first 6 digits need to agree.
const HS_CODE_GROUP_LENGTH = 6;

function hsCodePrefix(code: string): string {
  return code.slice(0, HS_CODE_GROUP_LENGTH);
}

interface ProrataGroup {
  product: string;
  origin: string;
  hsCode: string;
  prorata: number;
}

// Groups every declaration article by HS position + country of origin (e.g.
// several color variants of the same product/origin collapse into one row),
// summing each group's combined share of the declaration's total declared
// value — the same "one row per HS+origin group, PRORATA is the group's
// summed share" grouping HS total's own PRORATA column uses.
function groupArticlesByHsAndOrigin(declaration: Declaration): ProrataGroup[] {
  const declarationValeurDeclareeTotal = declaration.articles.reduce(
    (sum, article) => sum + article.valeurDeclaree,
    0
  );
  const groups = new Map<string, ProrataGroup>();
  for (const article of declaration.articles) {
    const key = `${hsCodePrefix(article.hsCode)}|${normalizeCountryName(article.pays)}`;
    const share =
      declarationValeurDeclareeTotal > 0 ? article.valeurDeclaree / declarationValeurDeclareeTotal : 0;
    const existing = groups.get(key);
    if (existing) {
      existing.prorata += share;
    } else {
      groups.set(key, {
        product: article.nomArticle,
        origin: article.pays,
        hsCode: article.hsCode,
        prorata: share,
      });
    }
  }
  return [...groups.values()];
}

// Adds the "prorata" sheet — one row per declaration HS+origin group
// (product/origin/HSC taken from the first article encountered in that
// group), with PRORATA summing every article in the group's own share of
// the declaration's total declared value.
export async function addProrataSheet(
  workbook: ExcelJS.Workbook,
  declaration: Declaration,
  branding: BrandingInfo,
  generatedAt: Date,
  sheetName = 'prorata'
): Promise<void> {
  const sheet = workbook.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 4 }] });

  sheet.columns = [
    { key: 'product', width: 36 },
    { key: 'origin', width: 24 },
    { key: 'hsCode', width: 20 },
    { key: 'prorata', width: 18 },
  ];

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

  const headerRow = sheet.addRow(['product', 'origin', 'HSC', 'PRORATA']);
  styleHeaderRowGrouped(headerRow, COLUMN_COUNT, COLUMN_GROUPS);

  const groups = groupArticlesByHsAndOrigin(declaration);
  groups.forEach((group, index) => {
    // Rounded to 4 decimal places of the raw fraction, matching every other
    // PRORATA column's 0.00% display elsewhere in the workbook.
    const prorata = Math.round(group.prorata * 10000) / 10000;
    const row = sheet.addRow({
      product: group.product,
      origin: group.origin,
      hsCode: group.hsCode,
      prorata,
    });
    styleDataRow(row, COLUMN_COUNT, index, new Set(), new Set([PRORATA_COLUMN]));
  });
}

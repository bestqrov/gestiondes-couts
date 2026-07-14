import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { styleHeaderRow, styleDataRow } from './excelStyling.js';

const COLUMN_COUNT = 5;
const VALEUR_DECLAREE_COLUMN = 4;

// Shared by the standalone File 1 generator below and by the combined
// (single-file, multi-sheet) generator — both need to add this exact sheet
// to a workbook writer they own, so the sheet-building logic lives here once.
export function addArticleSummarySheet(
  workbook: ExcelJS.stream.xlsx.WorkbookWriter,
  declaration: Declaration
): void {
  const sheet = workbook.addWorksheet('Articles', { views: [{ state: 'frozen', ySplit: 1 }] });

  // Column widths only here (no `header:`) — the header row is added and
  // styled explicitly below instead of relying on ExcelJS's implicit
  // header-from-columns behavior, so it can be colored/bolded.
  sheet.columns = [
    { key: 'nomArticle', width: 30 },
    { key: 'hsCode', width: 15 },
    { key: 'pays', width: 20 },
    { key: 'valeurDeclaree', width: 18 },
    { key: 'quantite', width: 18 },
  ];

  const headerRow = sheet.addRow(['Nom Article', 'HSC', 'Pays', 'Valeur déclarée', 'Unité (Quantity)']);
  styleHeaderRow(headerRow, COLUMN_COUNT);
  headerRow.commit();

  declaration.articles.forEach((article, index) => {
    const row = sheet.addRow({
      nomArticle: article.nomArticle,
      hsCode: article.hsCode,
      pays: article.pays,
      valeurDeclaree: article.valeurDeclaree,
      quantite: article.quantite,
    });
    styleDataRow(row, COLUMN_COUNT, index, new Set([VALEUR_DECLAREE_COLUMN]));
    row.commit();
  });

  sheet.commit();
}

export async function generateArticleSummaryExcel(
  declaration: Declaration,
  outputPath: string
): Promise<void> {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: true,
  });
  addArticleSummarySheet(workbook, declaration);
  await workbook.commit();
}

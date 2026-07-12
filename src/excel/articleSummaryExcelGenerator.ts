import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';

// Shared by the standalone File 1 generator below and by the combined
// (single-file, multi-sheet) generator — both need to add this exact sheet
// to a workbook writer they own, so the sheet-building logic lives here once.
export function addArticleSummarySheet(
  workbook: ExcelJS.stream.xlsx.WorkbookWriter,
  declaration: Declaration
): void {
  const sheet = workbook.addWorksheet('Articles');

  sheet.columns = [
    { header: 'Nom Article', key: 'nomArticle', width: 30 },
    { header: 'HSC', key: 'hsCode', width: 15 },
    { header: 'Pays', key: 'pays', width: 20 },
    { header: 'Valeur déclarée', key: 'valeurDeclaree', width: 18 },
    { header: 'Unité (Quantity)', key: 'quantite', width: 18 },
  ];

  for (const article of declaration.articles) {
    sheet
      .addRow({
        nomArticle: article.nomArticle,
        hsCode: article.hsCode,
        pays: article.pays,
        valeurDeclaree: article.valeurDeclaree,
        quantite: article.quantite,
      })
      .commit();
  }

  sheet.commit();
}

export async function generateArticleSummaryExcel(
  declaration: Declaration,
  outputPath: string
): Promise<void> {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: false,
  });
  addArticleSummarySheet(workbook, declaration);
  await workbook.commit();
}

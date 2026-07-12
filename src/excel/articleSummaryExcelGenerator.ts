import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';

export async function generateArticleSummaryExcel(
  declaration: Declaration,
  outputPath: string
): Promise<void> {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: false,
  });
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
  await workbook.commit();
}

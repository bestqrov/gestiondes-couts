import ExcelJS from 'exceljs';
import type { Declaration } from '../domain/types.js';
import { allocateTaxAcrossUnits, unionTaxCodes } from './unitLevelTaxHelpers.js';

export async function generateUnitLevelExcel(
  declaration: Declaration,
  outputPath: string
): Promise<void> {
  const taxCodes = unionTaxCodes(declaration.articles);

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: false,
  });
  const sheet = workbook.addWorksheet('Unit Detail');

  sheet.columns = [
    { header: 'Nom Article', key: 'nomArticle', width: 30 },
    { header: 'HSC', key: 'hsCode', width: 15 },
    { header: 'Serial Number', key: 'serialNumber', width: 15 },
    ...taxCodes.map((code) => ({ header: code, key: code, width: 14 })),
  ];

  for (const article of declaration.articles) {
    const quantite = Math.round(article.quantite);
    if (Math.abs(article.quantite - quantite) > 0.01) {
      throw new Error(
        `Article ${article.numero}: quantite (${article.quantite}) is not a whole number; cannot generate one row per unit`
      );
    }

    const perCodeAllocations = new Map<string, number[]>();
    for (const code of taxCodes) {
      const tax = article.taxes.find((t) => t.code === code);
      perCodeAllocations.set(
        code,
        tax ? allocateTaxAcrossUnits(tax.montant, quantite) : new Array(quantite).fill(0)
      );
    }

    for (let unit = 0; unit < quantite; unit++) {
      const row: Record<string, string | number> = {
        nomArticle: article.nomArticle,
        hsCode: article.hsCode,
        serialNumber: unit + 1,
      };
      for (const code of taxCodes) {
        row[code] = perCodeAllocations.get(code)![unit];
      }
      sheet.addRow(row).commit();
    }
  }

  sheet.commit();
  await workbook.commit();
}

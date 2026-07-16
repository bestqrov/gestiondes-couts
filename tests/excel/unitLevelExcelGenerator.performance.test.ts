import ExcelJS from 'exceljs';
import { describe, it, expect, afterEach } from 'vitest';
import { generateUnitLevelExcel } from '../../src/excel/unitLevelExcelGenerator.js';
import { createTempXlsxPath, cleanupTempDir } from './testHelpers.js';
import type { Declaration } from '../../src/domain/types.js';

function makeLargeDeclaration(quantite: number): Declaration {
  return {
    code: '999999',
    redevable: 'PERFORMANCE TEST CO',
    benNumero: '1',
    articles: [
      {
        numero: 1,
        hsCode: '6109100010',
        nomArticle: 'T-SHIRT',
        pays: 'ITALIE',
        paysCode: 'IT',
        valeurDeclaree: 100000,
        quantite,
        unite: 'U',
        taxes: [
          { code: '000110', assiette: 100000, taux: 0, montant: 0 },
          { code: '007217', assiette: 100000, taux: 0.25, montant: 250 },
          { code: '002109', assiette: 100000, taux: 20, montant: 20000 },
        ],
        totalArticle: 20250,
      },
    ],
  };
}

describe('generateUnitLevelExcel performance', () => {
  let tempDir: string | undefined;

  afterEach(() => {
    if (tempDir) cleanupTempDir(tempDir);
    tempDir = undefined;
  });

  it('generates 10,000+ rows via streaming within a reasonable time and with the correct row count', async () => {
    const quantite = 10000;
    const declaration = makeLargeDeclaration(quantite);
    const { filePath, dir } = createTempXlsxPath('unit-level-performance');
    tempDir = dir;

    const start = Date.now();
    await generateUnitLevelExcel(declaration, filePath, { companyName: null, brandColor: null });
    const durationMs = Date.now() - start;

    expect(durationMs).toBeLessThan(15000);

    // Use the streaming reader (not a full in-memory read) to count rows,
    // consistent with this file existing to prove the streaming path scales.
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {});
    let rowCount = 0;
    for await (const worksheetReader of reader) {
      for await (const _row of worksheetReader) {
        rowCount++;
      }
    }

    // quantite data rows + 2 title rows + 1 header row
    expect(rowCount).toBe(quantite + 3);
  }, 30000);
});

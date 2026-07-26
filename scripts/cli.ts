import 'dotenv/config';
import { readFile } from 'node:fs/promises';
import { extractDocumentText } from '../src/ocr/documentTextExtractor.js';
import { detectAndParsePair } from '../src/parser/detectAndParsePair.js';
import { mergeDeclaration } from '../src/merge/declarationMerger.js';
import { validateArticle } from '../src/domain/validators.js';
import { generateCombinedExcel } from '../src/excel/combinedExcelGenerator.js';
import { parsePackingList } from '../src/parser/packingList/packingListParser.js';

async function main() {
  const [, , liquidationPath, dumPath, packingListPath, outDir = '.'] = process.argv;

  if (!liquidationPath || !dumPath || !packingListPath) {
    console.error(
      'Usage: npm run generate -- <liquidation-file> <dum-file> <packing-list-file> [output-dir]'
    );
    process.exit(1);
  }

  console.log(`Reading Liquidation: ${liquidationPath}`);
  const liquidationOcr = await extractDocumentText(liquidationPath);
  console.log(`Reading DUM: ${dumPath}`);
  const dumOcr = await extractDocumentText(dumPath);
  console.log(`Reading Packing List: ${packingListPath}`);
  const packingListBuffer = await readFile(packingListPath);
  const packingListRows = await parsePackingList(packingListBuffer);
  console.log(`Packing list: ${packingListRows.length} row(s)`);

  console.log('\n--- Liquidation extracted text (confidence %s) ---', liquidationOcr.confidence);
  console.log(liquidationOcr.text);
  console.log('\n--- DUM extracted text (confidence %s) ---', dumOcr.confidence);
  console.log(dumOcr.text);

  console.log('\n--- Parsing ---');
  const { liquidation, dum, swapped } = detectAndParsePair(liquidationOcr.text, dumOcr.text);
  if (swapped) {
    console.log('(Note: files were auto-detected in reversed order from the arguments given)');
  }

  console.log('--- Merging ---');
  const declaration = mergeDeclaration(liquidation, dum);
  for (const article of declaration.articles) {
    validateArticle(article);
  }
  console.log(`Merged declaration: code=${declaration.code}, ${declaration.articles.length} article(s)`);

  const outputPath = `${outDir}/Declaration.xlsx`;
  // No app settings context from a CLI run — falls back to the default
  // branding (generic company name, indigo accent).
  await generateCombinedExcel(declaration, packingListRows, outputPath, {
    companyName: null,
    brandColor: null,
    logoDataUri: null,
  });

  console.log(`\nGenerated: ${outputPath} (3 sheets: Articles, Global, HS total)`);
}

main().catch((error) => {
  console.error('\nFAILED:', error instanceof Error ? error.message : error);
  process.exit(1);
});

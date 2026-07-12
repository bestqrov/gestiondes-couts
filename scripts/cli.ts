import { extractDocumentText } from '../src/ocr/documentTextExtractor.js';
import { detectAndParsePair } from '../src/parser/detectAndParsePair.js';
import { mergeDeclaration } from '../src/merge/declarationMerger.js';
import { validateArticle } from '../src/domain/validators.js';
import { generateArticleSummaryExcel } from '../src/excel/articleSummaryExcelGenerator.js';
import { generateUnitLevelExcel } from '../src/excel/unitLevelExcelGenerator.js';

async function main() {
  const [, , liquidationPath, dumPath, outDir = '.'] = process.argv;

  if (!liquidationPath || !dumPath) {
    console.error('Usage: npm run generate -- <liquidation-file> <dum-file> [output-dir]');
    process.exit(1);
  }

  console.log(`Reading Liquidation: ${liquidationPath}`);
  const liquidationOcr = await extractDocumentText(liquidationPath);
  console.log(`Reading DUM: ${dumPath}`);
  const dumOcr = await extractDocumentText(dumPath);

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

  const file1Path = `${outDir}/File1-ArticleSummary.xlsx`;
  const file2Path = `${outDir}/File2-UnitLevelDetail.xlsx`;
  await generateArticleSummaryExcel(declaration, file1Path);
  await generateUnitLevelExcel(declaration, file2Path);

  console.log(`\nGenerated:\n  ${file1Path}\n  ${file2Path}`);
}

main().catch((error) => {
  console.error('\nFAILED:', error instanceof Error ? error.message : error);
  process.exit(1);
});

import { describe, it, expect, afterEach } from 'vitest';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import PDFDocument from 'pdfkit';
import { createCanvas } from 'canvas';
import { extractPdfText } from '../../src/ocr/pdfTextExtractor.js';
import { parseLiquidation } from '../../src/parser/liquidation/liquidationParser.js';

async function buildPdf(draw: (doc: PDFKit.PDFDocument) => void): Promise<string> {
  const dir = await mkdtemp(path.join(tmpdir(), 'pdf-extractor-test-'));
  const filePath = path.join(dir, 'test.pdf');
  const doc = new PDFDocument({ size: 'A4' });
  const chunks: Buffer[] = [];
  doc.on('data', (chunk: Buffer) => chunks.push(chunk));
  const donePromise = new Promise<void>((resolve) => doc.on('end', () => resolve()));
  draw(doc);
  doc.end();
  await donePromise;
  await writeFile(filePath, Buffer.concat(chunks));
  return filePath;
}

describe('extractPdfText', () => {
  const tempDirs: string[] = [];

  afterEach(async () => {
    await Promise.all(tempDirs.splice(0).map((dir) => rm(dir, { recursive: true, force: true })));
  });

  it('reads the embedded text layer of a born-digital PDF directly (no OCR)', async () => {
    const filePath = await buildPdf((doc) => {
      doc.fontSize(20).text('DECLARATION 123456');
    });
    tempDirs.push(path.dirname(filePath));

    const result = await extractPdfText(filePath);

    expect(result.text).toContain('DECLARATION 123456');
    expect(result.confidence).toBe(1.0);
  });

  // Regression test for a real bug: a born-digital Liquidation PDF (e.g.
  // exported straight from the BADR customs portal, not scanned) has a real
  // text layer, so it never goes through OCR — but pdfjs's text items must
  // be rejoined into actual lines (using hasEOL) for the Liquidation
  // parser's line-based tax-row parsing to work at all. Without that, a
  // real user's PDF failed with "no tax rows found" even though the text
  // extracted perfectly fine as one long flattened line.
  it('preserves line breaks so a multi-line, box-drawn Liquidation table parses correctly', async () => {
    const filePath = await buildPdf((doc) => {
      doc.fontSize(9);
      const lines = [
        'REDEVABLE : MED AFRICA LOGISTICS CODE : 309536',
        "CATEGORIE D'ORDONNANCEMENT : Credit d'enlevement B E N° : 169 DU : 25/06/2026",
        '',
        'ARTICLE : 1 NUMERO SH : 6109100010 VALEUR : 27 147,00',
        'QUANTITE : 354.000 UNITE : NOMBRE',
        '',
        '! TAXE ! ASSIETTE ! TAUX ! S.TVA ! S.FR ! TAUX VIRTUEL ! MONTANT !',
        '! 000110 ! 27147.00 ! 0.0 ! T ! ! ! 0,00 !',
        '! 007217 ! 27147.00 ! 0.25 ! T ! ! ! 68,00 !',
        '! 002109 ! 27215.00 ! 20.0 ! ! ! ! 5 443,00 !',
        'TOTAL ARTICLE : 5 511,00',
      ];
      for (const line of lines) {
        doc.text(line || ' ');
      }
    });
    tempDirs.push(path.dirname(filePath));

    const result = await extractPdfText(filePath);
    // The core regression check: lines must actually be separated, not
    // flattened into one giant space-joined string.
    expect(result.text.split('\n').length).toBeGreaterThan(5);

    const parsed = parseLiquidation(result.text);
    expect(parsed.header.code).toBe('309536');
    expect(parsed.header.redevable).toBe('MED AFRICA LOGISTICS');
    expect(parsed.articles).toHaveLength(1);
    expect(parsed.articles[0].taxes).toHaveLength(3);
    expect(parsed.articles[0].taxes.map((t) => t.code)).toEqual(['000110', '007217', '002109']);
  });

  it('falls back to rasterize-then-OCR for a scanned PDF with no text layer', async () => {
    const canvas = createCanvas(800, 300);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, 800, 300);
    ctx.fillStyle = '#000000';
    ctx.font = '40px sans-serif';
    ctx.fillText('SCANNE 98765', 40, 150);
    const pageImage = canvas.toBuffer('image/png');

    const filePath = await buildPdf((doc) => {
      doc.image(pageImage, 50, 50, { width: 500 });
    });
    tempDirs.push(path.dirname(filePath));

    const result = await extractPdfText(filePath);

    expect(result.text.trim().length).toBeGreaterThan(0);
    expect(result.confidence).toBeLessThan(1.0);
    expect(result.confidence).toBeGreaterThan(0);
  }, 30000);
});

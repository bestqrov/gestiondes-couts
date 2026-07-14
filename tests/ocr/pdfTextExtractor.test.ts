import { describe, it, expect, afterEach } from 'vitest';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import PDFDocument from 'pdfkit';
import { createCanvas } from 'canvas';
import { extractPdfText } from '../../src/ocr/pdfTextExtractor.js';

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

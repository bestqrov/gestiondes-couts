import sharp from 'sharp';
import { createWorker, PSM } from 'tesseract.js';
import type { OcrResult } from './types.js';

// Liquidation-style documents are printed as a fixed-width ASCII table
// (monospace font, box-drawing borders) rather than natural prose. Upscaling
// and boosting contrast before OCR, plus telling Tesseract to expect a
// single uniform block of text (PSM.SINGLE_BLOCK) rather than guessing at a
// natural-document layout, measurably reduces character-level misreads on
// this kind of tabular content compared to feeding the raw photo/screenshot
// straight to Tesseract's defaults.
async function preprocessForOcr(filePath: string): Promise<Buffer> {
  return sharp(filePath)
    .resize({ width: 2400, withoutEnlargement: false })
    .grayscale()
    .normalize()
    .sharpen()
    .toBuffer();
}

// Scanned pages / photographed documents have no text layer, so we fall back
// to Tesseract (bundled WASM, no system install needed), per design spec §3.1.
export async function extractImageText(filePath: string): Promise<OcrResult> {
  const preprocessed = await preprocessForOcr(filePath);

  const worker = await createWorker('fra');
  try {
    await worker.setParameters({
      tessedit_pageseg_mode: PSM.SINGLE_BLOCK,
      preserve_interword_spaces: '1',
    });
    const { data } = await worker.recognize(preprocessed);
    return { text: data.text, confidence: data.confidence / 100 };
  } finally {
    await worker.terminate();
  }
}

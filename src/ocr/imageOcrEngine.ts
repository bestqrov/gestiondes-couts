import { tmpdir } from 'node:os';
import sharp from 'sharp';
import { createWorker, OEM, PSM } from 'tesseract.js';
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

  // OEM.TESSERACT_ONLY (the "legacy" recognizer) reads this document's
  // monospace/dot-matrix-style font dramatically more accurately than the
  // modern LSTM engine (tesseract.js's default) — empirically confirmed
  // against a real sample: LSTM misread digits throughout the tax table
  // (e.g. "0" as "e"), while the legacy engine transcribed every tax code
  // and amount correctly. legacyCore/legacyLang request the traineddata
  // variant that includes the legacy engine's tables (the default
  // LSTM-only data doesn't have them).
  //
  // Tesseract downloads/caches language data (fra.traineddata) on first
  // use. Point that at the OS temp dir rather than the process's cwd, since
  // cwd isn't guaranteed writable in a container.
  const worker = await createWorker(
    'fra',
    OEM.TESSERACT_ONLY,
    { legacyCore: true, legacyLang: true, cachePath: tmpdir() }
  );
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
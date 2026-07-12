import { createWorker } from 'tesseract.js';
import type { OcrResult } from './types.js';

// Scanned pages / photographed documents have no text layer, so we fall back
// to Tesseract (bundled WASM, no system install needed), per design spec §3.1.
export async function extractImageText(filePath: string): Promise<OcrResult> {
  const worker = await createWorker('fra');
  try {
    const { data } = await worker.recognize(filePath);
    return { text: data.text, confidence: data.confidence / 100 };
  } finally {
    await worker.terminate();
  }
}

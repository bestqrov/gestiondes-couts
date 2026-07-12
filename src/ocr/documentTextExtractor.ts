import path from 'node:path';
import { extractPdfText } from './pdfTextExtractor.js';
import { extractImageText } from './imageOcrEngine.js';
import type { OcrResult } from './types.js';

const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp']);

// Per design spec §3.1: try the PDF text layer first (fast, ~exact); if the
// PDF has no embedded text (a scanned PDF) or the file is an image, fall
// back to Tesseract OCR. Rasterizing a scanned PDF page-by-page before OCR
// is not yet implemented — today, only born-digital PDFs and plain image
// files are supported; a scanned PDF will currently return empty text.
export async function extractDocumentText(filePath: string): Promise<OcrResult> {
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.pdf') {
    return extractPdfText(filePath);
  }

  if (IMAGE_EXTENSIONS.has(ext)) {
    return extractImageText(filePath);
  }

  throw new Error(`Unsupported file type "${ext}" for ${filePath}`);
}

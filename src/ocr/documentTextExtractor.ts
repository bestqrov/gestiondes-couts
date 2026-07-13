import path from 'node:path';
import { extractPdfText } from './pdfTextExtractor.js';
import { extractImageText } from './imageOcrEngine.js';
import { extractImageTextViaGoogleVision } from './googleVisionOcrEngine.js';
import { extractImageTextViaOpenAi } from './openAiOcrEngine.js';
import type { OcrResult } from './types.js';

const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp']);

// Per design spec §3.1: try the PDF text layer first (fast, ~exact); if the
// PDF has no embedded text (a scanned PDF) or the file is an image, fall
// back to OCR. Rasterizing a scanned PDF page-by-page before OCR is not yet
// implemented — today, only born-digital PDFs and plain image files are
// supported; a scanned PDF will currently return empty text.
//
// For images, Tesseract is the default (free, fully local, no API key), but
// it has real accuracy limits on dense/tabular printed text (see the
// Liquidation document's tax table). Two opt-in, paid alternatives are
// available, in priority order:
//   1. OPENAI_API_KEY — GPT-4o vision, best accuracy on this kind of
//      fixed-width table, no dictionary-correction artifacts.
//   2. GOOGLE_VISION_API_KEY — Cloud Vision's DOCUMENT_TEXT_DETECTION.
// Both require a network call and a funded account; Tesseract remains the
// default precisely because it's free and fully local.
export async function extractDocumentText(filePath: string): Promise<OcrResult> {
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.pdf') {
    return extractPdfText(filePath);
  }

  if (IMAGE_EXTENSIONS.has(ext)) {
    if (process.env.OPENAI_API_KEY) {
      return extractImageTextViaOpenAi(filePath);
    }
    if (process.env.GOOGLE_VISION_API_KEY) {
      return extractImageTextViaGoogleVision(filePath);
    }
    return extractImageText(filePath);
  }

  throw new Error(`Unsupported file type "${ext}" for ${filePath}`);
}

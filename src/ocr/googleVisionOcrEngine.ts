import { readFile } from 'node:fs/promises';
import type { OcrResult } from './types.js';

interface GoogleVisionAnnotateResponse {
  responses: Array<{
    fullTextAnnotation?: { text: string };
    textAnnotations?: Array<{ description: string }>;
    error?: { message: string };
  }>;
}

// Google Cloud Vision's DOCUMENT_TEXT_DETECTION is a modern deep-learning OCR
// model, meaningfully more accurate than Tesseract on dense/tabular printed
// text (e.g. the Liquidation document's fixed-width ASCII table). Used as an
// opt-in alternative to Tesseract when GOOGLE_VISION_API_KEY is configured —
// see documentTextExtractor.ts for the selection logic. Confidence is not
// returned per-document by this endpoint, so we report a fixed high
// confidence (0.95) rather than a real measured value.
export async function extractImageTextViaGoogleVision(filePath: string): Promise<OcrResult> {
  const apiKey = process.env.GOOGLE_VISION_API_KEY;
  if (!apiKey) {
    throw new Error('GOOGLE_VISION_API_KEY is not set');
  }

  const imageBytes = await readFile(filePath);
  const base64Image = imageBytes.toString('base64');

  const response = await fetch(
    `https://vision.googleapis.com/v1/images:annotate?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        requests: [
          {
            image: { content: base64Image },
            features: [{ type: 'DOCUMENT_TEXT_DETECTION' }],
          },
        ],
      }),
    }
  );

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Google Vision API request failed (${response.status}): ${body}`);
  }

  const data = (await response.json()) as GoogleVisionAnnotateResponse;
  const result = data.responses[0];

  if (result?.error) {
    throw new Error(`Google Vision API error: ${result.error.message}`);
  }

  const text = result?.fullTextAnnotation?.text ?? '';
  return { text, confidence: 0.95 };
}

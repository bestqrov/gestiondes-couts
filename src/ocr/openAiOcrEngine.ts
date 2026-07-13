import { readFile } from 'node:fs/promises';
import path from 'node:path';
import type { OcrResult } from './types.js';

interface OpenAiChatResponse {
  choices: Array<{ message: { content: string } }>;
  error?: { message: string };
}

const MIME_TYPES: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.tif': 'image/tiff',
  '.tiff': 'image/tiff',
  '.bmp': 'image/bmp',
};

const TRANSCRIPTION_PROMPT = `You are transcribing a scanned Moroccan customs document (either a "Liquidation Douanière" tax table or a "DUM" declaration form). Output the raw text exactly as printed, preserving line breaks and the original layout as closely as possible — this text will be parsed by regex, so precision matters more than readability.

Rules:
- Transcribe every digit exactly as printed. Do not "correct", reformat, or guess at unclear digits — if genuinely illegible, mark it as [?] rather than substituting a plausible-looking character.
- Preserve "!" or "|" table border characters if present in a fixed-width ASCII table, on the same lines they appear.
- Do not translate, summarize, or omit any part of the document.
- Output only the transcribed text, no commentary before or after.`;

// GPT-4o's vision capability reads dense/tabular printed text (e.g. the
// Liquidation document's fixed-width ASCII table) far more reliably than
// Tesseract, and doesn't have Tesseract's dictionary-driven digit-to-letter
// substitution problem. Opt-in via OPENAI_API_KEY — see
// documentTextExtractor.ts for the engine-selection priority.
export async function extractImageTextViaOpenAi(filePath: string): Promise<OcrResult> {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY is not set');
  }

  const ext = path.extname(filePath).toLowerCase();
  const mimeType = MIME_TYPES[ext] ?? 'image/jpeg';
  const imageBytes = await readFile(filePath);
  const dataUrl = `data:${mimeType};base64,${imageBytes.toString('base64')}`;

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: 'gpt-4o',
      temperature: 0,
      messages: [
        {
          role: 'user',
          content: [
            { type: 'text', text: TRANSCRIPTION_PROMPT },
            { type: 'image_url', image_url: { url: dataUrl, detail: 'high' } },
          ],
        },
      ],
    }),
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`OpenAI API request failed (${response.status}): ${body}`);
  }

  const data = (await response.json()) as OpenAiChatResponse;

  if (data.error) {
    throw new Error(`OpenAI API error: ${data.error.message}`);
  }

  const text = data.choices?.[0]?.message?.content ?? '';
  return { text, confidence: 0.95 };
}
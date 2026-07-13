import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { extractImageTextViaOpenAi } from '../../src/ocr/openAiOcrEngine.js';

describe('extractImageTextViaOpenAi', () => {
  const originalKey = process.env.OPENAI_API_KEY;

  beforeEach(() => {
    delete process.env.OPENAI_API_KEY;
  });

  afterEach(() => {
    if (originalKey === undefined) {
      delete process.env.OPENAI_API_KEY;
    } else {
      process.env.OPENAI_API_KEY = originalKey;
    }
  });

  it('throws immediately (no network call) when OPENAI_API_KEY is not set', async () => {
    await expect(extractImageTextViaOpenAi('/some/path.jpg')).rejects.toThrow(
      'OPENAI_API_KEY is not set'
    );
  });
});
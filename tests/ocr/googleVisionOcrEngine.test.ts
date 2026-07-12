import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { extractImageTextViaGoogleVision } from '../../src/ocr/googleVisionOcrEngine.js';

describe('extractImageTextViaGoogleVision', () => {
  const originalKey = process.env.GOOGLE_VISION_API_KEY;

  beforeEach(() => {
    delete process.env.GOOGLE_VISION_API_KEY;
  });

  afterEach(() => {
    if (originalKey === undefined) {
      delete process.env.GOOGLE_VISION_API_KEY;
    } else {
      process.env.GOOGLE_VISION_API_KEY = originalKey;
    }
  });

  it('throws immediately (no network call) when GOOGLE_VISION_API_KEY is not set', async () => {
    await expect(extractImageTextViaGoogleVision('/some/path.jpg')).rejects.toThrow(
      'GOOGLE_VISION_API_KEY is not set'
    );
  });
});

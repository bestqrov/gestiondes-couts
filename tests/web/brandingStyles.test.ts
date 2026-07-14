import { describe, it, expect } from 'vitest';
import { renderBrandOverrideStyle, renderLogoImg } from '../../src/web/brandingStyles.js';
import type { AppSettings } from '../../src/db/appSettingsRepository.js';

const EMPTY_SETTINGS: AppSettings = {
  companyName: null,
  logoDataUri: null,
  brandColor: null,
  fontFamily: null,
};

describe('renderBrandOverrideStyle', () => {
  it('returns an empty string when no brand color or font is configured', () => {
    expect(renderBrandOverrideStyle(EMPTY_SETTINGS)).toBe('');
  });

  it('overrides --brand-600/700/soft when a valid brand color is set', () => {
    const style = renderBrandOverrideStyle({ ...EMPTY_SETTINGS, brandColor: '#4f46e5' });
    expect(style).toContain('--brand-600: #4f46e5;');
    expect(style).toContain('--brand-700:');
    expect(style).toContain('--brand-soft:');
  });

  it('ignores an invalid (malformed) brand color rather than emitting broken CSS', () => {
    const style = renderBrandOverrideStyle({ ...EMPTY_SETTINGS, brandColor: 'not-a-color' });
    expect(style).toBe('');
  });

  it('adds a body font-family rule for a known font choice', () => {
    const style = renderBrandOverrideStyle({ ...EMPTY_SETTINGS, fontFamily: 'serif' });
    expect(style).toContain('font-family: Georgia');
  });

  it('ignores an unrecognized font choice', () => {
    const style = renderBrandOverrideStyle({ ...EMPTY_SETTINGS, fontFamily: 'not-a-real-font' });
    expect(style).toBe('');
  });
});

describe('renderLogoImg', () => {
  it('returns an empty string when no logo is set', () => {
    expect(renderLogoImg(EMPTY_SETTINGS)).toBe('');
  });

  it('renders an img tag with the logo data URI when set', () => {
    const html = renderLogoImg({ ...EMPTY_SETTINGS, logoDataUri: 'data:image/png;base64,abc123' });
    expect(html).toContain('src="data:image/png;base64,abc123"');
    expect(html).toContain('<img');
  });

  it('uses the company name as alt text when available, and escapes it', () => {
    const html = renderLogoImg({
      ...EMPTY_SETTINGS,
      logoDataUri: 'data:image/png;base64,abc123',
      companyName: 'Acme & Sons',
    });
    expect(html).toContain('alt="Acme &amp; Sons"');
  });
});
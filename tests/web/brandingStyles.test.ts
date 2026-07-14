import { describe, it, expect } from 'vitest';
import {
  renderBrandOverrideStyle,
  renderLogoImg,
  renderLoginBadge,
  renderLoginTitle,
  renderContactRows,
} from '../../src/web/brandingStyles.js';
import type { AppSettings } from '../../src/db/appSettingsRepository.js';

const EMPTY_SETTINGS: AppSettings = {
  companyName: null,
  logoDataUri: null,
  brandColor: null,
  fontFamily: null,
  contactEmail: null,
  contactWhatsapp: null,
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

  it('also overrides --header-bg (to the brand color) and --header-ink (to white) for the superadmin topbar', () => {
    const style = renderBrandOverrideStyle({ ...EMPTY_SETTINGS, brandColor: '#4f46e5' });
    expect(style).toContain('--header-bg: #4f46e5;');
    expect(style).toContain('--header-ink: #ffffff;');
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

describe('renderLoginBadge', () => {
  it('falls back to the default document-icon SVG when no logo is set', () => {
    const html = renderLoginBadge(EMPTY_SETTINGS);
    expect(html).toContain('<svg');
    expect(html).not.toContain('<img');
  });

  it('renders the uploaded logo as an img when set', () => {
    const html = renderLoginBadge({ ...EMPTY_SETTINGS, logoDataUri: 'data:image/png;base64,abc123' });
    expect(html).toContain('<img');
    expect(html).toContain('src="data:image/png;base64,abc123"');
  });
});

describe('renderLoginTitle', () => {
  it('falls back to a generic branded title when no company name is set', () => {
    expect(renderLoginTitle(EMPTY_SETTINGS)).toBe('Gestion des <span class="accent">Coûts</span>');
  });

  it('uses the configured company name, escaped, when set', () => {
    expect(renderLoginTitle({ ...EMPTY_SETTINGS, companyName: 'Acme & Sons' })).toBe('Acme &amp; Sons');
  });
});

describe('renderContactRows', () => {
  it('returns an empty string when neither contact field is set', () => {
    expect(renderContactRows(EMPTY_SETTINGS)).toBe('');
  });

  it('renders only the email row when only contactEmail is set', () => {
    const html = renderContactRows({ ...EMPTY_SETTINGS, contactEmail: 'contact@acme.example' });
    expect(html).toContain('mailto:contact@acme.example');
    expect(html).not.toContain('wa.me');
  });

  it('renders only the WhatsApp row when only contactWhatsapp is set, as a digits-only wa.me link', () => {
    const html = renderContactRows({ ...EMPTY_SETTINGS, contactWhatsapp: '+212 6-00 00 00 00' });
    expect(html).toContain('https://wa.me/212600000000');
    expect(html).not.toContain('mailto:');
  });

  it('renders both rows when both are set', () => {
    const html = renderContactRows({
      ...EMPTY_SETTINGS,
      contactEmail: 'contact@acme.example',
      contactWhatsapp: '+212600000000',
    });
    expect(html).toContain('mailto:contact@acme.example');
    expect(html).toContain('https://wa.me/212600000000');
  });

  it('escapes the displayed email/phone text', () => {
    const html = renderContactRows({ ...EMPTY_SETTINGS, contactEmail: 'a<b>@acme.example' });
    expect(html).toContain('a&lt;b&gt;@acme.example');
  });
});
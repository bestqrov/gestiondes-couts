import type { AppSettings } from '../db/appSettingsRepository.js';
import { isValidHexColor, darken, lighten } from '../domain/colorUtils.js';

const FONT_STACKS: Record<string, string> = {
  system: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
  serif: "Georgia, 'Times New Roman', serif",
  mono: "'Courier New', Courier, monospace",
  rounded: "Verdana, Geneva, sans-serif",
};

export const FONT_OPTIONS: Array<{ value: string; label: string }> = [
  { value: 'system', label: 'Système (par défaut)' },
  { value: 'serif', label: 'Serif (Georgia)' },
  { value: 'mono', label: 'Monospace (Courier)' },
  { value: 'rounded', label: 'Rounded (Verdana)' },
];

function fontStack(fontFamily: string | null): string | undefined {
  if (!fontFamily) return undefined;
  return FONT_STACKS[fontFamily];
}

// Produces a <style> block that overrides a page's --brand-* CSS variables
// (and body font-family) with the superadmin-configured company branding,
// falling back to whatever the page already defines when no custom color/
// font has been set. Injected as a second <style> block right after the
// page's own — later declarations for the same custom property win at
// equal specificity, so this doesn't need !important.
export function renderBrandOverrideStyle(settings: AppSettings): string {
  const declarations: string[] = [];

  if (settings.brandColor && isValidHexColor(settings.brandColor)) {
    const brand600 = settings.brandColor;
    const brand700 = darken(brand600, 0.18);
    const brandSoft = lighten(brand600, 0.88);
    declarations.push(`--brand-600: ${brand600};`, `--brand-700: ${brand700};`, `--brand-soft: ${brandSoft};`);
    declarations.push(`--brand: ${brand600};`);
  }

  const stack = fontStack(settings.fontFamily);
  const fontRule = stack ? `body { font-family: ${stack}; }` : '';

  if (declarations.length === 0 && !fontRule) return '';

  return `<style>
    :root { ${declarations.join(' ')} }
    ${fontRule}
  </style>`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// A small logo shown next to the "Déconnexion" button on pages that have
// one — empty string (renders nothing) when no logo has been uploaded.
export function renderLogoImg(settings: AppSettings): string {
  if (!settings.logoDataUri) return '';
  const alt = settings.companyName ? escapeHtml(settings.companyName) : 'Logo';
  return `<img src="${settings.logoDataUri}" alt="${alt}" style="height:32px;max-width:140px;object-fit:contain;border-radius:6px;" />`;
}
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
    // Consumed by the superadmin sidebar shell's .topbar (background) and
    // its title/company-name text — a solid brand-colored header instead
    // of the plain card background, with white text for contrast. Falls
    // back to the shell's own --card-bg/--ink-900 via var(--header-bg,
    // var(--card-bg)) when unset, so pages without these rules are
    // unaffected.
    declarations.push(`--header-bg: ${brand600};`, `--header-ink: #ffffff;`);
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

const DEFAULT_BADGE_ICON =
  '<svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M7 3.5h7l5 5V19a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 5 19V5A1.5 1.5 0 0 1 6.5 3.5H7Z" stroke="white" stroke-width="1.6" stroke-linejoin="round"/><path d="M14 3.5V8a1 1 0 0 0 1 1h4.5" stroke="white" stroke-width="1.6" stroke-linejoin="round"/><path d="M8.5 13.2l2.1 2.1 4.3-4.5" stroke="white" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>';

// The logo badge shown on the login page (both the left brand panel and
// the form panel's small icon) — the uploaded logo when set, otherwise a
// generic document/checkmark icon so the badge is never empty.
export function renderLoginBadge(settings: AppSettings): string {
  if (!settings.logoDataUri) return DEFAULT_BADGE_ICON;
  const alt = settings.companyName ? escapeHtml(settings.companyName) : 'Logo';
  return `<img src="${settings.logoDataUri}" alt="${alt}" />`;
}

// The large title on the login page's left panel — the configured company
// name when set, otherwise a generic two-tone fallback matching the app's
// own name (so the page never looks unbranded/blank before Réglages has
// been filled in).
export function renderLoginTitle(settings: AppSettings): string {
  if (settings.companyName) return escapeHtml(settings.companyName);
  return 'Gestion des <span class="accent">Coûts</span>';
}

const MAIL_ICON =
  '<svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M3 5.5h14v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-9Z" stroke="currentColor" stroke-width="1.4"/><path d="M3.5 5.8l6.5 5 6.5-5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const PHONE_ICON =
  '<svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M5 3.5h2.2l1 4-1.6 1.3a9 9 0 0 0 4.6 4.6l1.3-1.6 4 1v2.2c0 .8-.7 1.4-1.4 1.3A13.5 13.5 0 0 1 3.7 4.9C3.6 4.2 4.2 3.5 5 3.5Z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/></svg>';

// A digits-only WhatsApp deep link (wa.me requires country code + digits,
// no "+", spaces, or punctuation) — built from whatever format the
// superadmin typed the number in, rather than requiring a specific format.
function toWhatsAppLink(rawNumber: string): string {
  return `https://wa.me/${rawNumber.replace(/[^0-9]/g, '')}`;
}

// Contact rows shown under the login form (email / WhatsApp) — only the
// ones actually configured in Réglages are rendered; returns '' entirely
// when neither is set, so no empty divider shows up.
export function renderContactRows(settings: AppSettings): string {
  const rows: string[] = [];
  if (settings.contactEmail) {
    const email = escapeHtml(settings.contactEmail);
    rows.push(
      `<div class="contact-row">${MAIL_ICON}<span>Email :</span> <a href="mailto:${email}">${email}</a></div>`
    );
  }
  if (settings.contactWhatsapp) {
    const phone = escapeHtml(settings.contactWhatsapp);
    rows.push(
      `<div class="contact-row">${PHONE_ICON}<span>WhatsApp :</span> <a href="${toWhatsAppLink(settings.contactWhatsapp)}" target="_blank" rel="noopener noreferrer">${phone}</a></div>`
    );
  }
  if (rows.length === 0) return '';
  return `<p class="contact-hint">Besoin d'aide ?</p>${rows.join('')}`;
}
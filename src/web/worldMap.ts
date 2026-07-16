import type { CountryProductCount } from '../db/transactionsRepository.js';

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// A simple ranked list of countries of origin — no map/pins, just the
// name and total unit quantity for every country seen across saved
// declarations, sorted by quantity descending. Deliberately the physical
// unit quantity (totalQuantite), not the article-line count (productCount)
// — one declaration line for 354 T-shirts is 1 "product line" but 354
// actual units, and showing "1" there reads as obviously wrong.
export function renderWorldMapPanel(countryCounts: CountryProductCount[]): string {
  if (countryCounts.length === 0) {
    return `
      <div class="card placeholder-card">
        <div class="placeholder-icon"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 2.5c-3 3-3 12 0 15M10 2.5c3 3 3 12 0 15M2.5 10h15" stroke="currentColor" stroke-width="1.4"/><circle cx="10" cy="10" r="7.5" stroke="currentColor" stroke-width="1.4"/></svg></div>
        <h2>Pas encore de données</h2>
        <p>La répartition géographique des produits s'affichera ici après la génération de déclarations.</p>
      </div>
    `;
  }

  const rows = countryCounts
    .map((c) => {
      const quantity = Math.round(c.totalQuantite);
      return `<div class="map-legend-row">
        <span class="map-legend-name">${escapeHtml(c.pays)}</span>
        <span class="map-legend-count">${quantity} unité${quantity > 1 ? 's' : ''}</span>
      </div>`;
    })
    .join('');

  return `
    <div class="card">
      <h2>Répartition géographique des produits</h2>
      <p class="lede" style="margin-top:-8px;">Pays d'origine des produits déclarés, toutes déclarations confondues.</p>
      <div class="map-legend">${rows}</div>
    </div>
  `;
}

export const WORLD_MAP_STYLE = `
  .map-legend { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 8px 16px; }
  .map-legend-row {
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
    font-size: 13px; padding: 9px 12px; background: var(--input-bg); border-radius: 8px;
  }
  .map-legend-name { color: var(--ink-700); font-weight: 600; }
  .map-legend-count { color: var(--brand-700); font-weight: 700; font-variant-numeric: tabular-nums; }
`;

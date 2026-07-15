import type { CountryProductCount } from '../db/transactionsRepository.js';

// Centroids [longitude, latitude] for countries as written (in French,
// uppercase) on Moroccan customs documents — covers the trading partners
// this app realistically sees. Deliberately a curated lookup rather than a
// full geo dataset/library: this app has no bundled map assets, and the
// country set on any given declaration is small and predictable. A country
// not in this table still shows up in the text legend, just without a pin.
const COUNTRY_COORDINATES: Record<string, [number, number]> = {
  MAROC: [-6.8, 31.8],
  ESPAGNE: [-3.7, 40.2],
  FRANCE: [2.5, 46.6],
  ITALIE: [12.6, 42.8],
  ALLEMAGNE: [10.4, 51.2],
  PORTUGAL: [-8.2, 39.6],
  BELGIQUE: [4.5, 50.6],
  'PAYS-BAS': [5.3, 52.2],
  'PAYS BAS': [5.3, 52.2],
  'ROYAUME-UNI': [-2.0, 54.0],
  ANGLETERRE: [-2.0, 54.0],
  SUISSE: [8.2, 46.8],
  AUTRICHE: [14.3, 47.6],
  POLOGNE: [19.4, 52.0],
  ROUMANIE: [24.9, 45.9],
  BULGARIE: [25.5, 42.7],
  GRECE: [21.8, 39.1],
  'REPUBLIQUE TCHEQUE': [15.5, 49.8],
  HONGRIE: [19.5, 47.2],
  SUEDE: [16.5, 62.2],
  DANEMARK: [9.5, 56.0],
  FINLANDE: [26.0, 64.0],
  IRLANDE: [-8.0, 53.4],
  TURQUIE: [35.0, 39.0],
  TUNISIE: [9.5, 34.0],
  ALGERIE: [2.6, 28.0],
  EGYPTE: [30.0, 26.5],
  CHINE: [104.2, 35.9],
  'HONG KONG': [114.2, 22.3],
  INDE: [79.0, 22.0],
  BANGLADESH: [90.4, 23.7],
  PAKISTAN: [69.3, 30.4],
  VIETNAM: [108.3, 14.1],
  INDONESIE: [113.9, -0.8],
  THAILANDE: [101.0, 15.9],
  MALAISIE: [101.9, 4.2],
  CAMBODGE: [104.9, 12.6],
  JAPON: [138.3, 36.2],
  'COREE DU SUD': [127.8, 36.3],
  'ETATS-UNIS': [-98.5, 39.8],
  USA: [-98.5, 39.8],
  CANADA: [-106.3, 56.1],
  BRESIL: [-51.9, -14.2],
  MEXIQUE: [-102.6, 23.6],
  'ARABIE SAOUDITE': [45.1, 23.9],
  'EMIRATS ARABES UNIS': [54.0, 24.0],
  SENEGAL: [-14.5, 14.5],
  'COTE D\'IVOIRE': [-5.5, 7.5],
  NIGERIA: [8.7, 9.1],
  'AFRIQUE DU SUD': [22.9, -30.6],
};

function projectToPercent([lon, lat]: [number, number]): { x: number; y: number } {
  return { x: ((lon + 180) / 360) * 100, y: ((90 - lat) / 180) * 100 };
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// A magnitude ("how many products came from where") map, not a precise
// atlas — pins are placed at real country centroids (correct relative
// geography) on a plain grid backdrop rather than traced coastlines, since
// this app has no bundled map/geo asset to draw borders from. Marker size
// (not color/hue) encodes magnitude, per a single-hue sequential scale —
// the brand color throughout, at a fixed opacity, sized by product count.
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

  const maxCount = Math.max(...countryCounts.map((c) => c.productCount));
  const plotted: string[] = [];
  const unplottedNames: string[] = [];

  for (const country of countryCounts) {
    const coords = COUNTRY_COORDINATES[country.pays.trim().toUpperCase()];
    if (!coords) {
      unplottedNames.push(country.pays);
      continue;
    }
    const { x, y } = projectToPercent(coords);
    // Square-root scale so area (not radius) is roughly proportional to
    // count — avoids a single large outlier visually swallowing the rest.
    const radius = 5 + Math.sqrt(country.productCount / maxCount) * 13;
    plotted.push(`
      <g class="map-pin">
        <circle cx="${x}%" cy="${y}%" r="${radius}" />
        <title>${escapeHtml(country.pays)} — ${country.productCount} produit${country.productCount > 1 ? 's' : ''}</title>
      </g>
    `);
  }

  const legendRows = countryCounts
    .slice(0, 8)
    .map(
      (c) => `<div class="map-legend-row">
        <span class="map-legend-name">${escapeHtml(c.pays)}</span>
        <span class="map-legend-count">${c.productCount}</span>
      </div>`
    )
    .join('');

  const unplottedNote = unplottedNames.length
    ? `<p class="map-note">Sans position sur la carte : ${unplottedNames.map(escapeHtml).join(', ')}.</p>`
    : '';

  const gridLines = [10, 20, 30, 40, 50, 60, 70, 80, 90]
    .map((pct) => `<line x1="0" y1="${pct}%" x2="100%" y2="${pct}%" class="map-grid-line" />`)
    .join('');
  const gridColumns = [10, 20, 30, 40, 50, 60, 70, 80, 90]
    .map((pct) => `<line x1="${pct}%" y1="0" x2="${pct}%" y2="100%" class="map-grid-line" />`)
    .join('');

  return `
    <div class="card">
      <h2>Répartition géographique des produits</h2>
      <p class="lede" style="margin-top:-8px;">Pays d'origine des produits déclarés, toutes déclarations confondues. La taille du point indique le nombre de produits.</p>
      <div class="map-wrap">
        <svg viewBox="0 0 1000 500" preserveAspectRatio="xMidYMid meet" class="world-map">
          <rect x="0" y="0" width="100%" height="100%" class="map-ocean" />
          ${gridLines}
          ${gridColumns}
          <line x1="0" y1="50%" x2="100%" y2="50%" class="map-equator" />
          ${plotted.join('')}
        </svg>
      </div>
      <div class="map-legend">${legendRows}</div>
      ${unplottedNote}
    </div>
  `;
}

export const WORLD_MAP_STYLE = `
  .map-wrap { border-radius: 12px; overflow: hidden; border: 1px solid var(--line); margin-bottom: 16px; }
  .world-map { display: block; width: 100%; height: auto; aspect-ratio: 2 / 1; }
  .map-ocean { fill: var(--input-bg); }
  .map-grid-line { stroke: var(--line); stroke-width: 1; opacity: 0.5; }
  .map-equator { stroke: var(--line); stroke-width: 1.5; opacity: 0.8; stroke-dasharray: 4 4; }
  .map-pin circle {
    fill: var(--brand-600); fill-opacity: 0.85; stroke: var(--card-bg); stroke-width: 2;
    transition: fill-opacity 0.15s;
  }
  .map-pin:hover circle { fill-opacity: 1; }
  .map-legend { display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 8px 16px; }
  .map-legend-row {
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
    font-size: 12.5px; padding: 6px 10px; background: var(--input-bg); border-radius: 7px;
  }
  .map-legend-name { color: var(--ink-700); font-weight: 600; }
  .map-legend-count { color: var(--brand-700); font-weight: 700; font-variant-numeric: tabular-nums; }
  .map-note { font-size: 11.5px; color: var(--ink-400); margin: 10px 0 0; }
`;

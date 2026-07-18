import type { CountryProductCount } from '../db/transactionsRepository.js';

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const MONTH_NAMES_FR = [
  'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
  'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre',
];

// "2026-07" -> "Juillet 2026" — falls back to the raw period string if it
// isn't in the expected shape, rather than throwing over a display label.
function formatPeriodLabel(period: string): string {
  const match = /^(\d{4})-(\d{2})$/.exec(period);
  if (!match) return period;
  const [, year, month] = match;
  const monthName = MONTH_NAMES_FR[Number.parseInt(month, 10) - 1];
  return monthName ? `${monthName} ${year}` : period;
}

function renderPeriodFilter(availablePeriods: string[], selectedPeriod: string): string {
  if (availablePeriods.length === 0) return '';
  const options = [
    `<option value=""${selectedPeriod ? '' : ' selected'}>Toutes les périodes</option>`,
    ...availablePeriods.map(
      (period) =>
        `<option value="${period}"${period === selectedPeriod ? ' selected' : ''}>${formatPeriodLabel(period)}</option>`
    ),
  ].join('');
  return `
    <form method="get" action="/superadmin/dashboard" class="map-period-filter">
      <label for="period">Période</label>
      <select id="period" name="period" onchange="this.form.submit()">${options}</select>
    </form>
  `;
}

// A simple ranked list of countries of origin — no map/pins, just the
// name and total unit quantity for every country seen across saved
// declarations, sorted by quantity descending. Deliberately the physical
// unit quantity (totalQuantite), not the article-line count (productCount)
// — one declaration line for 354 T-shirts is 1 "product line" but 354
// actual units, and showing "1" there reads as obviously wrong.
//
// `availablePeriods`/`selectedPeriod` back an optional month filter — only
// months that actually have at least one saved declaration are offered, so
// the dropdown never leads to a guaranteed-empty selection.
export function renderWorldMapPanel(
  countryCounts: CountryProductCount[],
  availablePeriods: string[] = [],
  selectedPeriod = ''
): string {
  const periodFilter = renderPeriodFilter(availablePeriods, selectedPeriod);

  if (countryCounts.length === 0) {
    const message = selectedPeriod
      ? `Aucune déclaration générée pour ${formatPeriodLabel(selectedPeriod)}.`
      : "La répartition géographique des produits s'affichera ici après la génération de déclarations.";
    return `
      ${periodFilter}
      <div class="card placeholder-card">
        <div class="placeholder-icon"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 2.5c-3 3-3 12 0 15M10 2.5c3 3 3 12 0 15M2.5 10h15" stroke="currentColor" stroke-width="1.4"/><circle cx="10" cy="10" r="7.5" stroke="currentColor" stroke-width="1.4"/></svg></div>
        <h2>Pas encore de données</h2>
        <p>${escapeHtml(message)}</p>
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
      <div class="map-header">
        <div>
          <h2>Répartition géographique des produits</h2>
          <p class="lede" style="margin:-4px 0 0;">Pays d'origine des produits déclarés${selectedPeriod ? ` — ${formatPeriodLabel(selectedPeriod)}` : ', toutes déclarations confondues'}.</p>
        </div>
        ${periodFilter}
      </div>
      <div class="map-legend">${rows}</div>
    </div>
  `;
}

export const WORLD_MAP_STYLE = `
  .map-header { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; margin-bottom: 18px; flex-wrap: wrap; }
  .map-period-filter { display: flex; align-items: center; gap: 8px; }
  .map-period-filter label { margin: 0; font-size: 12.5px; font-weight: 600; color: var(--ink-500); white-space: nowrap; }
  .map-period-filter select {
    padding: 8px 12px; font-size: 13px; font-family: inherit; color: var(--ink-900);
    background: var(--input-bg); border: 1.5px solid var(--line); border-radius: 8px; cursor: pointer;
  }
  .map-legend { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 8px 16px; }
  .map-legend-row {
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
    font-size: 13px; padding: 9px 12px; background: var(--input-bg); border-radius: 8px;
  }
  .map-legend-name { color: var(--ink-700); font-weight: 600; }
  .map-legend-count { color: var(--brand-700); font-weight: 700; font-variant-numeric: tabular-nums; }
`;

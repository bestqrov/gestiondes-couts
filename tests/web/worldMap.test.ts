import { describe, it, expect } from 'vitest';
import { renderWorldMapPanel } from '../../src/web/worldMap.js';

describe('renderWorldMapPanel', () => {
  it('shows the empty-state message with no period filter when there are no countries and no periods', () => {
    const html = renderWorldMapPanel([]);
    expect(html).toContain('Pas encore de données');
    expect(html).not.toContain('<select');
  });

  it('shows a period-specific empty-state message when a period is selected but has no data', () => {
    const html = renderWorldMapPanel([], ['2026-06', '2026-07'], '2026-05');
    expect(html).toContain('Aucune déclaration générée pour Mai 2026');
  });

  it('renders the period dropdown with "Toutes les périodes" plus each available period, formatted', () => {
    const html = renderWorldMapPanel(
      [{ pays: 'ITALIE', productCount: 1, totalQuantite: 354 }],
      ['2026-07', '2026-06'],
      ''
    );
    expect(html).toContain('<option value="" selected>Toutes les périodes</option>');
    expect(html).toContain('<option value="2026-07">Juillet 2026</option>');
    expect(html).toContain('<option value="2026-06">Juin 2026</option>');
  });

  it('marks the selected period as selected in the dropdown', () => {
    const html = renderWorldMapPanel(
      [{ pays: 'ITALIE', productCount: 1, totalQuantite: 354 }],
      ['2026-07', '2026-06'],
      '2026-06'
    );
    expect(html).toContain('<option value="2026-06" selected>Juin 2026</option>');
    expect(html).toContain('<option value="">Toutes les périodes</option>');
  });

  it('renders no period filter form when no periods are available', () => {
    const html = renderWorldMapPanel([{ pays: 'ITALIE', productCount: 1, totalQuantite: 354 }]);
    expect(html).not.toContain('<select');
  });

  it('renders country rows with the total quantity, unaffected by the period filter UI', () => {
    const html = renderWorldMapPanel(
      [{ pays: 'ITALIE', productCount: 1, totalQuantite: 354 }],
      ['2026-07'],
      '2026-07'
    );
    expect(html).toContain('ITALIE');
    expect(html).toContain('354 unités');
  });
});

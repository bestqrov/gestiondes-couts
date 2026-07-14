import type { Declaration } from '../domain/types.js';
import { allocateTaxAcrossUnits, unionTaxCodes } from '../excel/unitLevelTaxHelpers.js';

const UNIT_PREVIEW_ROW_LIMIT = 200;

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function renderArticleSummaryTable(declaration: Declaration): string {
  const rows = declaration.articles
    .map(
      (article) => `<tr>
        <td>${escapeHtml(article.nomArticle)}</td>
        <td>${escapeHtml(article.hsCode)}</td>
        <td>${escapeHtml(article.pays)}</td>
        <td>${article.valeurDeclaree.toFixed(2)}</td>
        <td>${article.quantite}</td>
      </tr>`
    )
    .join('');

  return `<table>
    <thead><tr><th>Nom Article</th><th>HSC</th><th>Pays</th><th>Valeur déclarée</th><th>Unité</th></tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function renderUnitLevelPreviewTable(declaration: Declaration): { html: string; totalRows: number } {
  const taxCodes = unionTaxCodes(declaration.articles);
  const headerCells = ['Nom Article', 'HSC', 'Serial Number', ...taxCodes]
    .map((h) => `<th>${escapeHtml(h)}</th>`)
    .join('');

  let totalRows = 0;
  const previewRows: string[] = [];

  for (const article of declaration.articles) {
    const quantite = Math.round(article.quantite);
    const perCodeAllocations = new Map<string, number[]>();
    for (const code of taxCodes) {
      const tax = article.taxes.find((t) => t.code === code);
      perCodeAllocations.set(
        code,
        tax ? allocateTaxAcrossUnits(tax.montant, quantite) : new Array(quantite).fill(0)
      );
    }

    for (let unit = 0; unit < quantite; unit++) {
      totalRows++;
      if (previewRows.length < UNIT_PREVIEW_ROW_LIMIT) {
        const cells = taxCodes
          .map((code) => `<td>${perCodeAllocations.get(code)![unit].toFixed(2)}</td>`)
          .join('');
        previewRows.push(
          `<tr><td>${escapeHtml(article.nomArticle)}</td><td>${escapeHtml(article.hsCode)}</td><td>${unit + 1}</td>${cells}</tr>`
        );
      }
    }
  }

  return {
    html: `<table><thead><tr>${headerCells}</tr></thead><tbody>${previewRows.join('')}</tbody></table>`,
    totalRows,
  };
}

// Renders just the results content (heading + the two tables), with no
// <html>/<head>/<body> wrapper and no download link — meant to be fetched
// and injected inline into the upload page's own success panel (see
// GET /last-declaration-results in server.ts), which supplies its own
// export buttons around this fragment. renderResultsPage() below still
// wraps the same tables in a full standalone page for the separate
// GET /results route.
export function renderResultsFragment(declaration: Declaration): string {
  const file1Table = renderArticleSummaryTable(declaration);
  const { html: file2Table, totalRows } = renderUnitLevelPreviewTable(declaration);
  const previewNote =
    totalRows > UNIT_PREVIEW_ROW_LIMIT
      ? `<p class="note">Aperçu limité à ${UNIT_PREVIEW_ROW_LIMIT} lignes sur ${totalRows} — téléchargez le fichier complet pour tout voir.</p>`
      : '';

  return `
    <h2 class="results-heading">Déclaration ${escapeHtml(declaration.code)} — ${escapeHtml(declaration.redevable)}</h2>
    <div class="results-columns">
      <div class="results-column results-column-a">
        <h3><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 4.5h12v11a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1v-11Z" stroke="currentColor" stroke-width="1.4"/><path d="M4 8h12M7.5 8v8.5" stroke="currentColor" stroke-width="1.4"/></svg> Résumé Articles</h3>
        <div class="results-table-scroll">${file1Table}</div>
      </div>
      <div class="results-column results-column-b">
        <h3><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M3 16.5V9M8 16.5V4M13 16.5v-6M17.5 16.5V7" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg> Détail par unité</h3>
        ${previewNote}
        <div class="results-table-scroll">${file2Table}</div>
      </div>
    </div>
  `;
}

export function renderResultsPage(declaration: Declaration): string {
  const file1Table = renderArticleSummaryTable(declaration);
  const { html: file2Table, totalRows } = renderUnitLevelPreviewTable(declaration);
  const previewNote =
    totalRows > UNIT_PREVIEW_ROW_LIMIT
      ? `<p class="note">Aperçu limité à ${UNIT_PREVIEW_ROW_LIMIT} lignes sur ${totalRows} — téléchargez le fichier complet pour tout voir.</p>`
      : '';

  return `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8" />
<title>Résultats</title>
<style>
  body { font-family: -apple-system, sans-serif; margin: 24px; color: #1a1a1a; }
  h1 { font-size: 18px; }
  .columns { display: flex; gap: 24px; flex-wrap: wrap; }
  .column { flex: 1; min-width: 320px; }
  .column h2 { font-size: 15px; }
  table { border-collapse: collapse; width: 100%; font-size: 12px; }
  th, td { border: 1px solid #ddd; padding: 4px 8px; text-align: left; white-space: nowrap; }
  th { background: #f3f4f6; position: sticky; top: 0; }
  .table-scroll { max-height: 500px; overflow: auto; border: 1px solid #ddd; }
  .note { font-size: 12px; color: #666; }
  a.download {
    display: inline-block; margin-top: 12px; padding: 10px 16px; background: #2563eb;
    color: white; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 13px;
  }
  a.back { display: inline-block; margin-bottom: 16px; font-size: 13px; }
</style>
</head>
<body>
  <a class="back" href="/">&larr; Nouvelle génération</a>
  <h1>Résultats — Déclaration ${escapeHtml(declaration.code)} (${escapeHtml(declaration.redevable)})</h1>
  <div class="columns">
    <div class="column">
      <h2>Feuille 1 — Article Summary</h2>
      <div class="table-scroll">${file1Table}</div>
      <a class="download" href="/download">Télécharger Excel (2 feuilles)</a>
    </div>
    <div class="column">
      <h2>Feuille 2 — Unit-Level Cost Detail</h2>
      ${previewNote}
      <div class="table-scroll">${file2Table}</div>
      <a class="download" href="/download">Télécharger Excel (2 feuilles)</a>
    </div>
  </div>
</body>
</html>`;
}

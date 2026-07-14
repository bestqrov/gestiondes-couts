import PDFDocument from 'pdfkit';
import type { Declaration } from '../domain/types.js';
import type { AppSettings } from '../db/appSettingsRepository.js';
import { allocateTaxAcrossUnits, unionTaxCodes } from '../excel/unitLevelTaxHelpers.js';
import { isValidHexColor, darken } from '../domain/colorUtils.js';

const DEFAULT_BRAND_COLOR = '#4F46E5';
const UNIT_PREVIEW_ROW_LIMIT = 200; // same cap as the on-screen results preview

const PAGE_MARGIN = 36;
const HEADER_HEIGHT = 64;
const FOOTER_HEIGHT = 28;
const LOGO_MAX_MIME = new Set(['image/png', 'image/jpeg']); // pdfkit has no SVG/WEBP support

interface Column {
  header: string;
  width: number;
  key: string;
}

function parseLogoBuffer(logoDataUri: string | null): Buffer | undefined {
  if (!logoDataUri) return undefined;
  const match = logoDataUri.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) return undefined;
  const [, mime, base64] = match;
  if (!LOGO_MAX_MIME.has(mime)) return undefined;
  try {
    return Buffer.from(base64, 'base64');
  } catch {
    return undefined;
  }
}

// Draws the letterhead (logo + company name + generation date, top) and
// footer (company name, bottom) on every buffered page — done as a final
// pass over doc.bufferedPageRange() rather than on each addPage(), so the
// table-drawing code below doesn't need to know about the letterhead at
// all; it just draws rows and calls addPage() when it runs out of room.
function drawLetterheadOnAllPages(
  doc: PDFKit.PDFDocument,
  companyName: string | null,
  logoBuffer: Buffer | undefined,
  generatedAt: Date
): void {
  const range = doc.bufferedPageRange();
  const dateLabel = generatedAt.toLocaleDateString('fr-FR', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });

  for (let i = range.start; i < range.start + range.count; i++) {
    doc.switchToPage(i);
    const pageWidth = doc.page.width;
    const pageHeight = doc.page.height;

    // Header
    let logoWidth = 0;
    if (logoBuffer) {
      try {
        doc.image(logoBuffer, PAGE_MARGIN, PAGE_MARGIN - 6, { fit: [50, 36] });
        logoWidth = 60;
      } catch {
        // Corrupt/unsupported image data — skip the logo rather than fail the whole PDF.
      }
    }
    doc
      .font('Helvetica-Bold')
      .fontSize(13)
      .fillColor('#0f172a')
      .text(companyName ?? 'Déclaration Douanière', PAGE_MARGIN + logoWidth, PAGE_MARGIN - 4, {
        width: pageWidth - PAGE_MARGIN * 2 - logoWidth - 140,
      });
    doc
      .font('Helvetica')
      .fontSize(9)
      .fillColor('#64748b')
      .text(`Généré le ${dateLabel}`, pageWidth - PAGE_MARGIN - 140, PAGE_MARGIN - 2, {
        width: 140,
        align: 'right',
      });
    doc
      .moveTo(PAGE_MARGIN, HEADER_HEIGHT)
      .lineTo(pageWidth - PAGE_MARGIN, HEADER_HEIGHT)
      .strokeColor('#e2e8f0')
      .lineWidth(1)
      .stroke();

    // Footer
    doc
      .moveTo(PAGE_MARGIN, pageHeight - FOOTER_HEIGHT)
      .lineTo(pageWidth - PAGE_MARGIN, pageHeight - FOOTER_HEIGHT)
      .strokeColor('#e2e8f0')
      .stroke();
    if (companyName) {
      doc
        .font('Helvetica')
        .fontSize(9)
        .fillColor('#94a3b8')
        .text(companyName, PAGE_MARGIN, pageHeight - FOOTER_HEIGHT + 8, {
          width: pageWidth - PAGE_MARGIN * 2,
          align: 'center',
        });
    }
  }
}

// Draws a colored-header, banded-row table starting at `startY`, paginating
// (adding a page and redrawing the column header) whenever a row would run
// past the reserved footer area. Returns the Y position immediately below
// the finished table, for whatever comes next.
function drawTable(
  doc: PDFKit.PDFDocument,
  columns: Column[],
  rows: string[][],
  startY: number,
  brand600: string
): number {
  const rowHeight = 18;
  const headerHeight = 20;
  const tableX = PAGE_MARGIN;
  const tableWidth = columns.reduce((sum, c) => sum + c.width, 0);
  const maxY = doc.page.height - FOOTER_HEIGHT - 12;

  function drawHeaderRow(y: number): number {
    doc.rect(tableX, y, tableWidth, headerHeight).fill(brand600);
    let x = tableX;
    doc.font('Helvetica-Bold').fontSize(8).fillColor('#ffffff');
    for (const col of columns) {
      doc.text(col.header, x + 4, y + 6, { width: col.width - 8, ellipsis: true });
      x += col.width;
    }
    return y + headerHeight;
  }

  let y = drawHeaderRow(startY);

  rows.forEach((row, index) => {
    if (y + rowHeight > maxY) {
      doc.addPage();
      y = drawHeaderRow(HEADER_HEIGHT + 12);
    }
    if (index % 2 === 1) {
      doc.rect(tableX, y, tableWidth, rowHeight).fill('#f3f4f6');
    }
    let x = tableX;
    doc.font('Helvetica').fontSize(8).fillColor('#0f172a');
    row.forEach((cell, colIndex) => {
      const col = columns[colIndex];
      doc.text(cell, x + 4, y + 5, { width: col.width - 8, ellipsis: true });
      x += col.width;
    });
    doc
      .moveTo(tableX, y + rowHeight)
      .lineTo(tableX + tableWidth, y + rowHeight)
      .strokeColor('#e2e8f0')
      .lineWidth(0.5)
      .stroke();
    y += rowHeight;
  });

  return y;
}

function buildArticleSummaryRows(declaration: Declaration): string[][] {
  return declaration.articles.map((article) => [
    article.nomArticle,
    article.hsCode,
    article.pays,
    article.valeurDeclaree.toFixed(2),
    String(article.quantite),
  ]);
}

function buildUnitLevelTable(declaration: Declaration): { columns: Column[]; rows: string[][] } {
  const taxCodes = unionTaxCodes(declaration.articles);
  const codeColumnWidth = Math.max(40, Math.min(60, 380 / Math.max(taxCodes.length, 1)));
  const columns: Column[] = [
    { header: 'Nom Article', width: 140, key: 'nomArticle' },
    { header: 'HSC', width: 80, key: 'hsCode' },
    { header: 'N°', width: 40, key: 'serial' },
    ...taxCodes.map((code) => ({ header: code, width: codeColumnWidth, key: code })),
  ];

  const rows: string[][] = [];
  outer: for (const article of declaration.articles) {
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
      if (rows.length >= UNIT_PREVIEW_ROW_LIMIT) break outer;
      rows.push([
        article.nomArticle,
        article.hsCode,
        String(unit + 1),
        ...taxCodes.map((code) => perCodeAllocations.get(code)![unit].toFixed(2)),
      ]);
    }
  }

  return { columns, rows };
}

// Generates a landscape A4 PDF replicating the on-screen "Afficher
// résultats" tables (Article Summary + a unit-level preview, same 200-row
// cap as the on-screen version) with the same colored-header/banded-row
// look as the Excel export, plus a letterhead (logo, company name,
// generation date) and footer (company name) repeated on every page.
export function generateDeclarationPdf(
  declaration: Declaration,
  settings: AppSettings
): PDFKit.PDFDocument {
  const brand600 =
    settings.brandColor && isValidHexColor(settings.brandColor)
      ? settings.brandColor
      : DEFAULT_BRAND_COLOR;

  const doc = new PDFDocument({
    layout: 'landscape',
    size: 'A4',
    margins: { top: HEADER_HEIGHT + 12, bottom: FOOTER_HEIGHT + 12, left: PAGE_MARGIN, right: PAGE_MARGIN },
    bufferPages: true,
  });

  let y = HEADER_HEIGHT + 20;
  doc.font('Helvetica-Bold').fontSize(12).fillColor('#0f172a').text(
    `Déclaration ${declaration.code} — ${declaration.redevable}`,
    PAGE_MARGIN,
    y
  );
  y += 22;

  doc.font('Helvetica-Bold').fontSize(10).fillColor('#334155').text('Résumé Articles', PAGE_MARGIN, y);
  y += 16;
  y = drawTable(
    doc,
    [
      { header: 'Nom Article', width: 200, key: 'nomArticle' },
      { header: 'HSC', width: 100, key: 'hsCode' },
      { header: 'Pays', width: 120, key: 'pays' },
      { header: 'Valeur déclarée', width: 120, key: 'valeurDeclaree' },
      { header: 'Unité', width: 80, key: 'quantite' },
    ],
    buildArticleSummaryRows(declaration),
    y,
    brand600
  );
  y += 24;

  if (y > doc.page.height - FOOTER_HEIGHT - 60) {
    doc.addPage();
    y = HEADER_HEIGHT + 20;
  }

  doc
    .font('Helvetica-Bold')
    .fontSize(10)
    .fillColor('#334155')
    .text('Détail par unité (aperçu)', PAGE_MARGIN, y);
  y += 16;
  const { columns, rows } = buildUnitLevelTable(declaration);
  drawTable(doc, columns, rows, y, darken(brand600, 0.1));

  const logoBuffer = parseLogoBuffer(settings.logoDataUri);
  drawLetterheadOnAllPages(doc, settings.companyName, logoBuffer, new Date());

  return doc;
}

import PDFDocument from 'pdfkit';
import type { Declaration } from '../domain/types.js';
import type { AppSettings } from '../db/appSettingsRepository.js';
import { allocateTaxAcrossUnits, unionTaxCodes } from '../excel/unitLevelTaxHelpers.js';
import { isValidHexColor } from '../domain/colorUtils.js';

const DEFAULT_BRAND_COLOR = '#4F46E5';

const PAGE_MARGIN = 36;
const HEADER_HEIGHT = 76;
const FOOTER_HEIGHT = 26;
const LOGO_MAX_MIME = new Set(['image/png', 'image/jpeg']); // pdfkit has no SVG/WEBP support

interface Column {
  header: string;
  width: number;
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

// Draws the letterhead — a solid brand-colored band (company name +
// document title, white text, logo if set) — and a footer (company name)
// on every buffered page. Done as a final pass over doc.bufferedPageRange()
// rather than on each addPage(), so the table-drawing code below doesn't
// need to know about the letterhead at all; it just draws rows and calls
// addPage() when it runs out of room.
function drawLetterheadOnAllPages(
  doc: PDFKit.PDFDocument,
  documentTitle: string,
  companyName: string | null,
  logoBuffer: Buffer | undefined,
  brand600: string,
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

    // Header band — filled with the app/brand color so the letterhead is
    // clearly branded, not just colored table headers further down.
    doc.rect(0, 0, pageWidth, HEADER_HEIGHT).fill(brand600);

    let logoWidth = 0;
    if (logoBuffer) {
      try {
        doc.image(logoBuffer, PAGE_MARGIN, 14, { fit: [46, 46] });
        logoWidth = 58;
      } catch {
        // Corrupt/unsupported image data — skip the logo rather than fail the whole PDF.
      }
    }
    const textX = PAGE_MARGIN + logoWidth;
    const textWidth = pageWidth - PAGE_MARGIN * 2 - logoWidth - 150;
    doc
      .font('Helvetica-Bold')
      .fontSize(15)
      .fillColor('#ffffff')
      .text(companyName ?? 'Déclaration Douanière', textX, 16, { width: textWidth });
    doc
      .font('Helvetica')
      .fontSize(10)
      .fillColor('#e0e7ff')
      .text(documentTitle, textX, 38, { width: textWidth });
    doc
      .font('Helvetica')
      .fontSize(9)
      .fillColor('#ffffff')
      .text(`Généré le ${dateLabel}`, pageWidth - PAGE_MARGIN - 150, 16, {
        width: 150,
        align: 'right',
      });

    // Footer
    doc
      .moveTo(PAGE_MARGIN, pageHeight - FOOTER_HEIGHT)
      .lineTo(pageWidth - PAGE_MARGIN, pageHeight - FOOTER_HEIGHT)
      .strokeColor('#e2e8f0')
      .lineWidth(1)
      .stroke();
    if (companyName) {
      doc
        .font('Helvetica')
        .fontSize(9)
        .fillColor('#94a3b8')
        .text(companyName, PAGE_MARGIN, pageHeight - FOOTER_HEIGHT + 7, {
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
  const rowHeight = 16;
  const headerHeight = 20;
  const tableX = PAGE_MARGIN;
  const tableWidth = columns.reduce((sum, c) => sum + c.width, 0);
  const maxY = doc.page.height - FOOTER_HEIGHT - 10;

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
      doc.text(cell, x + 4, y + 4, { width: col.width - 8, ellipsis: true });
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

// Mirrors the Excel export's "Global" sheet exactly: every article's unit
// rows, one after another, with Nom Article / HSC / Serial Number / [sorted
// tax codes] / Valeur Déclarée (per-unit, same value repeated across an
// article's rows) — no row cap, unlike the on-screen results preview,
// since this is meant to be the same complete data as the Excel sheet.
function buildGlobalSheetTable(declaration: Declaration): { columns: Column[]; rows: string[][] } {
  const taxCodes = unionTaxCodes(declaration.articles);
  const codeColumnWidth = Math.max(38, Math.min(56, 340 / Math.max(taxCodes.length, 1)));
  const columns: Column[] = [
    { header: 'Nom Article', width: 130 },
    { header: 'HSC', width: 75 },
    { header: 'N°', width: 36 },
    ...taxCodes.map((code) => ({ header: code, width: codeColumnWidth })),
    { header: 'Valeur Déclarée', width: 80 },
  ];

  const rows: string[][] = [];
  for (const article of declaration.articles) {
    const quantite = Math.round(article.quantite);
    const valeurDeclareePerUnit = article.valeurDeclaree / quantite;
    const perCodeAllocations = new Map<string, number[]>();
    for (const code of taxCodes) {
      const tax = article.taxes.find((t) => t.code === code);
      perCodeAllocations.set(
        code,
        tax ? allocateTaxAcrossUnits(tax.montant, quantite) : new Array(quantite).fill(0)
      );
    }
    for (let unit = 0; unit < quantite; unit++) {
      rows.push([
        article.nomArticle,
        article.hsCode,
        String(unit + 1),
        ...taxCodes.map((code) => perCodeAllocations.get(code)![unit].toFixed(2)),
        valeurDeclareePerUnit.toFixed(2),
      ]);
    }
  }

  return { columns, rows };
}

// Generates a landscape A4 PDF containing exactly the same data as the
// Excel export's "Global" sheet — every article's unit rows, colored like
// the Excel table — with a brand-colored letterhead (logo, company name,
// document title, generation date) and footer (company name) repeated on
// every page.
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

  const { columns, rows } = buildGlobalSheetTable(declaration);
  drawTable(doc, columns, rows, HEADER_HEIGHT + 12, brand600);

  const documentTitle = `Déclaration ${declaration.code} — ${declaration.redevable}`;
  const logoBuffer = parseLogoBuffer(settings.logoDataUri);
  drawLetterheadOnAllPages(doc, documentTitle, settings.companyName, logoBuffer, brand600, new Date());

  return doc;
}

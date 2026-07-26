import ExcelJS from 'exceljs';
import { describe, it, expect } from 'vitest';
import { parsePackingList } from '../../../src/parser/packingList/packingListParser.js';

async function buildFixtureBuffer(
  rows: (string | number)[][],
  headers: string[] = ['item', 'DESCRIPTION', 'color', 'pieces', 'unit', 'total', 'origin', 'HS CODE']
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.addRow(headers);
  for (const row of rows) {
    sheet.addRow(row);
  }
  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

describe('parsePackingList', () => {
  it('parses every field from real-shaped rows (matching the screenshot sample)', async () => {
    const buffer = await buildFixtureBuffer([
      ['AB0141DOAY16', 'VESTITO A FASCIA CORTO IN TULLE', 'PK2 PINK MEDIUM', 18, '€ 9,61', '€ 172,98', 'CHINA', '61044300'],
      ['BD0015DOAY16', 'BODY A COSTE', 'WH2 WHITE', 18, '€ 4,80', '€ 86,40', 'BANGLADESH', '61142000'],
    ]);

    const rows = await parsePackingList(buffer);

    expect(rows).toEqual([
      {
        item: 'AB0141DOAY16',
        description: 'VESTITO A FASCIA CORTO IN TULLE',
        color: 'PK2 PINK MEDIUM',
        pieces: 18,
        unit: 9.61,
        total: 172.98,
        origin: 'CHINA',
        hsCode: '61044300',
      },
      {
        item: 'BD0015DOAY16',
        description: 'BODY A COSTE',
        color: 'WH2 WHITE',
        pieces: 18,
        unit: 4.8,
        total: 86.4,
        origin: 'BANGLADESH',
        hsCode: '61142000',
      },
    ]);
  });

  it('matches headers case-insensitively, tolerating the real mixed-case source (item lowercase, DESCRIPTION uppercase, HS CODE uppercase)', async () => {
    const buffer = await buildFixtureBuffer(
      [['CF0093DOAY16', 'CAMICIA CROP IN SATIN', 'WH1 OFF WHITE', 2, 6.65, 13.3, 'CAMBOGIA', '62064000']],
      ['ITEM', 'Description', 'COLOR', 'Pieces', 'UNIT', 'Total', 'ORIGIN', 'hs code']
    );

    const rows = await parsePackingList(buffer);

    expect(rows).toHaveLength(1);
    expect(rows[0].item).toBe('CF0093DOAY16');
    expect(rows[0].hsCode).toBe('62064000');
  });

  it('parses unit/total from plain numeric cells, not just "€ X,XX"-formatted strings', async () => {
    const buffer = await buildFixtureBuffer([
      ['CF0093DOAY16', 'CAMICIA CROP IN SATIN', 'WH1 OFF WHITE', 2, 6.65, 13.3, 'CAMBOGIA', '62064000'],
    ]);

    const rows = await parsePackingList(buffer);

    expect(rows[0].unit).toBeCloseTo(6.65, 2);
    expect(rows[0].total).toBeCloseTo(13.3, 2);
  });

  it('skips a trailing all-blank row', async () => {
    const buffer = await buildFixtureBuffer([
      ['CF0093DOAY16', 'CAMICIA CROP IN SATIN', 'WH1 OFF WHITE', 2, '€ 6,65', '€ 13,30', 'CAMBOGIA', '62064000'],
      ['', '', '', '', '', '', '', ''],
    ]);

    const rows = await parsePackingList(buffer);

    expect(rows).toHaveLength(1);
  });

  it('preserves the source row order (no sorting/grouping by HS code)', async () => {
    const buffer = await buildFixtureBuffer([
      ['ZZ_LAST', 'ITEM Z', 'COLOR Z', 1, 1, 1, 'CHINA', '90000000'],
      ['AA_FIRST', 'ITEM A', 'COLOR A', 1, 1, 1, 'CHINA', '10000000'],
    ]);

    const rows = await parsePackingList(buffer);

    expect(rows.map((r) => r.item)).toEqual(['ZZ_LAST', 'AA_FIRST']);
  });

  it('throws when a required header is missing', async () => {
    const buffer = await buildFixtureBuffer(
      [['CF0093DOAY16', 'CAMICIA CROP IN SATIN', 'WH1 OFF WHITE', 2, '€ 6,65', '€ 13,30', 'CAMBOGIA']],
      ['item', 'DESCRIPTION', 'color', 'pieces', 'unit', 'total', 'origin']
    );

    await expect(parsePackingList(buffer)).rejects.toThrow(/hsCode/i);
  });
});

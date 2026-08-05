import { describe, it, expect } from 'vitest';
import { normalizeCountryName } from '../../src/domain/countryNames.js';

describe('normalizeCountryName', () => {
  it('translates a known English spelling to its French DUM/Liquidation equivalent', () => {
    expect(normalizeCountryName('CHINA')).toBe('CHINE');
    expect(normalizeCountryName('Italy')).toBe('ITALIE');
    expect(normalizeCountryName('india')).toBe('INDE');
  });

  it('leaves an already-French name unchanged (aside from uppercasing/trimming)', () => {
    expect(normalizeCountryName('CHINE')).toBe('CHINE');
    expect(normalizeCountryName('  Bangladesh  ')).toBe('BANGLADESH');
  });

  it('returns an unlisted country name unchanged, uppercased, so identical spellings still match', () => {
    expect(normalizeCountryName('Portugal')).toBe('PORTUGAL');
  });
});

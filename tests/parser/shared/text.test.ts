import { describe, it, expect } from 'vitest';
import { parseFrenchNumber, extractFirst } from '../../../src/parser/shared/text.js';

describe('parseFrenchNumber', () => {
  it('parses space-separated thousands with comma decimal', () => {
    expect(parseFrenchNumber('27 147,00')).toBeCloseTo(27147.0);
  });

  it('parses plain decimal with dot', () => {
    expect(parseFrenchNumber('354.000')).toBeCloseTo(354.0);
  });

  it('throws on unparseable input', () => {
    expect(() => parseFrenchNumber('abc')).toThrow();
  });

  it('parses comma decimal with no thousands separator', () => {
    expect(parseFrenchNumber('5,50')).toBeCloseTo(5.5);
  });

  it('parses a negative number with thousands separator', () => {
    expect(parseFrenchNumber('-1 234,56')).toBeCloseTo(-1234.56);
  });

  it('parses thousands separated by a non-breaking space (as found in real PDF text)', () => {
    expect(parseFrenchNumber('27 147,00')).toBeCloseTo(27147.0);
  });

  it('throws on empty input', () => {
    expect(() => parseFrenchNumber('')).toThrow();
  });

  it('throws on whitespace-only input', () => {
    expect(() => parseFrenchNumber('   ')).toThrow();
  });
});

describe('extractFirst', () => {
  it('returns the first capture group trimmed', () => {
    expect(extractFirst('REDEVABLE :  GLOBAL TRADE LOGISTICS SARL\n', /REDEVABLE\s*:\s*(.+)/)).toBe(
      'GLOBAL TRADE LOGISTICS SARL'
    );
  });

  it('returns undefined when no match', () => {
    expect(extractFirst('no match here', /FOO\s*:\s*(.+)/)).toBeUndefined();
  });

  it('returns an empty string (not undefined) when the label matches but the value is missing', () => {
    // Documents current behavior: a matching pattern with an empty capture group
    // yields '' rather than undefined, since the match itself succeeded.
    expect(extractFirst('LABEL:  \n', /LABEL\s*:\s*(.*)/)).toBe('');
  });
});

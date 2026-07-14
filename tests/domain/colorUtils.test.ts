import { describe, it, expect } from 'vitest';
import { isValidHexColor, darken, lighten } from '../../src/domain/colorUtils.js';

describe('isValidHexColor', () => {
  it('accepts a well-formed 6-digit hex color', () => {
    expect(isValidHexColor('#4f46e5')).toBe(true);
    expect(isValidHexColor('#FFFFFF')).toBe(true);
  });

  it('rejects malformed input', () => {
    expect(isValidHexColor('4f46e5')).toBe(false); // missing #
    expect(isValidHexColor('#fff')).toBe(false); // 3-digit shorthand not supported
    expect(isValidHexColor('#gggggg')).toBe(false); // invalid hex digits
    expect(isValidHexColor('red')).toBe(false); // named color
    expect(isValidHexColor('')).toBe(false);
  });
});

describe('darken', () => {
  it('moves a color toward black', () => {
    expect(darken('#ffffff', 0.5)).toBe('#808080');
    expect(darken('#4f46e5', 0)).toBe('#4f46e5'); // amount 0 is a no-op
  });
});

describe('lighten', () => {
  it('moves a color toward white', () => {
    expect(lighten('#000000', 0.5)).toBe('#808080');
    expect(lighten('#4f46e5', 0)).toBe('#4f46e5'); // amount 0 is a no-op
  });

  it('produces a very light tint near white at a high amount', () => {
    const tint = lighten('#4f46e5', 0.9);
    expect(tint).toBe('#ededfc');
  });
});
const HEX_COLOR_PATTERN = /^#[0-9a-fA-F]{6}$/;

export function isValidHexColor(value: string): boolean {
  return HEX_COLOR_PATTERN.test(value);
}

function hexToRgb(hex: string): [number, number, number] {
  const int = Number.parseInt(hex.slice(1), 16);
  return [(int >> 16) & 255, (int >> 8) & 255, int & 255];
}

function rgbToHex([r, g, b]: [number, number, number]): string {
  return '#' + [r, g, b].map((c) => Math.round(c).toString(16).padStart(2, '0')).join('');
}

function mix(hex: string, target: [number, number, number], amount: number): string {
  const [r, g, b] = hexToRgb(hex);
  return rgbToHex([
    r + (target[0] - r) * amount,
    g + (target[1] - g) * amount,
    b + (target[2] - b) * amount,
  ]);
}

/** Mixes toward black — used to derive a darker shade (e.g. a hover/700 tone) from a base color. */
export function darken(hex: string, amount: number): string {
  return mix(hex, [0, 0, 0], amount);
}

/** Mixes toward white — used to derive a light tint (e.g. a badge background) from a base color. */
export function lighten(hex: string, amount: number): string {
  return mix(hex, [255, 255, 255], amount);
}
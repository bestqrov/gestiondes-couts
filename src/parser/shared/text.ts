export function parseFrenchNumber(raw: string): number {
  const cleaned = raw
    .replace(/\s/g, '')
    .replace(/,/g, '.');
  const value = Number.parseFloat(cleaned);
  if (Number.isNaN(value)) {
    throw new Error(`Cannot parse number from "${raw}"`);
  }
  return value;
}

export function extractFirst(text: string, pattern: RegExp): string | undefined {
  const match = text.match(pattern);
  return match ? match[1].trim() : undefined;
}

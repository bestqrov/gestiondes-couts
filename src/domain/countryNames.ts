// Maps a country name spelled in another language (as seen on a supplier's
// packing list, typically English) to the French spelling the DUM/Liquidation
// use (e.g. article.pays) — needed to match a packing-list row's "origin" to
// a declaration article by country, not just by HS code, since the same HS
// position can be sourced from more than one country with different tax
// montants. Deliberately a fixed list of the countries seen in real
// documents rather than a full ISO country database — an unlisted country
// name is returned unchanged (uppercased), which still matches correctly
// whenever the packing list and the DUM happen to use the same spelling
// (e.g. "BANGLADESH", "PORTUGAL").
const COUNTRY_NAME_ALIASES: Record<string, string> = {
  CHINA: 'CHINE',
  ITALY: 'ITALIE',
  INDIA: 'INDE',
  TURKEY: 'TURQUIE',
  SPAIN: 'ESPAGNE',
  GERMANY: 'ALLEMAGNE',
  VIETNAM: 'VIETNAM',
  INDONESIA: 'INDONESIE',
  THAILAND: 'THAILANDE',
  CAMBODIA: 'CAMBODGE',
  MYANMAR: 'BIRMANIE',
  'UNITED KINGDOM': 'ROYAUME-UNI',
  UK: 'ROYAUME-UNI',
  USA: 'ETATS-UNIS',
  'UNITED STATES': 'ETATS-UNIS',
  BELGIUM: 'BELGIQUE',
  NETHERLANDS: 'PAYS-BAS',
  SWITZERLAND: 'SUISSE',
  POLAND: 'POLOGNE',
  EGYPT: 'EGYPTE',
  MOROCCO: 'MAROC',
  TUNISIA: 'TUNISIE',
  GREECE: 'GRECE',
  BULGARIA: 'BULGARIE',
  ROMANIA: 'ROUMANIE',
  'SOUTH KOREA': 'COREE DU SUD',
  KOREA: 'COREE DU SUD',
  JAPAN: 'JAPON',
  BRAZIL: 'BRESIL',
  MEXICO: 'MEXIQUE',
};

/**
 * Normalizes a country name for cross-document matching: trims, uppercases,
 * and translates known non-French spellings to their French DUM/Liquidation
 * equivalent. Two names that refer to the same country return the same
 * string, regardless of which language either was written in.
 */
export function normalizeCountryName(name: string): string {
  const upper = name.trim().toUpperCase();
  return COUNTRY_NAME_ALIASES[upper] ?? upper;
}

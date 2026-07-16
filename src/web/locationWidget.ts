// Best-effort "where is this visitor, what time/weather is it there" widget
// for the superadmin topbar — resolves the visitor's city from their IP
// address (no signup/API key required for either service used):
//   - ip-api.com's free, keyless endpoint for IP → city/lat/lon/timezone.
//   - open-meteo.com's free, keyless endpoint for lat/lon → current weather.
// Either lookup failing (offline, rate-limited, a local/private dev IP with
// no public geolocation) just means the widget renders nothing — this is
// cosmetic, never worth blocking or erroring a page over.

export interface VisitorLocation {
  city: string;
  country: string;
  lat: number;
  lon: number;
  timezone: string;
}

export interface CurrentWeather {
  tempC: number;
  description: string;
}

const LOCATION_CACHE_TTL_MS = 30 * 60 * 1000; // 30 minutes — an IP's city doesn't change minute to minute
const locationCache = new Map<string, { location: VisitorLocation | null; expiresAt: number }>();

// WMO weather codes (used by open-meteo) collapsed to short French labels —
// covers the common cases; anything unmapped falls back to "—".
const WEATHER_CODE_LABELS: Record<number, string> = {
  0: 'Ciel dégagé',
  1: 'Plutôt dégagé',
  2: 'Partiellement nuageux',
  3: 'Couvert',
  45: 'Brouillard',
  48: 'Brouillard givrant',
  51: 'Bruine légère',
  53: 'Bruine',
  55: 'Bruine forte',
  61: 'Pluie légère',
  63: 'Pluie',
  65: 'Pluie forte',
  71: 'Neige légère',
  73: 'Neige',
  75: 'Neige forte',
  80: 'Averses',
  81: 'Averses fortes',
  82: 'Averses violentes',
  95: 'Orage',
};

function isPrivateOrLocalIp(ip: string): boolean {
  return (
    ip === '::1' ||
    ip === '127.0.0.1' ||
    ip.startsWith('10.') ||
    ip.startsWith('192.168.') ||
    /^172\.(1[6-9]|2\d|3[0-1])\./.test(ip) ||
    ip.startsWith('::ffff:127.') ||
    ip === ''
  );
}

export async function getVisitorLocation(ip: string): Promise<VisitorLocation | null> {
  // A dev/internal-network IP has no meaningful public geolocation — skip
  // the network call entirely rather than let it fail slowly.
  if (isPrivateOrLocalIp(ip)) return null;

  const cached = locationCache.get(ip);
  if (cached && cached.expiresAt > Date.now()) return cached.location;

  let location: VisitorLocation | null = null;
  try {
    const response = await fetch(
      `http://ip-api.com/json/${encodeURIComponent(ip)}?fields=status,city,country,lat,lon,timezone`
    );
    if (response.ok) {
      const data = (await response.json()) as {
        status: string;
        city?: string;
        country?: string;
        lat?: number;
        lon?: number;
        timezone?: string;
      };
      if (data.status === 'success' && data.city && data.timezone) {
        location = {
          city: data.city,
          country: data.country ?? '',
          lat: data.lat ?? 0,
          lon: data.lon ?? 0,
          timezone: data.timezone,
        };
      }
    }
  } catch {
    location = null;
  }

  locationCache.set(ip, { location, expiresAt: Date.now() + LOCATION_CACHE_TTL_MS });
  return location;
}

export async function getCurrentWeather(lat: number, lon: number): Promise<CurrentWeather | null> {
  try {
    const response = await fetch(
      `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,weather_code`
    );
    if (!response.ok) return null;
    const data = (await response.json()) as {
      current?: { temperature_2m?: number; weather_code?: number };
    };
    if (data.current?.temperature_2m === undefined) return null;
    return {
      tempC: Math.round(data.current.temperature_2m),
      description: WEATHER_CODE_LABELS[data.current.weather_code ?? -1] ?? '—',
    };
  } catch {
    return null;
  }
}

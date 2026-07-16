import { describe, it, expect, vi, afterEach } from 'vitest';
import { getVisitorLocation, getCurrentWeather } from '../../src/web/locationWidget.js';

describe('getVisitorLocation', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('returns null without a network call for private/local IPs', async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal('fetch', fetchMock);

    expect(await getVisitorLocation('127.0.0.1')).toBeNull();
    expect(await getVisitorLocation('192.168.1.5')).toBeNull();
    expect(await getVisitorLocation('10.0.0.1')).toBeNull();
    expect(await getVisitorLocation('::1')).toBeNull();
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('parses a successful ip-api response into a VisitorLocation', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({
          status: 'success',
          city: 'Casablanca',
          country: 'Morocco',
          lat: 33.57,
          lon: -7.59,
          timezone: 'Africa/Casablanca',
        }),
      })
    );

    const location = await getVisitorLocation('41.140.0.1');
    expect(location).toEqual({
      city: 'Casablanca',
      country: 'Morocco',
      lat: 33.57,
      lon: -7.59,
      timezone: 'Africa/Casablanca',
    });
  });

  it('returns null when ip-api reports failure status', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue({ ok: true, json: async () => ({ status: 'fail' }) })
    );
    expect(await getVisitorLocation('41.140.0.2')).toBeNull();
  });

  it('returns null (not a thrown error) when the network call fails', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('network down')));
    await expect(getVisitorLocation('41.140.0.3')).resolves.toBeNull();
  });

  it('caches a successful lookup, not calling fetch again for the same IP', async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({
        status: 'success',
        city: 'Rabat',
        country: 'Morocco',
        lat: 34.02,
        lon: -6.83,
        timezone: 'Africa/Casablanca',
      }),
    });
    vi.stubGlobal('fetch', fetchMock);

    await getVisitorLocation('41.140.0.4');
    await getVisitorLocation('41.140.0.4');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe('getCurrentWeather', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('parses a successful open-meteo response, rounding the temperature', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ current: { temperature_2m: 21.6, weather_code: 1 } }),
      })
    );

    expect(await getCurrentWeather(33.57, -7.59)).toEqual({
      tempC: 22,
      description: 'Plutôt dégagé',
    });
  });

  it('falls back to "—" for an unrecognized weather code', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ current: { temperature_2m: 15, weather_code: 999 } }),
      })
    );
    expect((await getCurrentWeather(0, 0))?.description).toBe('—');
  });

  it('returns null when the network call fails', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('network down')));
    await expect(getCurrentWeather(0, 0)).resolves.toBeNull();
  });
});

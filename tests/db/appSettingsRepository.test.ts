import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';
import { getAppSettings, updateAppSettings } from '../../src/db/appSettingsRepository.js';

describe('appSettingsRepository', () => {
  it('returns all-null defaults when no settings have ever been saved', () => {
    const db = createDatabase(':memory:');
    expect(getAppSettings(db)).toEqual({
      companyName: null,
      logoDataUri: null,
      brandColor: null,
      fontFamily: null,
      contactEmail: null,
      contactWhatsapp: null,
    });
    db.close();
  });

  it('saves and reads back settings', () => {
    const db = createDatabase(':memory:');
    const result = updateAppSettings(db, {
      companyName: 'Acme Corp',
      brandColor: '#4f46e5',
      fontFamily: 'serif',
      logoDataUri: 'data:image/png;base64,abc123',
      contactEmail: 'contact@acme.example',
      contactWhatsapp: '+212600000000',
    });

    expect(result).toEqual({
      companyName: 'Acme Corp',
      brandColor: '#4f46e5',
      fontFamily: 'serif',
      logoDataUri: 'data:image/png;base64,abc123',
      contactEmail: 'contact@acme.example',
      contactWhatsapp: '+212600000000',
    });
    expect(getAppSettings(db)).toEqual(result);
    db.close();
  });

  it('a partial update only overwrites the fields provided, leaving the rest intact', () => {
    const db = createDatabase(':memory:');
    updateAppSettings(db, { companyName: 'Acme Corp', brandColor: '#4f46e5' });

    const result = updateAppSettings(db, { fontFamily: 'mono' });

    expect(result).toEqual({
      companyName: 'Acme Corp',
      brandColor: '#4f46e5',
      fontFamily: 'mono',
      logoDataUri: null,
      contactEmail: null,
      contactWhatsapp: null,
    });
    db.close();
  });

  it('calling updateAppSettings twice does not create a second row', () => {
    const db = createDatabase(':memory:');
    updateAppSettings(db, { companyName: 'First' });
    updateAppSettings(db, { companyName: 'Second' });

    const count = db.prepare('SELECT COUNT(*) as count FROM app_settings').get() as {
      count: number;
    };
    expect(count.count).toBe(1);
    expect(getAppSettings(db).companyName).toBe('Second');
    db.close();
  });
});

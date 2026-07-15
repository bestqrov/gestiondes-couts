import { describe, it, expect } from 'vitest';
import { createFakeCollection } from '../helpers/fakeMongoCollection.js';
import {
  getAppSettings,
  updateAppSettings,
  type AppSettingsDocument,
} from '../../src/db/appSettingsRepository.js';

function makeCollection() {
  return createFakeCollection<AppSettingsDocument>();
}

describe('appSettingsRepository', () => {
  it('returns all-null defaults when no settings have ever been saved', async () => {
    const collection = makeCollection();
    expect(await getAppSettings(collection)).toEqual({
      companyName: null,
      logoDataUri: null,
      brandColor: null,
      fontFamily: null,
      contactEmail: null,
      contactWhatsapp: null,
    });
  });

  it('saves and reads back settings', async () => {
    const collection = makeCollection();
    const result = await updateAppSettings(collection, {
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
    expect(await getAppSettings(collection)).toEqual(result);
  });

  it('a partial update only overwrites the fields provided, leaving the rest intact', async () => {
    const collection = makeCollection();
    await updateAppSettings(collection, { companyName: 'Acme Corp', brandColor: '#4f46e5' });

    const result = await updateAppSettings(collection, { fontFamily: 'mono' });

    expect(result).toEqual({
      companyName: 'Acme Corp',
      brandColor: '#4f46e5',
      fontFamily: 'mono',
      logoDataUri: null,
      contactEmail: null,
      contactWhatsapp: null,
    });
  });

  it('calling updateAppSettings twice does not create a second document', async () => {
    const collection = makeCollection();
    await updateAppSettings(collection, { companyName: 'First' });
    await updateAppSettings(collection, { companyName: 'Second' });

    expect(await collection.countDocuments()).toBe(1);
    expect((await getAppSettings(collection)).companyName).toBe('Second');
  });
});

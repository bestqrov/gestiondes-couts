import { describe, it, expect } from 'vitest';
import { createFakeCollection } from '../helpers/fakeMongoCollection.js';
import {
  getAppSettings,
  updateAppSettings,
  incrementGenerationCount,
  hashDeletePassword,
  verifyDeletePassword,
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
      generationCount: 0,
      deletePasswordHash: null,
      extraCostFields: [],
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
      generationCount: 0,
      deletePasswordHash: null,
      extraCostFields: [],
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
      generationCount: 0,
      deletePasswordHash: null,
      extraCostFields: [],
    });
  });

  it('calling updateAppSettings twice does not create a second document', async () => {
    const collection = makeCollection();
    await updateAppSettings(collection, { companyName: 'First' });
    await updateAppSettings(collection, { companyName: 'Second' });

    expect(await collection.countDocuments()).toBe(1);
    expect((await getAppSettings(collection)).companyName).toBe('Second');
  });

  it('incrementGenerationCount bumps the counter atomically, starting from 0', async () => {
    const collection = makeCollection();
    expect((await getAppSettings(collection)).generationCount).toBe(0);

    await incrementGenerationCount(collection);
    await incrementGenerationCount(collection);
    await incrementGenerationCount(collection);

    expect((await getAppSettings(collection)).generationCount).toBe(3);
  });

  it('reset (generationCount: 0) does not touch other fields', async () => {
    const collection = makeCollection();
    await updateAppSettings(collection, { companyName: 'Acme Corp' });
    await incrementGenerationCount(collection);
    await incrementGenerationCount(collection);

    const result = await updateAppSettings(collection, { generationCount: 0 });

    expect(result.generationCount).toBe(0);
    expect(result.companyName).toBe('Acme Corp');
  });

  it('hashDeletePassword/verifyDeletePassword round-trip, and reject the wrong password', async () => {
    const collection = makeCollection();
    const hash = hashDeletePassword('supersecret');
    const settings = await updateAppSettings(collection, { deletePasswordHash: hash });

    expect(verifyDeletePassword(settings, 'supersecret')).toBe(true);
    expect(verifyDeletePassword(settings, 'wrong')).toBe(false);
  });

  it('verifyDeletePassword rejects any password when none has been configured', async () => {
    const collection = makeCollection();
    const settings = await getAppSettings(collection);

    expect(verifyDeletePassword(settings, 'anything')).toBe(false);
  });
});

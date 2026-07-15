import type { Collection } from 'mongodb';

export const APP_SETTINGS_COLLECTION = 'app_settings';

// A single always-present document (fixed _id), upserted in place —
// app-wide settings have no natural owner/list semantics, unlike users or
// declarations, so there's exactly one document to read or write.
const SINGLETON_ID = 'singleton';

export interface AppSettings {
  companyName: string | null;
  logoDataUri: string | null;
  brandColor: string | null;
  fontFamily: string | null;
  contactEmail: string | null;
  contactWhatsapp: string | null;
}

export interface AppSettingsUpdate {
  companyName?: string | null;
  logoDataUri?: string | null;
  brandColor?: string | null;
  fontFamily?: string | null;
  contactEmail?: string | null;
  contactWhatsapp?: string | null;
}

export interface AppSettingsDocument extends AppSettings {
  _id: string;
  updatedAt: string;
}

export const DEFAULT_APP_SETTINGS: AppSettings = {
  companyName: null,
  logoDataUri: null,
  brandColor: null,
  fontFamily: null,
  contactEmail: null,
  contactWhatsapp: null,
};

export async function getAppSettings(
  collection: Collection<AppSettingsDocument>
): Promise<AppSettings> {
  const doc = await collection.findOne({ _id: SINGLETON_ID });
  if (!doc) return DEFAULT_APP_SETTINGS;
  return {
    companyName: doc.companyName,
    logoDataUri: doc.logoDataUri,
    brandColor: doc.brandColor,
    fontFamily: doc.fontFamily,
    contactEmail: doc.contactEmail,
    contactWhatsapp: doc.contactWhatsapp,
  };
}

export async function updateAppSettings(
  collection: Collection<AppSettingsDocument>,
  update: AppSettingsUpdate
): Promise<AppSettings> {
  const current = await getAppSettings(collection);
  const merged: AppSettings = {
    companyName: update.companyName !== undefined ? update.companyName : current.companyName,
    logoDataUri: update.logoDataUri !== undefined ? update.logoDataUri : current.logoDataUri,
    brandColor: update.brandColor !== undefined ? update.brandColor : current.brandColor,
    fontFamily: update.fontFamily !== undefined ? update.fontFamily : current.fontFamily,
    contactEmail: update.contactEmail !== undefined ? update.contactEmail : current.contactEmail,
    contactWhatsapp:
      update.contactWhatsapp !== undefined ? update.contactWhatsapp : current.contactWhatsapp,
  };

  await collection.updateOne(
    { _id: SINGLETON_ID },
    { $set: { ...merged, updatedAt: new Date().toISOString() } },
    { upsert: true }
  );

  return merged;
}

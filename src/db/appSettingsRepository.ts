import bcrypt from 'bcryptjs';
import type { Collection } from 'mongodb';

export const APP_SETTINGS_COLLECTION = 'app_settings';

// A single always-present document (fixed _id), upserted in place —
// app-wide settings have no natural owner/list semantics, unlike users or
// declarations, so there's exactly one document to read or write.
const SINGLETON_ID = 'singleton';

const SALT_ROUNDS = 10;

export interface AppSettings {
  companyName: string | null;
  logoDataUri: string | null;
  brandColor: string | null;
  fontFamily: string | null;
  contactEmail: string | null;
  contactWhatsapp: string | null;
  // Display-only tally of successful /generate calls, separate from the
  // transactions collection (the real, un-resettable declaration history
  // Historique/Costs pages rely on) — a superadmin can zero this out from
  // the dashboard without touching that history.
  generationCount: number;
  // bcrypt hash of the password required to reset generationCount — never
  // stored/returned in plaintext, mirroring usersRepository's passwordHash.
  deletePasswordHash: string | null;
}

export interface AppSettingsUpdate {
  companyName?: string | null;
  logoDataUri?: string | null;
  brandColor?: string | null;
  fontFamily?: string | null;
  contactEmail?: string | null;
  contactWhatsapp?: string | null;
  generationCount?: number;
  deletePasswordHash?: string | null;
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
  generationCount: 0,
  deletePasswordHash: null,
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
    // Both fields are new — documents written before this feature existed
    // won't have them, so fall back to the same defaults a fresh singleton
    // would get.
    generationCount: doc.generationCount ?? 0,
    deletePasswordHash: doc.deletePasswordHash ?? null,
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
    generationCount:
      update.generationCount !== undefined ? update.generationCount : current.generationCount,
    deletePasswordHash:
      update.deletePasswordHash !== undefined
        ? update.deletePasswordHash
        : current.deletePasswordHash,
  };

  await collection.updateOne(
    { _id: SINGLETON_ID },
    { $set: { ...merged, updatedAt: new Date().toISOString() } },
    { upsert: true }
  );

  return merged;
}

export async function incrementGenerationCount(
  collection: Collection<AppSettingsDocument>
): Promise<void> {
  const current = await getAppSettings(collection);
  await updateAppSettings(collection, { generationCount: current.generationCount + 1 });
}

export function hashDeletePassword(password: string): string {
  return bcrypt.hashSync(password, SALT_ROUNDS);
}

export function verifyDeletePassword(settings: AppSettings, password: string): boolean {
  return settings.deletePasswordHash !== null && bcrypt.compareSync(password, settings.deletePasswordHash);
}

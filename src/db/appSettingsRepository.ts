import type Database from 'better-sqlite3';

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

interface SettingsRow {
  company_name: string | null;
  logo_data_uri: string | null;
  brand_color: string | null;
  font_family: string | null;
  contact_email: string | null;
  contact_whatsapp: string | null;
}

const DEFAULTS: AppSettings = {
  companyName: null,
  logoDataUri: null,
  brandColor: null,
  fontFamily: null,
  contactEmail: null,
  contactWhatsapp: null,
};

export function getAppSettings(db: Database.Database): AppSettings {
  const row = db.prepare('SELECT * FROM app_settings WHERE id = 1').get() as
    | SettingsRow
    | undefined;
  if (!row) return DEFAULTS;
  return {
    companyName: row.company_name,
    logoDataUri: row.logo_data_uri,
    brandColor: row.brand_color,
    fontFamily: row.font_family,
    contactEmail: row.contact_email,
    contactWhatsapp: row.contact_whatsapp,
  };
}

// A single always-present row (id = 1), upserted in place — app-wide
// settings have no natural owner/list semantics, unlike users or
// declarations, so there's exactly one row to read or write.
export function updateAppSettings(db: Database.Database, update: AppSettingsUpdate): AppSettings {
  const current = getAppSettings(db);
  const merged: AppSettings = {
    companyName: update.companyName !== undefined ? update.companyName : current.companyName,
    logoDataUri: update.logoDataUri !== undefined ? update.logoDataUri : current.logoDataUri,
    brandColor: update.brandColor !== undefined ? update.brandColor : current.brandColor,
    fontFamily: update.fontFamily !== undefined ? update.fontFamily : current.fontFamily,
    contactEmail: update.contactEmail !== undefined ? update.contactEmail : current.contactEmail,
    contactWhatsapp:
      update.contactWhatsapp !== undefined ? update.contactWhatsapp : current.contactWhatsapp,
  };

  db.prepare(
    `INSERT INTO app_settings (
       id, company_name, logo_data_uri, brand_color, font_family,
       contact_email, contact_whatsapp, updated_at
     )
     VALUES (1, ?, ?, ?, ?, ?, ?, ?)
     ON CONFLICT (id) DO UPDATE SET
       company_name = excluded.company_name,
       logo_data_uri = excluded.logo_data_uri,
       brand_color = excluded.brand_color,
       font_family = excluded.font_family,
       contact_email = excluded.contact_email,
       contact_whatsapp = excluded.contact_whatsapp,
       updated_at = excluded.updated_at`
  ).run(
    merged.companyName,
    merged.logoDataUri,
    merged.brandColor,
    merged.fontFamily,
    merged.contactEmail,
    merged.contactWhatsapp,
    new Date().toISOString()
  );

  return merged;
}

import { defaultEmailSettings, type EmailSettings } from './email';
import { defaultLayout, type Column, type Layout } from './layout';

export interface Settings {
  version: 1;
  layout: Layout;
  flagHospitals: string[];
  hospitalCodes: string[];
  email: EmailSettings;
}

export type KeyValueStore = Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>;

const STORAGE_KEY = 'theatre-list-prepper.settings';

export function defaultSettings(): Settings {
  return {
    version: 1,
    layout: defaultLayout(),
    flagHospitals: ['POWH', 'STG'],
    hospitalCodes: ['SSEH', 'POWH', 'STG'],
    email: defaultEmailSettings(),
  };
}

export function loadSettings(store: KeyValueStore | null): Settings {
  try {
    const text = store?.getItem(STORAGE_KEY);
    return text ? importSettings(text) : defaultSettings();
  } catch {
    return defaultSettings();
  }
}

export function saveSettings(store: KeyValueStore | null, settings: Settings): void {
  try {
    store?.setItem(STORAGE_KEY, exportSettings(settings));
  } catch {
    // Storage unavailable (private window, blocked site data): settings last for this visit only.
  }
}

export function clearSettings(store: KeyValueStore | null): void {
  try {
    store?.removeItem(STORAGE_KEY);
  } catch {
    // Nothing stored to clear.
  }
}

export function exportSettings(settings: Settings): string {
  return JSON.stringify(settings, null, 2);
}

/** Reads settings JSON, filling anything missing or malformed with defaults. */
export function importSettings(text: string): Settings {
  const data: unknown = JSON.parse(text);
  if (!isObject(data)) throw new Error('This file isn’t a Theatre List Prepper settings file.');
  const defaults = defaultSettings();

  const layout = isObject(data.layout) ? data.layout : {};
  const columns = Array.isArray(layout.columns)
    ? layout.columns.filter(isObject).map(
        (c): Column => ({ heading: String(c.heading ?? ''), source: String(c.source ?? 'blank') }),
      )
    : defaults.layout.columns;

  return {
    version: 1,
    layout: {
      columns,
      includeHeadings: typeof layout.includeHeadings === 'boolean' ? layout.includeHeadings : defaults.layout.includeHeadings,
      pasteStart: typeof layout.pasteStart === 'string' ? layout.pasteStart : defaults.layout.pasteStart,
    },
    flagHospitals: Array.isArray(data.flagHospitals) ? stringList(data.flagHospitals) : defaults.flagHospitals,
    hospitalCodes: Array.isArray(data.hospitalCodes) ? stringList(data.hospitalCodes) : defaults.hospitalCodes,
    email: importEmail(data.email, defaults.email),
  };
}

function importEmail(value: unknown, defaults: EmailSettings): EmailSettings {
  if (!isObject(value)) return defaults;
  const text = (v: unknown, fallback: string) => (typeof v === 'string' && v.trim() ? v : fallback);
  return {
    to: stringList(value.to),
    cc: stringList(value.cc),
    name: typeof value.name === 'string' ? value.name : defaults.name,
    subject: text(value.subject, defaults.subject),
    message: text(value.message, defaults.message),
    outlookUrl: text(value.outlookUrl, defaults.outlookUrl),
  };
}

export function parseList(text: string): string[] {
  return text
    .split(/[,;\n]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function stringList(value: unknown): string[] {
  return Array.isArray(value) ? value.map((v) => String(v).trim()).filter(Boolean) : [];
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

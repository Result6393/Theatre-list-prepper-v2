// "Download offline copy with my settings": saves this page as one HTML file that has the user's
// settings (and kept Excel template) built in, so it works on a computer that blocks the web version
// or wipes browser data. The copy is made from the page as it was before the app drew anything.

const SETTINGS_ID = 'embedded-settings';
const TEMPLATE_ID = 'embedded-template';

/** The page as it was before the app drew anything. Call once, before the first render. */
export function capturePristine(): string {
  return '<!doctype html>\n' + document.documentElement.outerHTML;
}

function embeddedText(id: string): string | null {
  const text = document.getElementById(id)?.textContent?.trim();
  return text || null;
}

/** Settings baked into this file when it was saved from the app, if any. */
export function embeddedSettings(): string | null {
  return embeddedText(SETTINGS_ID);
}

/** When (ms since 1970) the settings were built into this file; 0 for a file saved before this was recorded. */
export function embeddedSavedAt(): number {
  const n = Number(document.getElementById(SETTINGS_ID)?.getAttribute('data-saved-at'));
  return Number.isFinite(n) && n > 0 ? n : 0;
}

/**
 * What to start with: the browser's own copy of something (settings, the kept template) or the one
 * built into an offline copy. The built-in one wins when the browser has none, or when the file was
 * saved after the browser's copy last changed (a newer file opened where an older one was used
 * before). Otherwise the browser's own wins, so edits made inside the file are kept.
 */
export function pickNewer(
  stored: string | null,
  storedUpdated: number,
  embedded: string | null,
  embeddedSavedAt: number,
): { text: string | null; fromEmbedded: boolean } {
  if (!embedded) return { text: stored || null, fromEmbedded: false };
  if (!stored) return { text: embedded, fromEmbedded: true };
  if (embeddedSavedAt > storedUpdated) return { text: embedded, fromEmbedded: true };
  return { text: stored, fromEmbedded: false };
}

/**
 * The Excel template to build into an offline copy: one chosen just for the copy, else the kept
 * one (unless left out), else none.
 */
export function templateForOffline(chosen: string | null, kept: string | null, includeKept: boolean): string | null {
  return chosen || (includeKept ? kept : null);
}

/** The name inside a stored template ({ name, data } as JSON), for showing. */
export function templateName(json: string | null): string | null {
  if (!json) return null;
  try {
    const name = (JSON.parse(json) as { name?: unknown } | null)?.name;
    return typeof name === 'string' && name ? name : null;
  } catch {
    return null;
  }
}

/** The kept Excel template baked into this file ({ name, data } as JSON), if any. */
export function embeddedTemplate(): string | null {
  return embeddedText(TEMPLATE_ID);
}

/**
 * The pristine page with settings (and optionally the kept template) built in. Done on a parsed
 * copy, not by searching the text, because the app's own code mentions these tags.
 */
export function withEmbedded(pristine: string, settingsJson: string, templateJson: string | null = null, savedAt = Date.now()): string {
  const doc = new DOMParser().parseFromString(pristine, 'text/html');
  const embed = (id: string, json: string | null) => {
    doc.getElementById(id)?.remove();
    if (!json) return;
    const tag = doc.createElement('script');
    tag.id = id;
    tag.type = 'application/json';
    // So a browser that already holds older settings can tell this copy is newer.
    tag.setAttribute('data-saved-at', String(savedAt));
    // "<" can't appear in a script's text, or a "</script" in the data would end it early.
    tag.textContent = json.replace(/</g, '\\u003c');
    doc.head.prepend(tag);
  };
  embed(TEMPLATE_ID, templateJson);
  embed(SETTINGS_ID, settingsJson);
  return '<!doctype html>\n' + doc.documentElement.outerHTML;
}

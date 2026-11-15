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

/** The kept Excel template baked into this file ({ name, data } as JSON), if any. */
export function embeddedTemplate(): string | null {
  return embeddedText(TEMPLATE_ID);
}

/**
 * The pristine page with settings (and optionally the kept template) built in. Done on a parsed
 * copy, not by searching the text, because the app's own code mentions these tags.
 */
export function withEmbedded(pristine: string, settingsJson: string, templateJson: string | null = null): string {
  const doc = new DOMParser().parseFromString(pristine, 'text/html');
  const embed = (id: string, json: string | null) => {
    doc.getElementById(id)?.remove();
    if (!json) return;
    const tag = doc.createElement('script');
    tag.id = id;
    tag.type = 'application/json';
    // "<" can't appear in a script's text, or a "</script" in the data would end it early.
    tag.textContent = json.replace(/</g, '\\u003c');
    doc.head.prepend(tag);
  };
  embed(TEMPLATE_ID, templateJson);
  embed(SETTINGS_ID, settingsJson);
  return '<!doctype html>\n' + doc.documentElement.outerHTML;
}

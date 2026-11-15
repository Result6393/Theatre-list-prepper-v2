import { describe, expect, it } from 'vitest';
import { pickNewer, templateForOffline, templateName } from '../src/offline';

describe('which copy an offline file starts with', () => {
  const stored = '{"stored":true}';
  const built = '{"built":true}';

  it("uses the browser's own copy when the file has none built in", () => {
    expect(pickNewer(stored, 100, null, 0)).toEqual({ text: stored, fromEmbedded: false });
    expect(pickNewer(null, 0, null, 0)).toEqual({ text: null, fromEmbedded: false });
  });

  it('uses the built-in copy when the browser has none', () => {
    expect(pickNewer(null, 0, built, 50)).toEqual({ text: built, fromEmbedded: true });
    expect(pickNewer(null, 0, built, 0)).toEqual({ text: built, fromEmbedded: true });
  });

  it("prefers a built-in copy saved after the browser's last changed", () => {
    expect(pickNewer(stored, 100, built, 200)).toEqual({ text: built, fromEmbedded: true });
  });

  it('keeps edits made in the browser after the file was saved', () => {
    expect(pickNewer(stored, 300, built, 200)).toEqual({ text: stored, fromEmbedded: false });
  });

  it("keeps the browser's copy for a file saved before the time was recorded", () => {
    expect(pickNewer(stored, 100, built, 0)).toEqual({ text: stored, fromEmbedded: false });
  });
});

describe('which Excel template goes into an offline copy', () => {
  const kept = '{"name":"kept.xlsx","data":"QQ=="}';
  const chosen = '{"name":"chosen.xlsx","data":"Qg=="}';

  it('uses a template chosen for the copy first', () => {
    expect(templateForOffline(chosen, null, true)).toBe(chosen);
    expect(templateForOffline(chosen, kept, true)).toBe(chosen);
  });

  it('otherwise the kept one, unless it is left out', () => {
    expect(templateForOffline(null, kept, true)).toBe(kept);
    expect(templateForOffline(null, kept, false)).toBeNull();
  });

  it('is none when there is nothing to include', () => {
    expect(templateForOffline(null, null, true)).toBeNull();
  });

  it('reads a template name for showing', () => {
    expect(templateName(kept)).toBe('kept.xlsx');
    expect(templateName(null)).toBeNull();
    expect(templateName('not json')).toBeNull();
    expect(templateName('{"data":"x"}')).toBeNull();
  });
});

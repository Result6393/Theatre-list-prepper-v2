import { describe, expect, it } from 'vitest';
import { pickNewer } from '../src/offline';

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

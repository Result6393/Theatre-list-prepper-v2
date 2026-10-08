import { describe, expect, it } from 'vitest';
import {
  clearSettings, defaultSettings, exportSettings, importSettings, loadSettings, saveSettings, type KeyValueStore,
} from '../src/settings';

function memoryStore(): KeyValueStore {
  const data = new Map<string, string>();
  return {
    getItem: (k) => data.get(k) ?? null,
    setItem: (k, v) => void data.set(k, v),
    removeItem: (k) => void data.delete(k),
  };
}

describe('settings storage', () => {
  it('starts empty with the default layout', () => {
    const s = loadSettings(memoryStore());
    expect(s.layout.columns).toHaveLength(10);
    expect(s.flagHospitals).toEqual(['POWH', 'STG']);
  });

  it('saves, loads and clears', () => {
    const store = memoryStore();
    const s = defaultSettings();
    s.flagHospitals = ['ABC'];
    saveSettings(store, s);
    expect(loadSettings(store)).toEqual(s);
    clearSettings(store);
    expect(loadSettings(store)).toEqual(defaultSettings());
  });

  it('round-trips through export/import and repairs partial files', () => {
    const s = defaultSettings();
    s.layout.columns = [{ heading: 'Time', source: 'field:time' }];
    expect(importSettings(exportSettings(s))).toEqual(s);
    const partial = importSettings('{"flagHospitals":["abc", 3],"layout":{"columns":[{"heading":"X"}]}}');
    expect(partial.flagHospitals).toEqual(['abc', '3']);
    expect(partial.layout.columns).toEqual([{ heading: 'X', source: 'blank' }]);
    expect(importSettings('{}').layout).toEqual(defaultSettings().layout);
    expect(() => importSettings('[1,2]')).toThrow();
  });

  it('falls back to defaults when storage is broken', () => {
    const store = memoryStore();
    store.setItem('theatre-list-prepper.settings', '{not json');
    expect(loadSettings(store)).toEqual(defaultSettings());
  });
});

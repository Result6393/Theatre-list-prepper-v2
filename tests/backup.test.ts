import { describe, expect, it } from 'vitest';
import { lensEmailText, lensLines, lensSummary } from '../src/lenslist';
import { defaultPrompts, readBackup, readLens, readTarget, writeBackup, writeLens, writeTarget } from '../src/prompts';
import { defaultSettings, exportSettings, importSettings } from '../src/settings';

const { lenses } = defaultPrompts();
const picked = writeTarget(writeLens('', lenses, { lens: 'DCB00', extra: '', power: '21.5' }), lenses, 'plano');

describe('backup lens', () => {
  it('goes on its own line under the lens and target, and reads back', () => {
    const cell = writeBackup(picked, lenses, { model: 'MA60AC', power: '20.5' });
    expect(cell).toBe('DCB00 +21.5D\nTarget plano\nBackup: MA60AC +20.5D');
    expect(readBackup(cell)).toEqual({ model: 'MA60AC', power: '+20.5D' });
    // The first lens and its target are read as before.
    expect(readLens(cell, lenses)).toMatchObject({ lens: 'DCB00', power: '+21.5D' });
    expect(readTarget(cell)).toBe('plano');
  });

  it('replaces the old backup, keeps your own notes, and comes off again', () => {
    const cell = writeBackup(`${picked}\nCTR on standby`, lenses, { model: 'MA60AC', power: '20' });
    const changed = writeBackup(cell, lenses, { model: 'AR40e', power: '19.5' });
    expect(changed).toBe('DCB00 +21.5D\nTarget plano\nBackup: AR40e +19.5D\nCTR on standby');
    expect(writeBackup(changed, lenses, null)).toBe(`${picked}\nCTR on standby`);
  });

  it('survives the first lens being changed or removed, and reads what you typed', () => {
    const cell = writeBackup(picked, lenses, { model: 'MA60AC', power: '20' });
    expect(writeLens(cell, lenses, { lens: 'DIB00', extra: '', power: '22' })).toBe('DIB00 +22.0D\nTarget plano\nBackup: MA60AC +20.0D');
    expect(readBackup(writeLens(cell, lenses, null))).toEqual({ model: 'MA60AC', power: '+20.0D' });
    expect(readBackup('backup lens - MN60AC 18')).toEqual({ model: 'MN60AC', power: '18' });
    expect(readBackup(picked)).toBeNull();
  });

  it('is on the lens list as its own line, and counted with the lenses needed', () => {
    const item = (who: string, cell: string) => ({ who, eye: 'RIGHT', cell });
    const { lines, missing } = lensLines(
      [item('1000001 TESTER, Alice', writeBackup(picked, lenses, { model: 'MA60AC', power: '20' })), item('1000002 SAMPLE, Bob', picked), item('1000003 EXAMPLE, Carol', 'Backup: MA60AC +20.0D')],
      lenses,
    );
    expect(lines.map((l) => [l.who.split(' ')[0], l.model, l.power, Boolean(l.backup)])).toEqual([
      ['1000001', 'DCB00', '+21.5D', false],
      ['1000001', 'MA60AC', '+20.0D', true],
      ['1000002', 'DCB00', '+21.5D', false],
      ['1000003', 'MA60AC', '+20.0D', true],
    ]);
    // Only a backup, no first lens: still to pick.
    expect(missing.map((m) => m.who)).toEqual(['1000003 EXAMPLE, Carol']);
    expect(lensSummary(lines)).toEqual([
      { model: 'DCB00', power: '+21.5D', count: 2 },
      { model: 'MA60AC', power: '+20.0D', count: 2 },
    ]);
    expect(lensEmailText(lines, { list: 'Smith', day: '', name: '' })).toContain('RE · 1000001 TESTER, Alice · Backup: MA60AC · +20.0D');
  });

  it('has its own list of lenses in settings, filled in for settings saved before it existed', () => {
    const old = JSON.parse(exportSettings(defaultSettings()));
    delete old.prompts.backupLenses;
    expect(importSettings(JSON.stringify(old)).prompts.backupLenses).toEqual(defaultPrompts().backupLenses);
    old.prompts.backupLenses = ['MN60AC'];
    expect(importSettings(JSON.stringify(old)).prompts.backupLenses).toEqual(['MN60AC']);
  });
});

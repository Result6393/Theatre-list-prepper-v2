import { describe, expect, it } from 'vitest';
import { lensLines, lensListText, shortEye, type LensItem } from '../src/lenslist';
import { defaultPrompts, writeLens, writeTarget } from '../src/prompts';

const { lenses } = defaultPrompts();
const item = (who: string, eye: string, cell: string): LensItem => ({ who, eye, cell });
const picked = (code: string, power: string, target: string | null = null) =>
  writeTarget(writeLens('', lenses, { lens: code, extra: '', power }), lenses, target);

describe('lens list for the theatre coordinator', () => {
  it('gives each patient their eye, lens, power and target', () => {
    const { lines, missing } = lensLines(
      [
        item('1000001 ALPHA, Ann', 'RIGHT', picked('DCB00', '21.5', 'plano')),
        item('1000002 BRAVO, Ben', 'LEFT', picked('DIB00', '-3', '-0.5')),
      ],
      lenses,
    );
    expect(missing).toEqual([]);
    expect(lensListText(lines, 'Lenses for Tuesday 13/10')).toBe(
      'Lenses for Tuesday 13/10\n\nRE · 1000001 ALPHA, Ann · DCB00 +21.5D · Target plano\nLE · 1000002 BRAVO, Ben · DIB00 -3.0D · Target -0.50',
    );
  });

  it('writes a lens with a typed number, such as a toric', () => {
    const cell = writeLens('', lenses, { lens: 'CNA0Tx', extra: '4', power: '20' });
    expect(lensLines([item('3000001 CHARLIE, Cat', 'BOTH', cell)], lenses).lines[0]).toEqual({
      eye: 'BE', who: '3000001 CHARLIE, Cat', lens: 'CNA0T4 +20.0D', target: '',
    });
  });

  it('leaves patients with no lens out, and says who', () => {
    const { lines, missing } = lensLines([item('A', 'RIGHT', picked('DCB00', '20')), item('B', 'LEFT', ''), item('C', 'LEFT', 'Target plano')], lenses);
    expect(lines.map((l) => l.who)).toEqual(['A']);
    expect(missing.map((m) => m.who)).toEqual(['B', 'C']);
  });

  it('uses what you typed when it is not one of your listed lenses', () => {
    const typed = 'Aim plano\nZCB00 +22.00 (-0.32)\nTarget -0.25';
    expect(lensLines([item('D', 'RIGHT', typed)], lenses).lines[0]).toMatchObject({ lens: 'Aim plano; ZCB00 +22.00 (-0.32)', target: 'Target -0.25' });
  });

  it('shortens the eye, and leaves it out when there is none', () => {
    expect(['RIGHT', 'left', 'Both', '', 'x'].map(shortEye)).toEqual(['RE', 'LE', 'BE', '', '']);
    const { lines } = lensLines([item('E', '', picked('DCB00', '20'))], lenses);
    expect(lensListText(lines, 'H')).toBe('H\n\nE · DCB00 +20.0D');
  });
});

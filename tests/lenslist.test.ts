import { describe, expect, it } from 'vitest';
import { lensEmailHtml, lensEmailText, lensLines, lensSummary, shortEye, type LensItem } from '../src/lenslist';
import { defaultPrompts, writeLens, writeTarget } from '../src/prompts';

const { lenses } = defaultPrompts();
const item = (who: string, eye: string, cell: string): LensItem => ({ who, eye, cell });
const picked = (code: string, power: string, target: string | null = null) =>
  writeTarget(writeLens('', lenses, { lens: code, extra: '', power }), lenses, target);
const info = { list: 'Smith', day: 'Tuesday 13/10', name: 'Sam' };

describe('lens email for the theatre coordinator', () => {
  it('gives each patient their eye, lens, power and target', () => {
    const { lines, missing } = lensLines(
      [
        item('1000001 ALPHA, Ann', 'RIGHT', picked('DCB00', '21.5', 'plano')),
        item('1000002 BRAVO, Ben', 'LEFT', picked('DIB00', '-3', '-0.5')),
      ],
      lenses,
    );
    expect(missing).toEqual([]);
    expect(lines).toEqual([
      { eye: 'RE', who: '1000001 ALPHA, Ann', model: 'DCB00', power: '+21.5D', target: 'plano' },
      { eye: 'LE', who: '1000002 BRAVO, Ben', model: 'DIB00', power: '-3.0D', target: '-0.50' },
    ]);
  });

  it('writes a lens with a typed number, such as a toric', () => {
    const cell = writeLens('', lenses, { lens: 'CNA0Tx', extra: '4', power: '20' });
    expect(lensLines([item('3000001 CHARLIE, Cat', 'BOTH', cell)], lenses).lines[0]).toEqual({
      eye: 'BE', who: '3000001 CHARLIE, Cat', model: 'CNA0T4', power: '+20.0D', target: '',
    });
  });

  it('leaves patients with no lens out, and says who', () => {
    const { lines, missing } = lensLines([item('A', 'RIGHT', picked('DCB00', '20')), item('B', 'LEFT', ''), item('C', 'LEFT', 'Target plano')], lenses);
    expect(lines.map((l) => l.who)).toEqual(['A']);
    expect(missing.map((m) => m.who)).toEqual(['B', 'C']);
  });

  it('uses what you typed when it is not one of your listed lenses', () => {
    const typed = 'Aim plano\nZCB00 +22.00 (-0.32)\nTarget -0.25';
    expect(lensLines([item('D', 'RIGHT', typed)], lenses).lines[0]).toMatchObject({ model: 'Aim plano; ZCB00 +22.00 (-0.32)', power: '', target: '-0.25' });
  });

  it('shortens the eye', () => {
    expect(['RIGHT', 'left', 'Both', '', 'x'].map(shortEye)).toEqual(['RE', 'LE', 'BE', '', '']);
  });

  it('counts the lenses needed by model, then power', () => {
    const { lines } = lensLines(
      [
        item('A', 'RIGHT', picked('DIB00', '20')),
        item('B', 'LEFT', picked('DCB00', '22')),
        item('C', 'LEFT', picked('DCB00', '-1')),
        item('D', 'RIGHT', picked('DCB00', '22')),
        item('E', 'RIGHT', picked('DCB00', '9.5')),
      ],
      lenses,
    );
    expect(lensSummary(lines)).toEqual([
      { model: 'DCB00', power: '-1.0D', count: 1 },
      { model: 'DCB00', power: '+9.5D', count: 1 },
      { model: 'DCB00', power: '+22.0D', count: 2 },
      { model: 'DIB00', power: '+20.0D', count: 1 },
    ]);
  });

  it('drafts the whole email as text', () => {
    const { lines } = lensLines(
      [item('1000001 ALPHA, Ann', 'RIGHT', picked('DCB00', '21.5', 'plano')), item('1000002 BRAVO, Ben', 'LEFT', picked('DCB00', '21.5'))],
      lenses,
    );
    expect(lensEmailText(lines, info)).toBe(
      [
        'Hi All,',
        '',
        'Here is the list of lenses for the Smith list on Tuesday 13/10.',
        '',
        'RE · 1000001 ALPHA, Ann · DCB00 · +21.5D · Target plano',
        'LE · 1000002 BRAVO, Ben · DCB00 · +21.5D',
        '',
        'In summary, here are the lenses needed:',
        '',
        'DCB00 · +21.5D × 2',
        '',
        'All the best,',
        'Sam',
      ].join('\n'),
    );
  });

  it('puts the patients and the summary in tables in the HTML', () => {
    const { lines } = lensLines([item('1000001 ALPHA & <B>', 'RIGHT', picked('DCB00', '21.5', 'plano'))], lenses);
    const html = lensEmailHtml(lines, { ...info, name: '' });
    const rows = [...html.matchAll(/<tr>(.*?)<\/tr>/g)].map((m) => [...m[1].matchAll(/<t[hd][^>]*>(.*?)<\/t[hd]>/g)].map((c) => c[1]));
    expect(rows).toEqual([
      ['Eye', 'Patient', 'Lens', 'Power', 'Target'],
      ['RE', '1000001 ALPHA &amp; &lt;B&gt;', 'DCB00', '+21.5D', 'plano'],
      ['Lens', 'Power', 'Number'],
      ['DCB00', '+21.5D', '1'],
    ]);
    expect(html.match(/<table/g)).toHaveLength(2);
    expect(html.startsWith('<div>Hi All,</div>')).toBe(true);
    expect(html.endsWith('<div>All the best,</div>')).toBe(true);
  });

  it('words the intro without a list name or day', () => {
    const { lines } = lensLines([item('A', 'RIGHT', picked('DCB00', '20'))], lenses);
    expect(lensEmailText(lines, { list: '', day: '', name: '' })).toContain('Here is the list of lenses for the list.');
  });
});

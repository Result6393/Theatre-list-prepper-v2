import { describe, expect, it } from 'vitest';
import { checkLens, combineReadings, parseConstants, readNumber, readReport, simulateLens, type OcrWord, type PickedLens, type Simulation } from '../src/iolreport';
import { srktPower, srktRefraction } from '../src/srkt';

/**
 * An invented report laid out like an IOL calculation printout, as OCR words: each line is a row,
 * and each piece of text sits at the character column given (10px per character).
 */
function page(lines: [number, string][][]): OcrWord[] {
  return lines.flatMap((parts, row) =>
    parts.flatMap(([col, text]) =>
      [...text.matchAll(/\S+/g)].map((m) => ({ text: m[0], x0: (col + m.index!) * 10, x1: (col + m.index! + m[0].length) * 10, y0: row * 30, y1: row * 30 + 20 })),
    ),
  );
}

/** The same text in each of the four lens columns (two per eye), or across each eye for a toric. */
const four = (texts: string[]): [number, string][] => texts.map((t, i) => [i * 42 + 2, t]);
const two = (left: string, right: string): [number, string][] => [[2, left], [86, right]];

const REPORT = page([
  [[0, 'Patient'], [12, 'SMITH, Jan']],
  [[0, 'Date of birth'], [14, '01/01/1950'], [40, 'Gender Female']],
  [[0, 'Patient ID'], [12, '1000001'], [110, 'Some Clinic']],
  [[0, 'Date of measurement:'], [22, '01/02/2026']],
  [[0, '!'], [3, 'Axial length of the right eye is 0.40 mm longer than axial length of the left eye.']],
  [[4, 'OD'], [70, 'IOL calculation'], [160, 'OS']],
  two('Target ref.: -0.50 D', 'Target ref.: -0.25 D'),
  two('AL: 23.50 mm     SD: 10 um', 'AL: 23.20 mm     SD: 12 um'),
  two('SE: 4450 D   SD: 0.02 D   K1: 44.20 D @ 96°', 'SE: 44.60 D   SD: 0.01 D   K1: 44.40 D @ 68°'),
  two('ΔK: +0.47 D @ 6°   K2: 44.67 D @ 6°', 'AK: +0.54 D @158°   K2: 44.94 D @158°'),
  four(['TK Acme Vision Te', 'TK Acme Vision Se', 'TK Acme Vision Te', 'TK Acme Vision Se']),
  four(['cnis DCBOO Preload', 'nsar AR40 3-Piece', 'cnis DCB00 Preload', 'nsar AR40 3-Piece']),
  four(['- Barrett Universal II TK -', '- Barrett Universal II TK -', '- Barrett Universal II TK -', '- Barrett Universal II TK -']),
  four(['LF: +2.09 DF: +4.0', 'LF: +1.73 DF: +1.5', 'LF: +2.09 DF: +4.0', 'LF: +1.73 DF: +1.5']),
  four(['IOL (D)    Ref (D)', 'IOL (D)    Ref (D)', 'IOL (D)    Ref (D)', 'I0L (D)    Ref (D)']),
  four(['+21.50     -1.03', '+21.00     -1.18', '+23.00     -1.23', '+22.50     -1.41']),
  four(['+2100      -0.68', '+20.50     -0.81', '+22.50     -0.87', '+22.00     -1.04']),
  four(['+20.50     -0.33', '+20.00     -0.45', '+22.00     -0.51', '+21.50     -0.67']),
  four(['+20.00     +0.02', '+19.50     -0.09', '421.50     -0.16', '+21.00     -0.31']),
  four(['+20.03 Emmetropia', '+19.37 Emmetropia', '+21.27 Emmetropia', '+20.58 Emmetropia']),
  two('TK      Acme Vision Tecnis Toric II ZCU', 'TK]     Acme Vision Tecnis Toric II ZCU'),
  two('         - Barrett Toric TK -', '         - Barrett Toric TK -'),
  two('IOL SE IOL Cyl IOL axis Ref SE Ref Sph', 'IOL SE IOL Cyl IOL axis Ref SE Ref Sph'),
  two('100 +21.00 +1.00 3° -0.68 -0.81 +0.27 3°', '150 +22.50 +1.50 167° -0.87 -0.99 +0.25 167°'),
  two('100 +20.50 +1.00 3* -0.33 -046 +0.27 3°', ']150 +22.00 +1.50 167° -0.51 -0.63 +0.25 167°'),
  two('--- +1.36 3° Emmetropia', '--- +1.32 167° Emmetropia'),
]);

const lens = (over: Partial<PickedLens>): PickedLens => ({ eye: 'RIGHT', code: 'DCB00', power: 20.5, target: -0.5, mrn: '1000001', surname: 'SMITH', givenName: 'Jan', ...over });

describe('reading an IOL calculation report', () => {
  const report = readReport(REPORT)!;

  it('reads the patient, the report’s warning and each eye’s target', () => {
    expect(report).toMatchObject({ name: 'SMITH, Jan', id: '1000001' });
    expect(report.warnings).toEqual(['Axial length of the right eye is 0.40 mm longer than axial length of the left eye.']);
    expect([report.right?.target, report.left?.target]).toEqual([-0.5, -0.25]);
    expect([report.right?.axialLength, report.right?.k, report.left?.axialLength, report.left?.k]).toEqual([23.5, 44.5, 23.2, 44.6]);
  });

  it('reads each lens’s table, keeping the columns and eyes apart', () => {
    expect(report.right!.tables.map((t) => [t.toric, t.rows.map((r) => [r.power, r.ref])])).toEqual([
      [false, [[21.5, -1.03], [21, -0.68], [20.5, -0.33], [20, 0.02]]],
      [false, [[21, -1.18], [20.5, -0.81], [20, -0.45], [19.5, -0.09]]],
      [true, [[21, -0.68], [20.5, -0.33]]],
    ]);
    expect(report.right!.tables[0].title).toContain('DCBOO');
    expect(report.left!.tables[0].rows.map((r) => r.power)).toEqual([23, 22.5, 22, 21.5]);
    expect(report.left!.tables[2].rows).toMatchObject([
      { model: '150', power: 22.5, ref: -0.87, cyl: 1.5, axis: 167 },
      { model: '150', power: 22, ref: -0.51, cyl: 1.5, axis: 167 },
    ]);
  });

  it('copes with what OCR does to numbers', () => {
    expect(['+20.50', '-0.33', '+2050', '-046', '421.00', '+0,25', '−1.03', '3°', 'abc'].map(readNumber)).toEqual([20.5, -0.33, 20.5, -0.46, 21, 0.25, -1.03, null, null]);
  });

  it('still splits the eyes when OCR misreads or misses the OD / OS headings', () => {
    const misread = readReport(REPORT.map((w) => (w.text === 'OS' ? { ...w, text: 'O5.' } : w)))!;
    expect(misread.left?.tables[0].rows[0].power).toBe(23);
    const missing = readReport(REPORT.filter((w) => w.text !== 'OD' && w.text !== 'OS'))!;
    expect(missing.right?.tables[0].rows.map((r) => r.power)).toEqual([21.5, 21, 20.5, 20]);
    expect(missing.left?.tables[0].rows.map((r) => r.power)).toEqual([23, 22.5, 22, 21.5]);
    expect(missing.right?.target).toBe(-0.5);
  });

  it('reads a report with lens tables for one eye only (the other already operated on)', () => {
    // The left eye's tables gone (its target stays), with and without the OD / OS headings.
    const oneEye = REPORT.filter((w) => w.y0 < 7 * 30 || w.x0 < 860);
    for (const words of [oneEye, oneEye.filter((w) => w.text !== 'OD' && w.text !== 'OS')]) {
      const r = readReport(words)!;
      expect(r.right?.tables.map((t) => t.rows.length)).toEqual([4, 4, 2]);
      expect(r.left?.tables ?? []).toEqual([]);
      expect(checkLens(r, lens({})).facts).toEqual(['Right eye: DCB00 +20.50 → predicted −0.33']);
    }
  });

  it('isn’t fooled by a page that isn’t an IOL calculation', () => {
    expect(readReport(page([[[0, 'Biometry'], [20, 'AL: 23.53 mm']]]))).toBeNull();
  });
});

describe('checking the lens you picked against the report', () => {
  const report = readReport(REPORT)!;

  it('finds the lens for the eye being operated on, with its predicted refraction', () => {
    expect(checkLens(report, lens({}))).toMatchObject({ level: 'ok', facts: ['Right eye: DCB00 +20.50 → predicted −0.33'], flags: [] });
    expect(checkLens(report, lens({ code: 'AR40', power: 20 })).facts).toEqual(['Right eye: AR40 +20.00 → predicted −0.45']);
  });

  it('flags a power from the other eye’s column', () => {
    const check = checkLens(report, lens({ power: 22 }));
    expect(check.level).toBe('error');
    expect(check.flags[0].text).toBe('+22.00 is in the LEFT eye’s column of the report, not the right’s.');
  });

  it('gives the axis for a toric, matched by its number', () => {
    const toric = lens({ eye: 'LEFT', code: 'ZCU150', family: { base: 'ZCU', number: '150' }, power: 22, target: -0.25 });
    expect(checkLens(report, toric)).toMatchObject({ level: 'ok', facts: ['Left eye: ZCU150 +22.00, axis 167° → predicted −0.51'] });
    // A ZCU100 isn't what the report calculated for the left eye.
    expect(checkLens(report, { ...toric, code: 'ZCU100', family: { base: 'ZCU', number: '100' } }).level).toBe('warn');
  });

  it('warns about a different target, a power off the table or a lens not on the report', () => {
    // Plano and −0.50 are the usual targets (the IOLMaster is set to −0.50): no note for those.
    expect(checkLens(report, lens({ target: 0 })).flags).toEqual([]);
    expect(checkLens(report, lens({ target: -0.5 })).flags).toEqual([]);
    expect(checkLens(report, lens({ target: -1 })).flags.map((f) => f.text)).toEqual(['The report was calculated for a target of −0.50 D (right eye); you picked −1.00.']);
    expect(checkLens(report, lens({ power: 24 })).flags[0].text).toBe('DCB00 +24.00 isn’t among the right eye’s powers on the report (+20.00 to +21.50).');
    expect(checkLens(report, lens({ code: 'CNA0T4', family: { base: 'CNA0T', number: '4' } })).flags[0].text).toBe('CNA0T4 isn’t on the right eye’s side of the report (add its A-constant in Settings to estimate it).');
  });

  it('flags a report with another patient’s name, first and as an error; a different ID is only a note', () => {
    const wrong = checkLens(report, lens({ mrn: '2000002', surname: 'JONES', givenName: 'Kim' }));
    expect(wrong.level).toBe('error');
    expect(wrong.nameMismatch).toBe('NAME MISMATCH: the report is for “SMITH, Jan”, not “JONES, Kim”. Check you have this patient’s report.');
    expect(wrong.flags[0].text).toBe(wrong.nameMismatch);
    expect(wrong.idNote).toMatch(/^The report’s patient ID \(1000001\) isn’t the MRN on the list \(2000002\)/);
    // Only the ID differs (another system's number): the lens check still passes.
    expect(checkLens(report, lens({ mrn: '2000002', givenName: 'Jan' })).level).toBe('ok');
  });

  it('compares the given name too, as family members share a surname', () => {
    expect(checkLens(report, lens({ givenName: 'Ron' })).nameMismatch).toBeDefined();
    expect(checkLens(report, lens({ givenName: 'J' })).nameMismatch).toBeUndefined();
    // A letter misread by OCR in a long name is allowed, and part of a double-barrelled name matches.
    expect(checkLens({ ...report, name: 'SMlTHERS, Jan' }, lens({ surname: 'SMITHERS', givenName: 'Jan' })).level).toBe('ok');
    expect(checkLens({ ...report, name: 'SMITH-JONES, Jan' }, lens({ surname: 'SMITH', givenName: 'Jan' })).level).toBe('ok');
    // A short name must match exactly.
    expect(checkLens({ ...report, name: 'OH, Ann' }, lens({ surname: 'NG', givenName: 'Ann' })).nameMismatch).toBeDefined();
  });

  it('says so when the name on the report can’t be read', () => {
    expect(checkLens({ ...report, name: '' }, lens({})).flags.map((f) => f.text)).toContain('Couldn’t read the patient’s name on the report: check it’s this patient’s.');
  });

  it('checks both eyes when the eye isn’t set, and reads several screenshots as one', () => {
    const check = checkLens(report, lens({ eye: '', power: 22 }));
    expect(check.facts).toEqual(['Left eye: DCB00 +22.00 → predicted −0.51']);
    expect(check.flags.some((f) => f.text.startsWith('The eye isn’t set'))).toBe(true);
    const combined = combineReadings([{ ...report, left: undefined }, { ...report, right: undefined, warnings: [] }])!;
    expect(checkLens(combined, lens({ eye: 'LEFT', power: 22, target: -0.25 })).level).toBe('ok');
  });
});

describe('SRK/T', () => {
  it('gives back the target refraction for the power it calculates', () => {
    for (const [al, k, a, target] of [[23.5, 44, 118.4, 0], [21, 46, 119.3, -0.5], [26.5, 42, 119.1, -2]]) {
      expect(srktRefraction(al, k, a, srktPower(al, k, a, target))).toBeCloseTo(target, 6);
    }
  });

  it('needs more power for a shorter eye, a flatter cornea or a higher A-constant', () => {
    const p = (al: number, k: number, a: number) => srktPower(al, k, a, 0);
    expect(p(22, 44, 119)).toBeGreaterThan(p(24, 44, 119));
    expect(p(23.5, 42, 119)).toBeGreaterThan(p(23.5, 45, 119));
    expect(p(23.5, 44, 119.3)).toBeGreaterThan(p(23.5, 44, 119.1));
    // An average eye with a typical A-constant: about +20 D for emmetropia (SRK II gives +20.05).
    expect(p(23.5, 44, 118.4)).toBeCloseTo(20.04, 2);
  });
});

describe('estimating a lens that isn’t on the report', () => {
  const report = readReport(REPORT)!;
  const constants = parseConstants(['DCB00 119.3', 'CNA0Tx 119.1', 'not a constant']);
  const cna = lens({ code: 'CNA0T0', family: { base: 'CNA0T', number: '0' }, power: 20.5 });

  it('reads the A-constants from Settings', () => {
    expect(constants).toEqual([{ pattern: 'DCB00', value: 119.3 }, { pattern: 'CNA0Tx', value: 119.1 }]);
  });

  it('works it out from the report’s numbers for a lens with a known A-constant, showing the working', () => {
    const check = checkLens(report, cna, constants);
    const shift = srktPower(23.5, 44.5, 119.1, -0.5) - srktPower(23.5, 44.5, 119.3, -0.5);
    expect(check.estimate!.shift).toBeCloseTo(shift, 2);
    expect(check.estimate!.shift).toBeLessThan(-0.2);
    expect(check.level).toBe('warn');
    expect(check.flags[0].text).toBe('CNA0T0 isn’t on the report: estimated from DCB00. Check the working.');
    const working = check.estimate!.working.join('\n');
    expect(working).toContain('A-constants used (SRK/T, from Settings): CNA0T0 119.1 (as CNA0Tx), DCB00 119.3.');
    expect(working).toContain('CNA0T0 has a lower A-constant than DCB00 (119.1 vs 119.3), so it needs a lower power for the same result.');
    expect(working).toMatch(/SRK\/T for this eye \(AL 23\.50 mm, K 44\.50 D, target −0\.50\): DCB00 \(A 119\.3\) \+\d+\.\d\d, CNA0T0 \(A 119\.1\) \+\d+\.\d\d, so CNA0T0 needs 0\.\d\d D lower\./);
    expect(working).toContain('Keeping the same power with CNA0T0 gives a more myopic result');
    // +20.50 of CNA0T0 is like DCB00 +20.75 or so: between −0.33 (+20.50) and −0.68 (+21.00).
    const predicted = -0.33 + ((20.5 - check.estimate!.shift - 20.5) / 0.5) * (-0.68 + 0.33);
    expect(check.facts).toEqual([`Right eye (estimated): CNA0T0 +20.50 → about −${Math.abs(Math.round(predicted * 100) / 100).toFixed(2)}`]);
    expect(check.estimate!.table.map((t) => t.power)).toEqual([21, 20.5, 20]);
  });

  it('says how to estimate a lens it has no A-constant for', () => {
    expect(checkLens(report, lens({ code: 'SN60WF', power: 20 }), constants).flags[0].text).toBe(
      'SN60WF isn’t on the right eye’s side of the report (add its A-constant in Settings to estimate it).',
    );
  });

  it('falls back to the A-constants’ difference when the eye’s AL and K weren’t read', () => {
    const blind = { ...report, right: { ...report.right!, axialLength: undefined, k: undefined } };
    const check = checkLens(blind, cna, constants);
    expect(check.estimate!.shift).toBeCloseTo(-0.2, 5);
    expect(check.estimate!.working.join(' ')).toContain('taken as the A-constants’ difference: 119.1 − 119.3 = −0.20 D');
  });
});

describe('highlighting the screenshot', () => {
  const report = readReport(REPORT)!;
  const boxes = (p: PickedLens) => checkLens(report, p, parseConstants(['DCB00 119.3', 'CNA0Tx 119.1'])).highlights;
  const at = (row: number) => row * 30;

  it('marks the name, MRN, the eye, ΔK, the lens and the power, in the eye’s colour', () => {
    const marks = boxes(lens({}));
    expect(marks.every((m) => m.eye === 'right' && !m.dashed)).toBe(true);
    expect(marks.map((m) => m.y0)).toEqual([at(0), at(1), at(2), at(5), at(9), at(10), at(17)]);
    // ΔK stops before K2.
    expect(marks[4].x1).toBeLessThan(200);
  });

  it('uses blue for the left eye, and dashes the other eye’s column when the power is in it', () => {
    expect(new Set(boxes(lens({ eye: 'LEFT', power: 22, target: -0.25 })).map((m) => m.eye))).toEqual(new Set(['left']));
    const wrong = boxes(lens({ power: 22 }));
    expect(wrong.filter((m) => m.dashed).map((m) => m.eye)).toEqual(['left', 'left']);
  });

  it('dashes the lens an estimate comes from, and the powers either side', () => {
    const marks = boxes(lens({ code: 'CNA0T0', family: { base: 'CNA0T', number: '0' }, power: 20.5 }));
    expect(marks.filter((m) => m.dashed).length).toBe(3);
  });
});

describe('trying another lens in the report view', () => {
  const report = readReport(REPORT)!;
  const constants = parseConstants(['DCB00 119.3', 'CNA0Tx 119.1']);
  const sim = (over: Partial<PickedLens>, side: 'right' | 'left' = 'right') => simulateLens(report, side, lens({ power: null, ...over }), constants);
  const table = (s: Simulation | string) => (s as Simulation).rows.map((r) => [r.power, r.ref, r.outside ?? false]);

  it('lists powers highest first, from the report, with the best for the target', () => {
    const s = sim({}) as Simulation;
    expect(s.source).toBe('report');
    expect(s.best).toMatchObject({ power: 20.5, ref: -0.33 });
    expect(table(s)).toEqual([[21.5, -1.03, false], [21, -0.68, false], [20.5, -0.33, false], [20, 0.02, false], [19.5, 0.37, true]]);
    expect(s.notes).toContain('Powers marked ≈ are beyond those on the report, read along its slope: less certain.');
  });

  it('finds the best power for another target, and keeps a chosen power in the list', () => {
    expect((sim({ target: 0 }) as Simulation).best.power).toBe(20);
    const s = sim({ target: 0, power: 22.5 }) as Simulation;
    expect(s.selected).toMatchObject({ power: 22.5, outside: true });
    expect(s.rows[0].power).toBe(22.5);
  });

  it('gives a toric’s cylinder, axis and residual cylinder from the report’s toric calculation', () => {
    const s = sim({ eye: 'LEFT', code: 'ZCU150', family: { base: 'ZCU', number: '150' }, target: -0.25 }, 'left') as Simulation;
    expect(s.source).toBe('report');
    expect(s.rows.find((r) => r.power === 22)).toMatchObject({ ref: -0.51, cyl: 1.5, axis: 167, refCyl: 0.25 });
    expect(s.notes[0]).toBe('From the report’s toric calculation (Barrett Toric) for ZCU150: axis 167°.');
  });

  it('says when the report chose another toric, and leaves out the axis', () => {
    const s = sim({ eye: 'LEFT', code: 'ZCU100', family: { base: 'ZCU', number: '100' } }, 'left') as Simulation;
    expect(s.notes[0]).toMatch(/^The report’s toric calculation \(Barrett Toric\) chose ZCU150 for this eye at 167°\./);
    expect(s.rows.every((r) => r.axis === undefined)).toBe(true);
  });

  it('estimates a lens that isn’t on the report, with the working', () => {
    const s = sim({ code: 'CNA0T0', family: { base: 'CNA0T', number: '0' } }) as Simulation;
    expect(s.source).toBe('estimate');
    expect(s.reference).toBe('DCB00');
    expect(s.working[1]).toBe('A-constants used (SRK/T, from Settings): CNA0T0 119.1 (as CNA0Tx), DCB00 119.3.');
    expect(s.best.power).toBe(20.5);
    // A Clareon toric is the spherical equivalent only.
    expect((sim({ code: 'CNA0T4', family: { base: 'CNA0T', number: '4' } }) as Simulation).notes).toContain('Spherical equivalent only: choose the cylinder and axis with a toric calculator.');
  });

  it('says why it can’t work a lens out', () => {
    expect(sim({ code: 'DIU150', family: { base: 'DIU', number: '150' } })).toBe('DIU150 isn’t on the right eye’s side of the report. Add its A-constant in Settings to estimate it from a lens that is.');
    expect(sim({ code: 'ZCU', family: { base: 'ZCU', number: '' } })).toBe('Choose the number for ZCU… (e.g. the toric model).');
  });
});

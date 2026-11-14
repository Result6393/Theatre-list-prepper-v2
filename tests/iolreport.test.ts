import { describe, expect, it } from 'vitest';
import { checkLens, combineReadings, readNumber, readReport, type OcrWord, type PickedLens } from '../src/iolreport';

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

const lens = (over: Partial<PickedLens>): PickedLens => ({ eye: 'RIGHT', code: 'DCB00', power: 20.5, target: -0.5, mrn: '1000001', surname: 'SMITH', ...over });

describe('reading an IOL calculation report', () => {
  const report = readReport(REPORT)!;

  it('reads the patient, the report’s warning and each eye’s target', () => {
    expect(report).toMatchObject({ name: 'SMITH, Jan', id: '1000001' });
    expect(report.warnings).toEqual(['Axial length of the right eye is 0.40 mm longer than axial length of the left eye.']);
    expect([report.right?.target, report.left?.target]).toEqual([-0.5, -0.25]);
  });

  it('reads each lens’s table, keeping the columns and eyes apart', () => {
    expect(report.right!.tables.map((t) => [t.toric, t.rows.map((r) => [r.power, r.ref])])).toEqual([
      [false, [[21.5, -1.03], [21, -0.68], [20.5, -0.33], [20, 0.02]]],
      [false, [[21, -1.18], [20.5, -0.81], [20, -0.45], [19.5, -0.09]]],
      [true, [[21, -0.68], [20.5, -0.33]]],
    ]);
    expect(report.right!.tables[0].title).toContain('DCBOO');
    expect(report.left!.tables[0].rows.map((r) => r.power)).toEqual([23, 22.5, 22, 21.5]);
    expect(report.left!.tables[2].rows).toEqual([
      { model: '150', power: 22.5, ref: -0.87, cyl: 1.5, axis: 167 },
      { model: '150', power: 22, ref: -0.51, cyl: 1.5, axis: 167 },
    ]);
  });

  it('copes with what OCR does to numbers', () => {
    expect(['+20.50', '-0.33', '+2050', '-046', '421.00', '+0,25', '−1.03', '3°', 'abc'].map(readNumber)).toEqual([20.5, -0.33, 20.5, -0.46, 21, 0.25, -1.03, null, null]);
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
    expect(checkLens(report, lens({ target: 0 })).flags.map((f) => f.text)).toEqual(['The report was calculated for a target of −0.50 D (right eye); you picked plano.']);
    expect(checkLens(report, lens({ power: 24 })).flags[0].text).toBe('DCB00 +24.00 isn’t among the right eye’s powers on the report (+20.00 to +21.50).');
    expect(checkLens(report, lens({ code: 'CNA0T4', family: { base: 'CNA0T', number: '4' } })).flags[0].text).toBe('CNA0T4 isn’t on the right eye’s side of the report.');
  });

  it('flags a report for someone else', () => {
    expect(checkLens(report, lens({ mrn: '2000002', surname: 'JONES' })).flags.map((f) => f.text)).toEqual([
      'The report is for patient ID 1000001, not 2000002.',
      'The report’s name (SMITH, Jan) doesn’t match JONES.',
    ]);
    // A letter misread by OCR in a long name is allowed.
    expect(checkLens({ ...report, name: 'SMlTHERS, Jan' }, lens({ surname: 'SMITHERS' })).level).toBe('ok');
  });

  it('checks both eyes when the eye isn’t set, and reads several screenshots as one', () => {
    const check = checkLens(report, lens({ eye: '', power: 22 }));
    expect(check.facts).toEqual(['Left eye: DCB00 +22.00 → predicted −0.51']);
    expect(check.flags.some((f) => f.text.startsWith('The eye isn’t set'))).toBe(true);
    const combined = combineReadings([{ ...report, left: undefined }, { ...report, right: undefined, warnings: [] }])!;
    expect(checkLens(combined, lens({ eye: 'LEFT', power: 22, target: -0.25 })).level).toBe('ok');
  });
});

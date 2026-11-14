import { describe, expect, it } from 'vitest';
import { defaultLayout } from '../src/layout';
import { absentText, addToken, cycleFinding, defaultPrompts, findingGroups, findingState, formatPower, formatTarget, hasToken, promptColumns, readEyeOrder, readGrading, readLens, readTarget, removeToken, toggleFinding, writeEyeOrder, writeGrading, writeLens, writeTarget } from '../src/prompts';
import { defaultSettings, importSettings } from '../src/settings';

const { findings, lenses } = defaultPrompts();
const groups = findingGroups(findings);
const group = (token: string) => groups.find((g) => g.options.some((o) => o.token === token))!;
const toggle = (cell: string, token: string) => toggleFinding(cell, group(token), token, findings);

describe('cataract prompt findings', () => {
  it('turns pick-one choices into a labelled group', () => {
    expect(groups.map((g) => [g.label, g.options.map((o) => o.short)])).toEqual([
      ['', ['PXF']],
      ['', ['Guttata']],
      ['Dilation', ['Good', 'Mod', 'Poor']],
      ['', ['Phacodonesis']],
      ['', ['Tamsulosin']],
    ]);
  });

  it('adds findings on one line under what is already there, and takes them out again', () => {
    let cell = 'Cantonese interpreter\nDiabetic';
    cell = toggle(cell, 'PXF');
    expect(cell).toBe('Cantonese interpreter\nDiabetic\nPXF');
    cell = toggle(cell, 'Guttata');
    cell = toggle(cell, 'Tamsulosin');
    expect(cell).toBe('Cantonese interpreter\nDiabetic\nPXF, Guttata, Tamsulosin');
    cell = toggle(cell, 'Guttata');
    expect(cell).toBe('Cantonese interpreter\nDiabetic\nPXF, Tamsulosin');
    cell = toggle(toggle(cell, 'PXF'), 'Tamsulosin');
    expect(cell).toBe('Cantonese interpreter\nDiabetic');
  });

  it('keeps only one choice from a pick-one group', () => {
    let cell = toggle('', 'Good dilation');
    expect(cell).toBe('Good dilation');
    cell = toggle(cell, 'Poor dilation');
    expect(cell).toBe('Poor dilation');
    expect(toggle(cell, 'Poor dilation')).toBe('');
  });

  it('recognises findings you typed yourself, as whole words only', () => {
    expect(hasToken('pxf, tamsulosin', 'PXF')).toBe(true);
    expect(hasToken('PXFX', 'PXF')).toBe(false);
    expect(removeToken('Known PXF; on drops', 'PXF')).toBe('Known; on drops');
    expect(addToken('', 'PXF', findings)).toBe('PXF');
    expect(addToken('PXF', 'pxf', findings)).toBe('PXF');
  });
});

describe('lens and target', () => {
  it('writes the lens on the first line and swaps it when another is picked', () => {
    let cell = 'Order backup lens';
    cell = writeLens(cell, lenses, { lens: 'DCB00', extra: '', power: '21.5' });
    expect(cell).toBe('DCB00 +21.5D\nOrder backup lens');
    expect(readLens(cell, lenses)).toEqual({ lens: 'DCB00', extra: '', power: '+21.5D' });
    cell = writeLens(cell, lenses, { lens: 'CNA0Tx', extra: '4', power: '+21.5D' });
    expect(cell).toBe('CNA0T4 +21.5D\nOrder backup lens');
    expect(readLens(cell, lenses)).toEqual({ lens: 'CNA0Tx', extra: '4', power: '+21.5D' });
    expect(writeLens(cell, lenses, null)).toBe('Order backup lens');
  });

  it('reads a lens typed by hand, and a toric number on ZCU', () => {
    expect(readLens('zcu225 +20D', lenses)).toEqual({ lens: 'ZCUx', extra: '225', power: '+20D' });
    expect(readLens('DIB00', lenses)).toEqual({ lens: 'DIB00', extra: '', power: '' });
    expect(readLens('Discuss with boss', lenses)).toBeNull();
  });

  it('puts the target under the lens, with plano one click away', () => {
    let cell = writeLens('', lenses, { lens: 'DCB00', extra: '', power: '' });
    cell = writeTarget(cell, lenses, 'plano');
    expect(cell).toBe('DCB00\nTarget plano');
    expect(readTarget(cell)).toBe('plano');
    cell = writeTarget(cell, lenses, '-0.5');
    expect(cell).toBe('DCB00\nTarget -0.50');
    expect(writeTarget(cell, lenses, null)).toBe('DCB00');
    expect(writeTarget('', lenses, 'plano')).toBe('Target plano');
  });

  it('tidies powers and targets but never guesses a sign', () => {
    expect(formatPower('21')).toBe('+21.0D');
    expect(formatPower('-1.5d')).toBe('-1.5D');
    expect(formatPower('TBC')).toBe('TBC');
    expect(formatTarget('0')).toBe('plano');
    expect(formatTarget('PL')).toBe('plano');
    expect(formatTarget('+0.25')).toBe('+0.25');
    expect(formatTarget('0.5')).toBe('0.5');
    expect(formatTarget('-2.5 (monovision)')).toBe('-2.5 (monovision)');
  });
});

describe('prompt settings', () => {
  it('finds the default cheat-sheet columns automatically', () => {
    expect(promptColumns(defaultLayout(), defaultPrompts())).toEqual({ findings: 3, lens: 6 });
    expect(promptColumns(defaultLayout(), { ...defaultPrompts(), lensColumn: 'Plan & notes' }).lens).toBe(9);
    expect(promptColumns(defaultLayout(), { ...defaultPrompts(), lensColumn: 'Gone' }).lens).toBe(-1);
  });

  it('fills in defaults for settings files from before prompts existed', () => {
    const old = JSON.parse(JSON.stringify(defaultSettings()));
    delete old.prompts;
    expect(importSettings(JSON.stringify(old)).prompts).toEqual(defaultPrompts());
    const custom = importSettings('{"prompts":{"enabled":false,"lenses":["SN60WF"]}}').prompts;
    expect(custom).toEqual({ ...defaultPrompts(), enabled: false, lenses: ['SN60WF'] });
  });
});

describe('first or second eye', () => {
  it('reads 1st or 2nd eye from anywhere in the cell, in either spelling', () => {
    expect(readEyeOrder('NS2. 1st eye. Macula ok.')).toBe('first');
    expect(readEyeOrder('Cantonese interpreter\n2nd eye')).toBe('second');
    expect(readEyeOrder('First eye')).toBe('first');
    expect(readEyeOrder('second EYE, PXF')).toBe('second');
    expect(readEyeOrder('NS2, macula ok')).toBeNull();
    expect(readEyeOrder('')).toBeNull();
  });

  it('puts it on its own line at the top, above what is already there', () => {
    expect(writeEyeOrder('Cantonese interpreter\nDiabetic', 'first')).toBe('1st eye\nCantonese interpreter\nDiabetic');
    expect(writeEyeOrder('', 'second')).toBe('2nd eye');
  });

  it('switches 1st to 2nd without leaving the old mention behind', () => {
    expect(writeEyeOrder('1st eye\nDiabetic', 'second')).toBe('2nd eye\nDiabetic');
    expect(writeEyeOrder('NS2. 1st eye. Macula ok.', 'second')).toBe('2nd eye\nNS2. Macula ok.');
  });

  it('removes it when you pick it again, keeping everything else', () => {
    expect(writeEyeOrder('1st eye\nDiabetic\nPXF', null)).toBe('Diabetic\nPXF');
    expect(writeEyeOrder('NS2, 2nd eye, PXF', null)).toBe('NS2, PXF');
    expect(writeEyeOrder('NS2, 2nd eye', null)).toBe('NS2');
    expect(writeEyeOrder('2nd eye', null)).toBe('');
    expect(writeEyeOrder('', null)).toBe('');
  });

  it('leaves findings and lens text alone', () => {
    const cell = writeEyeOrder(toggle('Diabetic', 'PXF'), 'first');
    expect(cell).toBe('1st eye\nDiabetic\nPXF');
    expect(toggle(cell, 'Guttata')).toBe('1st eye\nDiabetic\nPXF, Guttata');
    expect(writeEyeOrder(cell, 'second')).toBe('2nd eye\nDiabetic\nPXF');
  });
});

describe('findings with three states', () => {
  const { findings } = defaultPrompts();

  it('goes present, then not present, then back to nothing', () => {
    const present = cycleFinding('', 'PXF', findings);
    expect(present).toBe('PXF');
    const absent = cycleFinding(present, 'PXF', findings);
    expect(absent).toBe('No PXF');
    expect(findingState(absent, 'PXF')).toBe('absent');
    expect(cycleFinding(absent, 'PXF', findings)).toBe('');
  });

  it('keeps the other findings on the same line', () => {
    let cell = 'NS2\nPXF, Poor dilation';
    cell = cycleFinding(cell, 'Guttata', findings);
    expect(cell).toBe('NS2\nPXF, Poor dilation, Guttata');
    cell = cycleFinding(cell, 'Guttata', findings);
    expect(cell).toBe('NS2\nPXF, Poor dilation, No guttata');
    expect([findingState(cell, 'PXF'), findingState(cell, 'Guttata'), findingState(cell, 'Tamsulosin')]).toEqual(['present', 'absent', null]);
    expect(cycleFinding(cell, 'Guttata', findings)).toBe('NS2\nPXF, Poor dilation');
  });

  it('reads “no” and “nil” typed by hand', () => {
    expect(findingState('nil PXF', 'PXF')).toBe('absent');
    expect(findingState('no tamsulosin', 'Tamsulosin')).toBe('absent');
    expect(findingState('PXF OU', 'PXF')).toBe('present');
    expect([absentText('PXF'), absentText('Guttata'), absentText('Tamsulosin')]).toEqual(['No PXF', 'No guttata', 'No tamsulosin']);
  });
});

describe('cataract grade at the top of the diagnosis box', () => {
  it('goes after 1st / 2nd eye on the top line, before the findings', () => {
    let cell = writeEyeOrder('PXF, Poor dilation\nPOAG', 'first');
    cell = writeGrading(cell, 'NS2+ C1');
    expect(cell).toBe('1st eye: NS2+ C1\nPXF, Poor dilation\nPOAG');
    expect([readGrading(cell), readEyeOrder(cell)]).toEqual(['NS2+ C1', 'first']);
    expect(writeEyeOrder(cell, 'second')).toBe('2nd eye: NS2+ C1\nPXF, Poor dilation\nPOAG');
  });

  it('stands on its own line until the eye is picked, and keeps its place when it is', () => {
    const cell = writeGrading('PXF', 'NS3');
    expect(cell).toBe('Cataract: NS3\nPXF');
    expect(readGrading(cell)).toBe('NS3');
    expect(writeEyeOrder(cell, 'second')).toBe('2nd eye: NS3\nPXF');
    expect(writeEyeOrder('2nd eye: NS3\nPXF', null)).toBe('Cataract: NS3\nPXF');
  });

  it('is cleared, and read from a box typed by hand', () => {
    expect(writeGrading('1st eye: NS2\nPXF', '')).toBe('1st eye\nPXF');
    expect(readGrading('2nd eye. NS2 C1\nGuttata')).toBe('NS2 C1');
    expect(readGrading('PXF\n1st eye')).toBe('');
  });
});

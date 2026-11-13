import { describe, expect, it } from 'vitest';
import { allergyColumn, allergyPending, lacksAllergyDetails, pendingAllergies } from '../src/allergy';
import { extractSessions } from '../src/extract';
import { defaultLayout } from '../src/layout';
import { freshRows } from '../src/merge';
import { parseTable } from '../src/parse';
import { booking, runSheetCsv, slot } from './fixtures';

const layout = defaultLayout();
const col = allergyColumn(layout);
const or1 = slot('TEH OR 01', 'Open Unallocated', '8:00', '12:00');
const person = (mrn: string, allergy: string) =>
  booking(or1, { CONSULTANT: 'Cedar, M', MRN: mrn, PT_NAME: `TESTER${mrn}, Pat`, SCHED_TM: '08:00', ALLERGY: allergy });
const rowsFor = (...allergies: string[]) => {
  const table = parseTable(runSheetCsv([or1, ...allergies.map((a, i) => person(`500000${i}`, a))]));
  const [session] = extractSessions(table, { flagHospitals: [], hospitalCodes: [] });
  return freshRows(session.patients, layout);
};

describe('allergies needing details', () => {
  it('finds the allergy column in the layout', () => {
    expect(col).toBe(5);
    expect(allergyColumn({ ...layout, columns: [{ heading: 'X', source: 'blank' }] })).toBe(-1);
  });

  it('marks only patients the run sheet says have an allergy', () => {
    const rows = rowsFor('Y', 'NKA', 'Not Recorded', 'Penicillin');
    expect(rows.map((r) => r.patient!.allergyPending)).toEqual([true, false, false, false]);
    expect(rows.map((r) => allergyPending(r, col))).toEqual([true, false, false, false]);
  });

  it('clears once you type the details, and comes back if you delete them', () => {
    const [row] = rowsFor('Y');
    expect(allergyPending(row, col)).toBe(true);
    row.cells[col] = 'Penicillin: rash';
    expect(allergyPending(row, col)).toBe(false);
    row.cells[col] = '';
    expect(allergyPending(row, col)).toBe(true);
  });

  it('clears when acknowledged, without details', () => {
    const [row] = rowsFor('Y');
    row.allergyAck = true;
    expect(allergyPending(row, col)).toBe(false);
  });

  it('counts only ticked patients, and ignores cancelled ones', () => {
    const rows = rowsFor('Y', 'Y', 'Y');
    rows[1].included = false;
    rows[2].status = 'cancelled';
    expect(pendingAllergies(rows, col)).toEqual([rows[0]]);
  });

  it('does nothing when there is no allergy column', () => {
    const [row] = rowsFor('Y');
    expect(allergyPending(row, -1)).toBe(false);
  });

  it('treats a bare "Allergy: Y" as no details, but real text as details', () => {
    expect(['', 'Allergy: Y', 'allergy y', 'Allergy: Yes', 'Y', 'yes.', ' Allergy: '].map(lacksAllergyDetails)).toEqual([true, true, true, true, true, true, true]);
    expect(['Allergy: Yeast', 'Penicillin', 'Allergy not recorded', 'NKA', 'Latex'].map(lacksAllergyDetails)).toEqual([false, false, false, false, false]);
  });
});

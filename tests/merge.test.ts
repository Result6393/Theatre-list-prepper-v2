import { describe, expect, it } from 'vitest';
import { toTsv } from '../src/clipboard';
import { extractSessions, type Session } from '../src/extract';
import { buildRow, defaultLayout, type Layout } from '../src/layout';
import { ageOn, backfillFromBanner, backfillProblem, backfillRow, CANCELLED_LABEL, copyBlock, copyPictures, freshRows, inSections, isRectangular, keepReports, mergePrepped, MergeError, normaliseMrn, parsePastedRows, rowMrn, sheetRows, type GridRow } from '../src/merge';
import type { ReportImage } from '../src/report';
import { parseTable } from '../src/parse';
import { booking, runSheetCsv, slot } from './fixtures';

// Invented patients on an invented list, the night before and on the morning.
const OPTIONS = { flagHospitals: ['POWH', 'STG'], hospitalCodes: ['TEH'] };
const or4 = slot('TEH OR 04', 'Willow, R Eye - General', '8:00', '12:00');
const or1 = slot('TEH OR 01', 'Open Unallocated', '8:00', '12:00');
const cataract = (side: string) => `${side} cataract extraction and insertion of intraocular lens implant`;
const person = (mrn: string, name: string, time: string, side = 'right', s = or4) =>
  booking(s, { CONSULTANT: 'Willow, R', MRN: mrn, PT_NAME: name, SCHED_TM: time, PLANNED_PROCEDURE: cataract(side) });

const nightBefore = [
  or4,
  person('3000001', 'ALPHA, Ann', '08:00'),
  person('3000002', 'BRAVO, Ben', '08:30'),
  person('3000003', 'CHARLIE, Cat', '09:00'),
  person('3000004', 'DELTA, Don', '09:30'),
  person('3000005', 'ECHO, Eva', '10:00'),
];

// Bravo and Delta dropped off (Delta moved to OR 1), Charlie is now first,
// Echo's eye changed, and Foxtrot was added.
const morning = [
  or4,
  person('3000003', 'CHARLIE, Cat', '08:00'),
  person('3000001', 'ALPHA, Ann', '08:30'),
  person('3000005', 'ECHO, Eva', '10:00', 'left'),
  person('3000006', 'FOXTROT, Fay', '10:30'),
  or1,
  person('3000004', 'DELTA, Don', '08:00', 'right', or1),
];

const sessionsOf = (rows: Record<string, string>[]) => extractSessions(parseTable(runSheetCsv(rows)), OPTIONS);
const listOf = (sessions: Session[]) => sessions.find((s) => s.or === 'OR 4')!;
const describe_ = (s: Session) => `${s.or} · ${s.ampm}`;

/** Your sheet the night before: the app's rows plus your own notes in Exam and history. */
function preppedSheet(layout: Layout = defaultLayout()): string[][] {
  return listOf(sessionsOf(nightBefore)).patients.map((p) => {
    const row = buildRow(p, layout);
    row[3] = `${row[3]}\nPOAG on 2 drops`.trim();
    row[4] = `VA 6/12 IOP ${p.mrn.slice(-1)}`;
    return row;
  });
}

function merge(pasted: string[][], rows = morning, layout = defaultLayout()) {
  const sessions = sessionsOf(rows);
  return mergePrepped(pasted, listOf(sessions), sessions, layout, describe_);
}

const mrnOf = (cells: string[]) => cells[1].split(' ')[0];

describe('update from my sheet', () => {
  it('keeps your notes, follows the new order, adds new patients and keeps cancelled ones at the bottom', () => {
    const result = merge(preppedSheet());
    expect(result.rows.map((r) => [mrnOf(r.cells), r.status])).toEqual([
      ['3000003', 'kept'],
      ['3000001', 'kept'],
      ['3000005', 'kept'],
      ['3000006', 'new'],
      ['3000002', 'cancelled'],
      ['3000004', 'cancelled'],
    ]);
    expect(result.rows[0].cells[4]).toBe('VA 6/12 IOP 3');
    expect(result.rows[0].cells[3]).toBe('POAG on 2 drops');
    expect(result.rows[3].cells[4]).toBe('');
    expect(result.summary).toEqual({ kept: 3, added: 1, cancelled: 2, moved: 2, changed: 1 });
    expect(result.padTo).toBe(5);
  });

  it('says where a cancelled patient has gone if they are on another list', () => {
    const [bravo, delta] = merge(preppedSheet()).rows.filter((r) => r.status === 'cancelled');
    expect(bravo.note).toBeUndefined();
    expect(delta.note).toBe('Now on OR 1 · AM');
  });

  it('flags run-sheet changes without overwriting your cell, but not your own additions', () => {
    const echo = merge(preppedSheet()).rows[2];
    expect(echo.cells[0]).toBe('RIGHT');
    expect(echo.changes).toEqual({ 0: 'LEFT', 2: 'LE CE/IOL' });
    expect(merge(preppedSheet()).rows[0].changes).toEqual({});
  });

  it('marks only the patients that moved out of order', () => {
    const moved = merge(preppedSheet()).rows.filter((r) => r.movedFrom !== undefined);
    // Charlie jumped ahead of Alpha: both changed places relative to each other.
    expect(moved.map((r) => [mrnOf(r.cells), r.movedFrom])).toEqual([['3000003', 3], ['3000001', 1]]);
  });

  it('marks both patients when two swap places', () => {
    const swapped = [or4, nightBefore[1], { ...nightBefore[3], SCHED_TM: '08:30' }, { ...nightBefore[2], SCHED_TM: '09:00' }, nightBefore[4], nightBefore[5]];
    const result = merge(preppedSheet(), swapped);
    expect(result.rows.map((r) => [mrnOf(r.cells), r.movedFrom ?? null])).toEqual([
      ['3000001', null],
      ['3000003', 3],
      ['3000002', 2],
      ['3000004', null],
      ['3000005', null],
    ]);
    expect(result.summary.moved).toBe(2);
  });

  it('marks just the one patient pushed to the end, not everyone who shifted up', () => {
    const reordered = [or4, ...nightBefore.slice(2).map((r, i) => ({ ...r, SCHED_TM: `0${8 + i}:00` })), { ...nightBefore[1], SCHED_TM: '12:00' }];
    const result = merge(preppedSheet(), reordered);
    expect(result.rows.map((r) => r.movedFrom ?? null)).toEqual([null, null, null, null, 1]);
    expect(result.summary).toMatchObject({ kept: 5, added: 0, cancelled: 0, moved: 1, changed: 0 });
  });

  it('copies the list, then the cancelled patients under a label, then blanks over leftover rows', () => {
    const layout = defaultLayout();
    const result = merge(preppedSheet());
    const block = copyBlock(result.rows, layout, 9);
    expect(block.map((r) => r[1].split(' ')[0])).toEqual([
      '3000003', '3000001', '3000005', '3000006', '', 'Not', '3000002', '3000004', '',
    ]);
    expect(block[5][0]).toBe(CANCELLED_LABEL);
    expect(block.every((r) => r.length === layout.columns.length)).toBe(true);

    result.rows[4].included = false;
    expect(copyBlock(result.rows, layout).length).toBe(7);
    result.rows[5].included = false;
    expect(copyBlock(result.rows, layout)).toHaveLength(4);
  });

  it('copies exactly as before when nothing was merged', () => {
    const rows = freshRows(listOf(sessionsOf(nightBefore)).patients, defaultLayout());
    rows[1].included = false;
    expect(copyBlock(rows, defaultLayout()).map(mrnOf)).toEqual(['3000001', '3000003', '3000004', '3000005']);
  });

  it('round-trips through Excel again, bringing back a patient who is back on the list', () => {
    const layout = defaultLayout();
    const first = merge(preppedSheet());
    const pasted = parsePastedRows(toTsv(copyBlock(first.rows, layout, first.padTo)) + '\r\n');
    const later = [...morning, person('3000002', 'BRAVO, Ben', '11:00')];
    const second = merge(pasted, later);
    expect(second.rows.map((r) => [mrnOf(r.cells), r.status, r.reinstated ?? false])).toEqual([
      ['3000003', 'kept', false],
      ['3000001', 'kept', false],
      ['3000005', 'kept', false],
      ['3000006', 'kept', false],
      ['3000002', 'kept', true],
      ['3000004', 'cancelled', false],
    ]);
    expect(second.rows[4].cells[4]).toBe('VA 6/12 IOP 2');
    expect(second.summary.moved).toBe(0);
    expect(second.padTo).toBe(8);
  });

  it('keeps multi-line cells intact through copy and paste', () => {
    const rows = parsePastedRows(toTsv(preppedSheet()));
    expect(isRectangular(rows)).toBe(true);
    expect(rows[0][3]).toBe('POAG on 2 drops');
    expect(rows).toEqual(preppedSheet());
  });

  it('copes with an extra column copied on the left and a heading row', () => {
    const layout = { ...defaultLayout(), includeHeadings: true };
    const pasted = [layout.columns.map((c) => c.heading), ...preppedSheet()].map((r, i) => [String(i), ...r]);
    pasted[0][0] = '#';
    const result = merge(pasted, morning, layout);
    expect(result.summary.kept).toBe(3);
    expect(result.rows[0].cells[4]).toBe('VA 6/12 IOP 3');
    expect(result.padTo).toBe(5);
  });

  it('matches an MRN-only column even if Excel dropped leading zeros', () => {
    expect(normaliseMrn(' 0012345 ')).toBe('12345');
    expect(normaliseMrn('A0012')).toBe('A0012');
    const layout: Layout = {
      ...defaultLayout(),
      columns: [
        { heading: 'MRN', source: 'field:mrn' },
        { heading: 'Notes', source: 'blank' },
      ],
    };
    const rows = [...morning];
    rows[1] = { ...rows[1], MRN: '0003000003' };
    const result = merge([['3000003', 'my note']], rows, layout);
    expect(result.rows[0].status).toBe('kept');
    expect(result.rows[0].cells).toEqual(['3000003', 'my note']);
  });

  it('refuses rows from a different list', () => {
    const sessions = sessionsOf(morning);
    const or1List = sessions.find((s) => s.or === 'OR 1')!;
    const pasted = preppedSheet().filter((r) => mrnOf(r) !== '3000004');
    expect(() => mergePrepped(pasted, or1List, sessions, defaultLayout(), describe_)).toThrow(MergeError);
  });

  it('needs a Patient or MRN column to match on', () => {
    const layout: Layout = { ...defaultLayout(), columns: [{ heading: 'Eye', source: 'field:eye' }] };
    expect(() => merge([['RIGHT']], morning, layout)).toThrow(/Patient or MRN column/);
  });

  it('ignores trailing blank rows when reading a paste', () => {
    expect(parsePastedRows('a\tb\r\n\t\r\n\r\n')).toEqual([['a', 'b']]);
  });
});

describe('reading a paste', () => {
  it('splits rows correctly when a quoted multi-line cell comes first', () => {
    expect(parsePastedRows('"POAG\r\nOn drops"\tA\r\nx\tB\r\n')).toEqual([['POAG\nOn drops', 'A'], ['x', 'B']]);
  });

  it('keeps an unquoted line break inside a cell when rows end with \\r\\n', () => {
    expect(parsePastedRows('a\tPOAG\nOn drops\tb\r\nc\td\te\r\n')).toEqual([['a', 'POAG\nOn drops', 'b'], ['c', 'd', 'e']]);
  });

  it('notices when line breaks have broken the rows', () => {
    expect(isRectangular(parsePastedRows('a\tPOAG\nOn drops\tb\nc\td\te\n'))).toBe(false);
  });
});

describe('IOL report screenshots follow their patient', () => {
  const pic = (tag: number): ReportImage => ({ bytes: new Uint8Array([tag]), type: 'image/png', width: 4, height: 3 });
  // One screenshot each for Alpha, Bravo (cancelled in the morning) and Charlie.
  const pictures = [[pic(1)], [pic(2)], [pic(3), pic(33)], [], []];

  it('moves with the patient when the list is reordered, and stays with a cancelled patient', () => {
    const result = mergePrepped(preppedSheet(), listOf(sessionsOf(morning)), sessionsOf(morning), defaultLayout(), describe_, pictures);
    expect(result.rows.map((r) => [mrnOf(r.cells), r.reports?.map((p) => p.bytes[0])])).toEqual([
      ['3000003', [3, 33]],
      ['3000001', [1]],
      ['3000005', undefined],
      ['3000006', undefined],
      ['3000002', [2]],
      ['3000004', undefined],
    ]);
    // Saved on the same rows as the copy: list, blank, CANCELLED label, cancelled patients, blanks.
    const copied = copyPictures(result.rows, result.padTo);
    expect(copied.map((p) => p.map((i) => i.bytes[0]))).toEqual([[3, 33], [1], [], [], [], [], [2], []]);
    expect(copied).toHaveLength(copyBlock(result.rows, defaultLayout(), result.padTo).length);
  });

  it('keeps reports already added when rows pasted from your sheet (with no pictures) replace the list', () => {
    const before = merge(preppedSheet()).rows;
    before[1].reports = [pic(7)];
    const after = merge(preppedSheet()).rows;
    keepReports(after, before, defaultLayout());
    expect(after.find((r) => mrnOf(r.cells) === '3000001')?.reports).toEqual([pic(7)]);
    expect(after.filter((r) => r.reports)).toHaveLength(1);
  });

  it('comes with the rows of a sheet opened without a run sheet', () => {
    const rows = sheetRows([['RIGHT', '3000001 ALPHA, Ann'], ['', ''], ['LEFT', '3000002 BRAVO, Ben']], 's', [[pic(1)], [], [pic(2)]]);
    expect(rows.map((r) => r.reports?.[0].bytes[0])).toEqual([1, 2]);
  });
});

describe('backfill cases and patients you cancel', () => {
  const layout = defaultLayout();
  const details = { mrn: '3000009', name: 'GOLF, Gus', ageSex: '70m', eye: 'LEFT' as const, procedure: 'LE CE/IOL' };
  const prepped = () => {
    const rows = merge(preppedSheet()).rows;
    const golf = backfillRow(details, layout, 'backfill:1');
    golf.cells[4] = 'VA 6/9';
    // You cancel Alpha on the day, though the run sheet still lists them.
    const alpha = rows.find((r) => mrnOf(r.cells) === '3000001')!;
    alpha.statusBefore = alpha.status;
    alpha.status = 'cancelled';
    alpha.cancelledByYou = true;
    return inSections([...rows, golf]);
  };

  it('fills in a backfill case’s columns from what you typed, ready to prep', () => {
    const golf = backfillRow(details, layout, 'backfill:1');
    expect(golf.cells.slice(0, 3)).toEqual(['LEFT', '3000009 GOLF, Gus 70M', 'LE CE/IOL']);
    expect(golf).toMatchObject({ patient: null, status: 'sheet', backfill: true, unlocked: true, included: true });
    expect(rowMrn(golf, layout)).toBe('3000009');
  });

  it('needs an MRN and a name, so the row can be read back from your sheet', () => {
    expect(backfillProblem({ ...details, mrn: '' })).toMatch(/MRN/);
    expect(backfillProblem({ ...details, mrn: '12' })).toMatch(/at least 4 digits/);
    expect(backfillProblem({ ...details, name: ' ' })).toMatch(/name/);
    expect(backfillProblem(details)).toBeNull();
  });

  it('copies the list, then BACKFILL, then CANCELLED and the patients you cancelled under their own row', () => {
    const block = copyBlock(prepped(), layout);
    const firsts = block.map((r) => (r[0] === 'BACKFILL' || r[0] === 'CANCELLED' ? `${r[0]} ${r[1]}` : mrnOf(r) || '-'));
    expect(firsts).toEqual([
      '3000003', '3000005', '3000006',
      '-', 'BACKFILL Added on the day, not on the run sheet', '3000009',
      '-', 'CANCELLED Not on the latest run sheet', '3000002', '3000004',
      '-', 'CANCELLED Cancelled on the day', '3000001',
    ]);
  });

  it('reads the sections back from your sheet', () => {
    const rows = sheetRows(copyBlock(prepped(), layout), 'sheet');
    expect(rows.map((r) => [mrnOf(r.cells), r.status, Boolean(r.backfill), Boolean(r.cancelledByYou)])).toEqual([
      ['3000003', 'sheet', false, false],
      ['3000005', 'sheet', false, false],
      ['3000006', 'sheet', false, false],
      ['3000009', 'sheet', true, false],
      ['3000002', 'cancelled', false, false],
      ['3000004', 'cancelled', false, false],
      ['3000001', 'cancelled', false, true],
    ]);
  });

  it('keeps a backfill case and your cancellations through an update from the run sheet', () => {
    const result = merge(copyBlock(prepped(), layout));
    const by = (mrn: string) => result.rows.find((r) => mrnOf(r.cells) === mrn)!;
    expect(result.rows.map((r) => mrnOf(r.cells))).toEqual(['3000003', '3000005', '3000006', '3000009', '3000002', '3000004', '3000001']);
    expect(by('3000009')).toMatchObject({ status: 'sheet', backfill: true, patient: null });
    expect(by('3000009').cells[4]).toBe('VA 6/9');
    // Still on the run sheet, but you cancelled them: they stay cancelled, not "Back on list".
    expect(by('3000001')).toMatchObject({ status: 'cancelled', cancelledByYou: true, statusBefore: 'kept' });
    expect(by('3000001').reinstated).toBeUndefined();
    expect(by('3000001').patient?.mrn).toBe('3000001');
    expect(result.summary.cancelled).toBe(2);
  });

  it('a backfill case booked on the run sheet since joins the list, with your notes', () => {
    const booked = [...morning.slice(0, 5), person('3000009', 'GOLF, Gus', '11:00', 'left'), ...morning.slice(5)];
    const golf = merge(copyBlock(prepped(), layout), booked).rows.find((r) => mrnOf(r.cells) === '3000009')!;
    expect(golf).toMatchObject({ status: 'kept', fromBackfill: true });
    expect(golf.backfill).toBeUndefined();
    expect(golf.cells[4]).toBe('VA 6/9');
    expect(golf.movedFrom).toBeUndefined();
  });

  it('puts rows in their sections: the list, the backfill cases, the cancelled', () => {
    const row = (id: string, extra: Partial<GridRow>): GridRow => ({ patient: null, id, cells: [], included: true, status: 'sheet', changes: {}, ...extra });
    const rows = [row('c', { status: 'cancelled' }), row('b', { backfill: true }), row('m', {}), row('cb', { status: 'cancelled', backfill: true })];
    expect(inSections(rows).map((r) => r.id)).toEqual(['m', 'b', 'c', 'cb']);
  });
});

describe('a backfill case from a screenshot of the patient banner', () => {
  it('works out the age on the day of the list', () => {
    expect(ageOn('15/06/1950', '14/06/2026')).toBe(75);
    expect(ageOn('15/06/1950', '15/06/2026')).toBe(76);
    expect(ageOn('01/02/1950', '')).toBeNull();
    expect(ageOn('', '14/06/2026')).toBeNull();
  });

  it('fills the MRN, the name as on the run sheet and the age', () => {
    expect(backfillFromBanner({ first: 'Mary Ann', surname: 'Sample-Jones', dob: '01/02/1960', mrn: '0012345' }, '14/10/2026')).toEqual({
      mrn: '0012345',
      name: 'SAMPLE-JONES, Mary Ann',
      ageSex: '66',
    });
  });

  it('leaves out what the banner didn’t give', () => {
    expect(backfillFromBanner({ first: '', surname: '', dob: '01/01/1900', mrn: '1234567' }, 'not a date')).toEqual({ mrn: '1234567' });
  });
});

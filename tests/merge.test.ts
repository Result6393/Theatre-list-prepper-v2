import { describe, expect, it } from 'vitest';
import { toTsv } from '../src/clipboard';
import { extractSessions, type Session } from '../src/extract';
import { buildRow, defaultLayout, type Layout } from '../src/layout';
import { CANCELLED_LABEL, copyBlock, freshRows, isRectangular, mergePrepped, MergeError, normaliseMrn, parsePastedRows } from '../src/merge';
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
    expect(result.summary).toEqual({ kept: 3, added: 1, cancelled: 2, moved: 1, changed: 1 });
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
    expect(moved.map((r) => [mrnOf(r.cells), r.movedFrom])).toEqual([['3000003', 3]]);
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

import { describe, expect, it } from 'vitest';
import { extractSessions } from '../src/extract';
import { combineSessions, listOfPatient, orderedSelection } from '../src/lists';
import { parseTable } from '../src/parse';
import { runSheetCsv } from './fixtures';

const sessions = extractSessions(parseTable(runSheetCsv()), { flagHospitals: [], hospitalCodes: [] }).filter((s) => s.patients.length);
const [am, pm] = sessions.filter((s) => s.or === 'OR 1');

describe('lists that run on together', () => {
  it('combines lists into one, in order, naming each part once', () => {
    const both = combineSessions([am, pm]);
    expect(both.patients.map((p) => p.mrn)).toEqual([...am.patients, ...pm.patients].map((p) => p.mrn));
    expect(both.or).toBe('OR 1');
    expect(both.ampm).toBe('AM + PM');
    expect(both.date).toBe('14/10/2026');
    expect(both.consultant).toBe(`${am.consultant} / ${pm.consultant}`);
    expect(both.start).toBe(am.start);
    expect(both.end).toBe(pm.end);
    expect(both.id).not.toBe(am.id);
  });

  it('leaves a single list as it is', () => {
    expect(combineSessions([am])).toBe(am);
  });

  it('remembers which list each patient is on', () => {
    const lists = listOfPatient([am, pm]);
    expect(lists.get(pm.patients[0])).toBe(pm);
    expect(lists.get(am.patients[1])).toBe(am);
  });

  it('keeps ticked lists in run-sheet order and drops lists no longer on the sheet', () => {
    expect(orderedSelection(sessions, [pm.id, 'gone', am.id])).toEqual([am.id, pm.id]);
  });
});

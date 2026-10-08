import { describe, expect, it } from 'vitest';
import { toTsv } from '../src/clipboard';
import { extractSessions } from '../src/extract';
import { buildRow, defaultLayout, type Layout } from '../src/layout';
import { parseTable } from '../src/parse';
import { runSheetCsv } from './fixtures';

const OPTIONS = { flagHospitals: ['POWH', 'STG'], hospitalCodes: ['TEH', 'POWH', 'STG'] };
const patients = extractSessions(parseTable(runSheetCsv()), OPTIONS).flatMap((s) => s.patients);
const byMrn = (mrn: string) => patients.find((p) => p.mrn === mrn)!;

describe('layout', () => {
  it('default layout fills cheat-sheet columns B–K', () => {
    expect(buildRow(byMrn('1000006'), defaultLayout())).toEqual([
      'LEFT',
      '1000006 MOCK, Frank 70F',
      'LE CE/IOL',
      'Croation interpreter',
      '',
      'NKA',
      '',
      '',
      'Risk of falls',
      'STG patient',
    ]);
  });

  it('supports custom order, blank columns and raw run-sheet columns', () => {
    const layout: Layout = {
      includeHeadings: true,
      pasteStart: 'A2',
      columns: [
        { heading: 'Time', source: 'field:time' },
        { heading: 'Notes', source: 'blank' },
        { heading: 'Ward', source: 'raw:CURRENT_WARD' },
        { heading: 'MRN', source: 'field:mrn' },
        { heading: 'Missing', source: 'raw:NOT_A_COLUMN' },
      ],
    };
    expect(buildRow(byMrn('1000001'), layout)).toEqual(['09:20', '', 'DPU TEH', '1000001', '']);
  });
});

describe('toTsv', () => {
  it('quotes cells with line breaks, tabs or quotes so Excel keeps them in one cell', () => {
    expect(toTsv([['a', 'two\nlines', 'say "hi"'], ['b', 'tab\there', '']])).toBe(
      'a\t"two\nlines"\t"say ""hi"""\r\nb\t"tab\there"\t',
    );
  });
});

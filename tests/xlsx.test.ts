import { describe, expect, it } from 'vitest';
import { extractSessions } from '../src/extract';
import { buildRow, defaultLayout } from '../src/layout';
import { findPreppedRows, mergePrepped } from '../src/merge';
import { parseTable } from '../src/parse';
import { readWorkbook, WorkbookError } from '../src/xlsx';
import { runSheetCsv } from './fixtures';
import { makeWorkbook, makeZip } from './xlsx-helper';

const sessions = extractSessions(parseTable(runSheetCsv()), { flagHospitals: [], hospitalCodes: [] }).filter((s) => s.patients.length);
const or1am = sessions.find((s) => s.or === 'OR 1' && s.ampm === 'AM')!;
const layout = defaultLayout();
const prepped = () => or1am.patients.map((p, i) => {
  const row = buildRow(p, layout);
  row[4] = `VA 6/${9 + i}\nIOP 1${i}`;
  return row;
});

describe('reading Excel workbooks', () => {
  it('reads every sheet, with shared strings, numbers, line breaks and empty cells', async () => {
    const book = await makeWorkbook({ First: [['a', null, 3.5], ['line 1\nline 2', 1234567]], 'Second & more': [['x']] }, 2, 1);
    const sheets = await readWorkbook(book);
    expect(sheets.map((s) => s.name)).toEqual(['First', 'Second & more']);
    expect(sheets[0].rows).toEqual([
      ['', '', '', ''],
      ['', 'a', '', '3.5'],
      ['', 'line 1\nline 2', '1234567', ''],
    ]);
  });

  it('reads rich text, inline strings, booleans and cached formula values', async () => {
    const files = {
      'xl/workbook.xml': '<workbook><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="/xl/worksheets/s.xml"/></Relationships>',
      'xl/sharedStrings.xml': '<sst><si><r><rPr><b/></rPr><t>PXF</t></r><r><t xml:space="preserve">, guttata</t></r><rPh><t>skip</t></rPh></si></sst>',
      'xl/worksheets/s.xml':
        '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>Tom &amp; Jo</t></is></c><c r="C1" t="b"><v>1</v></c><c r="D1"><f>1+1</f><v>2</v></c><c r="E1"><v>0.30000000000000004</v></c></row></sheetData></worksheet>',
    };
    const [sheet] = await readWorkbook(await makeZip(files, false));
    expect(sheet.rows).toEqual([['PXF, guttata', 'Tom & Jo', 'TRUE', '2', '0.3']]);
  });

  it('says so for files that aren’t workbooks', async () => {
    await expect(readWorkbook(new TextEncoder().encode('MRN,NAME\n1,A').buffer as ArrayBuffer)).rejects.toThrow(WorkbookError);
  });
});

describe('finding your prepped rows in a workbook', () => {
  it('finds the rows under the headings on the sheet with this list, and merges them', async () => {
    const header = [['Consultant', 'Cedar', null, 'Date', '14/10/2026'], [], [...layout.columns.map((c) => c.heading)]];
    const book = await makeWorkbook({
      'Other list': [['RIGHT', '9999999 NOBODY, Zed 50F']],
      'OR1 AM': [...header, ...prepped(), [], ['Footer note: bring loupes']],
    }, 1, 1);
    const found = findPreppedRows(await readWorkbook(book), or1am, layout)!;
    expect(found.sheet).toBe('OR1 AM');
    expect(found.cell).toBe('B4');
    expect(found.matches).toBe(2);
    expect(found.rows).toEqual(prepped());
    const result = mergePrepped(found.rows, or1am, sessions, layout, (s) => s.or);
    expect(result.summary).toMatchObject({ kept: 2, added: 0, cancelled: 0 });
    expect(result.rows.map((r) => r.cells[4])).toEqual(prepped().map((r) => r[4]));
  });

  it('still finds them when the sheet starts in another column than the settings say', async () => {
    const book = await makeWorkbook({ Sheet1: prepped() }, 7, 3);
    const found = findPreppedRows(await readWorkbook(book), or1am, layout)!;
    expect(found.cell).toBe('D7');
    expect(found.rows).toEqual(prepped());
  });

  it('keeps cancelled rows at the bottom of the block', async () => {
    const rows = [...prepped(), [], ['CANCELLED', 'Not on the latest run sheet'], ['LEFT', '7654321 GONE, Gail 80F', 'LE CE/IOL']];
    const found = findPreppedRows(await readWorkbook(await makeWorkbook({ S: rows }, 5, 1)), or1am, layout)!;
    expect(found.rows).toHaveLength(5);
    const result = mergePrepped(found.rows, or1am, sessions, layout, (s) => s.or);
    expect(result.summary.cancelled).toBe(1);
    expect(result.padTo).toBe(5);
  });

  it('prefers the sheet made for this list when a patient is on two sheets', async () => {
    const [cedar, other] = prepped();
    const stranger = (n: number) => ['LEFT', `80000${n}0 STRANGER, Sam 60M`, 'LE CE/IOL'];
    const book = await makeWorkbook({
      'Other list': [other, stranger(1), stranger(2), stranger(3)],
      'This list': [cedar, other],
    }, 5, 1);
    const found = findPreppedRows(await readWorkbook(book), or1am, layout)!;
    expect(found.sheet).toBe('This list');
    expect(found).toMatchObject({ matches: 2, others: 0 });
  });

  it('finds nothing when no patient on this list is in the workbook', async () => {
    const book = await makeWorkbook({ S: [['RIGHT', '9999999 NOBODY, Zed 50F']] });
    expect(findPreppedRows(await readWorkbook(book), or1am, layout)).toBeNull();
  });
});

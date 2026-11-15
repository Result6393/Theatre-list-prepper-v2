import { describe, expect, it } from 'vitest';
import { involvementColumns, missingHeadings, readInvolvement, withInvolvementColumns } from '../src/involvement';
import { defaultLayout, type Layout } from '../src/layout';
import { defaultSettings, exportSettings, importSettings } from '../src/settings';
import { openWorkbook, saveWorkbook } from '../src/xlsx';
import { makeZip } from './xlsx-helper';

const layout = defaultLayout();
const headings = layout.columns.map((c) => c.heading);

describe('involvement', () => {
  it('has its own two columns after Plan & notes by default', () => {
    expect(headings.slice(-3)).toEqual(['Plan & notes', 'Involvement', 'Involvement notes']);
    expect(involvementColumns(layout)).toEqual({ level: 10, notes: 11 });
  });

  it('finds the columns by heading, wherever they are, or says there are none', () => {
    const custom: Layout = { ...layout, columns: [{ heading: 'Involvement Notes', source: 'blank' }, { heading: 'Patient', source: 'field:patient' }, { heading: 'My role', source: 'blank' }] };
    expect(involvementColumns(custom)).toEqual({ level: 2, notes: 0 });
    expect(involvementColumns({ ...layout, columns: layout.columns.slice(0, 10) })).toEqual({ level: -1, notes: -1 });
  });

  it('reads a level however it’s typed, and leaves other text alone', () => {
    expect(readInvolvement(' supervised')).toBe('Supervised');
    expect(readInvolvement('UNSUPERVISED')).toBe('Unsupervised');
    expect(readInvolvement('')).toBe('');
    expect(readInvolvement('did phaco only')).toBe('');
  });

  it('adds the columns to saved settings still on the old default columns, and only those', () => {
    const old = defaultSettings();
    old.layout.columns = old.layout.columns.slice(0, 10);
    expect(importSettings(exportSettings(old)).layout.columns.map((c) => c.heading)).toEqual(headings);

    const custom = defaultSettings();
    custom.layout.columns = [...custom.layout.columns.slice(0, 10), { heading: 'Time', source: 'field:time' }];
    expect(importSettings(exportSettings(custom)).layout.columns).toHaveLength(11);
    expect(withInvolvementColumns(layout.columns)).toHaveLength(12);
  });

  it('finds the headings an older sheet is missing, only beside your other headings', () => {
    const sheet = { name: 'List', rows: [['Consultant:', 'Smith'], [], [], ['#', ...headings.slice(0, 10)], ['1', 'RIGHT', '1000001 TESTER, Alice']] };
    expect(missingHeadings(sheet, 'B5', layout, [10, 11])).toEqual([
      { cell: 'L4', heading: 'Involvement' },
      { cell: 'M4', heading: 'Involvement notes' },
    ]);
    const done = { name: 'List', rows: [[], [], [], ['#', ...headings], []] };
    expect(missingHeadings(done, 'B5', layout, [10, 11])).toEqual([]);
    // No heading row above the list: nothing is written.
    const bare = { name: 'List', rows: [['Some title'], ['1', 'RIGHT', '1000001 TESTER, Alice']] };
    expect(missingHeadings(bare, 'B2', layout, [10, 11])).toEqual([]);
  });

  it('a new column’s heading and cells are formatted like their neighbours', async () => {
    const sheetXml =
      '<worksheet><sheetData>' +
      '<row r="4"><c r="A4" s="7" t="inlineStr"><is><t>Plan</t></is></c></row>' +
      '<row r="5"><c r="A5" s="3" t="inlineStr"><is><t>x</t></is></c></row>' +
      '</sheetData></worksheet>';
    const book = await openWorkbook(await makeZip({
      '[Content_Types].xml': '<Types/>',
      'xl/workbook.xml': '<workbook xmlns:r="r"><sheets><sheet name="List" sheetId="1" r:id="rId1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml': sheetXml,
    }));
    const result = await saveWorkbook(book, [
      { sheet: 'List', cell: 'A5', rows: [['x', 'Supervised']], styleFromLeft: true },
      { sheet: 'List', cell: 'B4', rows: [['Involvement']], styleFromLeft: true },
    ]);
    const saved = await openWorkbook(result.bytes.buffer as ArrayBuffer);
    expect(saved.sheets[0].rows[3]).toEqual(['Plan', 'Involvement']);
    const entry = saved.entries.find((e) => e.name === 'xl/worksheets/sheet1.xml')!;
    const xml = entry.method === 0 ? new TextDecoder().decode(entry.data) : await new Response(new Blob([entry.data as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
    // The heading looks like the heading beside it (not like the list), the cell like the cell beside it (not like the heading).
    expect(xml).toContain('<c r="B4" s="7" t="inlineStr">');
    expect(xml).toContain('<c r="B5" s="3" t="inlineStr">');
    const again = await saveWorkbook(saved, [{ sheet: 'List', cell: 'B4', rows: [['Involvement']], styleFromLeft: true }]);
    expect(again.changed).toBe(0);
  });
});

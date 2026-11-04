import { describe, expect, it } from 'vitest';
import { extractSessions } from '../src/extract';
import { buildRow, defaultLayout } from '../src/layout';
import { copyBlock, findListBlock, findPreppedRows, mergePrepped, sheetRows } from '../src/merge';
import { parseTable } from '../src/parse';
import { openWorkbook, overwrittenBelow, readWorkbook, saveWorkbook, sheetNameProblem, withEdit, WorkbookError } from '../src/xlsx';
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

describe('saving changes into the workbook', () => {
  const sheetXml =
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="2" max="2" width="30"/></cols><sheetData>' +
    '<row r="4" spans="1:4"><c r="B4" s="2" t="s"><v>0</v></c><c r="C4" s="2" t="s"><v>1</v></c></row>' +
    '<row r="5" spans="1:4"><c r="A5" s="1"><f>ROW()-4</f><v>1</v></c><c r="B5" s="3" t="s"><v>2</v></c><c r="C5" s="3"/><c r="D5" s="4"><v>1234567</v></c></row>' +
    '<row r="6"><c r="A6" s="1"><f>ROW()-4</f><v>2</v></c><c r="B6" s="3" t="s"><v>3</v></c></row>' +
    '<row r="9"><c r="B9" t="s"><v>4</v></c></row>' +
    '</sheetData><mergeCells count="1"><mergeCell ref="E1:F1"/></mergeCells></worksheet>';
  const files = () => ({
    '[Content_Types].xml': '<Types/>',
    'xl/workbook.xml': '<workbook><sheets><sheet name="List" r:id="rId1"/><sheet name="Other" r:id="rId2"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><t>Eye</t></si><si><t>Patient</t></si><si><t>RIGHT</t></si><si><t>LEFT</t></si><si><t>Notes &amp; more</t></si></sst>',
    'xl/worksheets/sheet1.xml': sheetXml,
    'xl/worksheets/sheet2.xml': '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>',
    'xl/styles.xml': '<styleSheet/>',
  });

  it('writes only the changed cells, keeping formatting, formulas and everything else', async () => {
    const book = await openWorkbook(await makeZip(files()));
    const edit = { sheet: 'List', cell: 'B5', rows: [['RIGHT', 'PXF, guttata\nPoor dilation', '1234567'], ['', '', '']] };
    const result = await saveWorkbook(book, [edit]);
    expect(result).toMatchObject({ changed: 2, formulas: 0 });
    const saved = await openWorkbook(result.bytes.buffer as ArrayBuffer);
    const list = saved.sheets.find((s) => s.name === 'List')!;
    expect(list.rows[4].slice(0, 4)).toEqual(['1', 'RIGHT', 'PXF, guttata\nPoor dilation', '1234567']);
    expect(list.rows[5].slice(0, 3)).toEqual(['2', '', '']);
    expect(list.rows[8][1]).toBe('Notes & more');
    expect(saved.sheets.find((s) => s.name === 'Other')!.rows).toEqual([['Eye']]);

    const xml = await sheetText(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="A5" s="1"><f>ROW()-4</f><v>1</v></c>');
    expect(xml).toContain('<c r="C5" s="3" t="inlineStr"><is><t xml:space="preserve">PXF, guttata\nPoor dilation</t></is></c>');
    expect(xml).toContain('<c r="D5" s="4"><v>1234567</v></c>');
    expect(xml).toContain('<c r="B6" s="3"/>');
    expect(xml).toContain('<cols><col min="2" max="2" width="30"/></cols>');
    expect(xml).toContain('<mergeCell ref="E1:F1"/>');
    expect(xml).not.toContain('spans=');
    // Untouched parts are copied byte for byte.
    expect(await sheetText(result.bytes, 'xl/styles.xml')).toBe('<styleSheet/>');
  });

  it('adds rows below with the formatting of the row above, and never overwrites a formula', async () => {
    const book = await openWorkbook(await makeZip(files()));
    const result = await saveWorkbook(book, [{ sheet: 'List', cell: 'A6', rows: [['x', 'LEFT'], ['', 'RIGHT']] }]);
    expect(result).toMatchObject({ changed: 1, formulas: 1 }); // B6 already says LEFT
    const xml = await sheetText(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="A6" s="1"><f>ROW()-4</f><v>2</v></c>');
    expect(xml).toMatch(/<row r="7"><c r="B7" s="3" t="inlineStr"><is><t xml:space="preserve">RIGHT<\/t><\/is><\/c><\/row><row r="9">/);
  });

  it('warns about text a longer list would write over', async () => {
    const book = await openWorkbook(await makeZip(files()));
    const list = book.sheets[0];
    const edit = { sheet: 'List', cell: 'B5', rows: [['RIGHT'], ['LEFT'], ['LEFT'], ['LEFT'], ['RIGHT']] };
    expect(overwrittenBelow(list, edit, 2)).toEqual(['Notes & more']);
    expect(overwrittenBelow(list, edit, 5)).toEqual([]);
    expect(withEdit(list, edit).rows[8][1]).toBe('RIGHT');
  });

  it('writes a valid zip: each part’s checksum matches', async () => {
    const book = await openWorkbook(await makeZip(files()));
    const result = await saveWorkbook(book, [{ sheet: 'List', cell: 'B5', rows: [['LEFT']] }]);
    const { crc32 } = await import('node:zlib');
    const view = new DataView(result.bytes.buffer);
    let at = 0;
    while (view.getUint32(at, true) === 0x04034b50) {
      const method = view.getUint16(at + 8, true);
      const crc = view.getUint32(at + 14, true);
      const size = view.getUint32(18 + at, true);
      const nameLength = view.getUint16(at + 26, true);
      const data = result.bytes.subarray(at + 30 + nameLength, at + 30 + nameLength + size);
      const raw = method === 8 ? new Uint8Array(await new Response(new Blob([data as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer()) : data;
      const name = new TextDecoder().decode(result.bytes.subarray(at + 30, at + 30 + nameLength));
      if (name === 'xl/worksheets/sheet1.xml') expect(crc).toBe(crc32(raw));
      at += 30 + nameLength + size;
    }
    expect(at).toBeGreaterThan(0);
  });
});

async function sheetText(bytes: Uint8Array, path: string): Promise<string> {
  const book = await openWorkbook(bytes.buffer as ArrayBuffer);
  const entry = book.entries.find((e) => e.name === path)!;
  if (entry.method === 0) return new TextDecoder().decode(entry.data);
  return new Response(new Blob([entry.data as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
}

describe('working from your prepped sheet without a run sheet', () => {
  it('finds each sheet’s list and makes its rows editable, cancelled ones last', async () => {
    const rows = [...prepped(), [], ['CANCELLED', 'Not on the latest run sheet'], ['LEFT', '7654321 GONE, Gail 80F', 'LE CE/IOL']];
    const sheets = await readWorkbook(await makeWorkbook({ 'OR1 AM': [['Eye', 'Patient'], ...rows], Template: [['Eye', 'Patient']] }, 4, 1));
    expect(findListBlock(sheets[1], layout)).toBeNull();
    const block = findListBlock(sheets[0], layout)!;
    expect(block).toMatchObject({ sheet: 'OR1 AM', cell: 'B5', matches: 3 });
    const list = sheetRows(block.rows, 'sheet:OR1 AM');
    expect(list.map((r) => [r.cells[1].split(' ')[0], r.status])).toEqual([
      ['1000002', 'sheet'],
      ['1000001', 'sheet'],
      ['7654321', 'cancelled'],
    ]);
    expect(new Set(list.map((r) => r.id)).size).toBe(3);
    // Copying (or saving) it back gives the same block.
    expect(copyBlock(list, layout, block.rows.length)).toEqual(block.rows.map((r) => r.map((c) => c)));
  });
});

describe('saving a list as a new sheet', () => {
  const template =
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><cols><col min="2" max="2" width="30"/></cols><sheetData>' +
    '<row r="1"><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>' +
    '<row r="5"><c r="A5" s="1"><f>ROW()-4</f></c><c r="B5" s="3" t="s"><v>2</v></c><c r="C5" s="3" t="s"><v>3</v></c></row>' +
    '<row r="6"><c r="A6" s="1"><f>ROW()-4</f></c><c r="B6" s="3" t="s"><v>2</v></c><c r="C6" s="3" t="s"><v>4</v></c></row>' +
    '</sheetData><conditionalFormatting sqref="B5:B20"><cfRule type="expression" priority="1"><formula>B5="RIGHT"</formula><extLst><ext uri="x"><x14:id xmlns:x14="y">{A}</x14:id></ext></extLst></cfRule></conditionalFormatting>' +
    '<pageSetup orientation="landscape" r:id="rId1"/><drawing r:id="rId2"/><tableParts count="1"><tablePart r:id="rId3"/></tableParts>' +
    '<extLst><ext uri="z"><x14:conditionalFormattings xmlns:x14="y"/></ext></extLst></worksheet>';
  const files = () => ({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    'xl/workbook.xml': '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Wed OR4" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="styles" Target="styles.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><t>Consultant</t></si><si><t>Willow</t></si><si><t>RIGHT</t></si><si><t>1000001 OLD, One 70F</t></si><si><t>1000002 OLD, Two 70F</t></si></sst>',
    'xl/worksheets/sheet1.xml': template,
    'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships/>',
  });

  it('copies a sheet’s layout into a new tab and puts the list there, leaving the original alone', async () => {
    const book = await openWorkbook(await makeZip(files()));
    const rows = [['LEFT', '2000001 NEW, Nina 60F'], ['', '']];
    const result = await saveWorkbook(book, [{ sheet: 'Thu OR2', cell: 'B5', rows }], [{ name: 'Thu OR2', copyOf: 'Wed OR4' }]);
    const saved = await openWorkbook(result.bytes.buffer as ArrayBuffer);
    expect(saved.sheets.map((s) => s.name)).toEqual(['Wed OR4', 'Thu OR2']);
    const [original, added] = saved.sheets;
    expect(original.rows[4].slice(1, 3)).toEqual(['RIGHT', '1000001 OLD, One 70F']);
    expect(added.rows[0].slice(1, 3)).toEqual(['Consultant', 'Willow']);
    expect(added.rows[4].slice(1, 3)).toEqual(['LEFT', '2000001 NEW, Nina 60F']);
    expect(added.rows[5].slice(1, 3)).toEqual(['', '']);

    const xml = await sheetText(result.bytes, 'xl/worksheets/sheet2.xml');
    expect(xml).toContain('<cols><col min="2" max="2" width="30"/></cols>');
    expect(xml).toContain('<c r="A5" s="1"><f>ROW()-4</f></c>');
    expect(xml).toContain('<c r="C5" s="3" t="inlineStr">');
    expect(xml).toContain('<formula>B5="RIGHT"</formula><extLst><ext uri="x">');
    expect(xml).toContain('<extLst><ext uri="z">');
    expect(xml).toContain('<pageSetup orientation="landscape"/>');
    expect(xml).not.toMatch(/tabSelected|<drawing|tableParts|r:id/);
    expect(await sheetText(result.bytes, 'xl/worksheets/sheet1.xml')).toBe(template);
    expect(await sheetText(result.bytes, 'xl/workbook.xml')).toContain('<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Thu OR2" sheetId="2" r:id="rIdTlp2"/></sheets>');
    expect(await sheetText(result.bytes, 'xl/_rels/workbook.xml.rels')).toContain('Id="rIdTlp2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>');
    expect(await sheetText(result.bytes, '[Content_Types].xml')).toContain('<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  });

  it('declares the relationship prefix on the new sheet, for files that declare it per sheet', async () => {
    const f = files();
    f['xl/workbook.xml'] = '<workbook xmlns="main"><sheets><sheet xmlns:r="rels" name="Wed OR4" sheetId="1" r:id="rId1"/></sheets></workbook>';
    const result = await saveWorkbook(await openWorkbook(await makeZip(f)), [], [{ name: 'New', copyOf: 'Wed OR4' }]);
    const xml = await sheetText(result.bytes, 'xl/workbook.xml');
    const added = xml.match(/<sheet [^>]*name="New"[^>]*\/>/)![0];
    expect(added).toContain('xmlns:r=');
    expect(added).toContain('r:id="rIdTlp2"');
  });

  it('checks the new sheet’s name the way Excel does', async () => {
    expect(sheetNameProblem('Thu OR2', ['Wed OR4'])).toBeNull();
    expect(sheetNameProblem('wed or4', ['Wed OR4'])).toMatch(/already/);
    expect(sheetNameProblem('14/10 OR4', [])).toMatch(/can’t contain/);
    expect(sheetNameProblem('x'.repeat(32), [])).toMatch(/31/);
    expect(sheetNameProblem('  ', [])).toMatch(/name/);
    const book = await openWorkbook(await makeZip(files()));
    await expect(saveWorkbook(book, [], [{ name: 'Wed OR4', copyOf: 'Wed OR4' }])).rejects.toThrow(/already/);
  });
});

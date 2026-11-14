import { describe, expect, it } from 'vitest';
import { imageSize, reportImage, sniffImageType, type ReportImage } from '../src/report';
import { blockPictures, openWorkbook, saveWorkbook, withEdit, type Workbook } from '../src/xlsx';
import { makeZip } from './xlsx-helper';

/** Just enough of a PNG (signature and IHDR) for its type and size; `tag` keeps them apart. */
function png(width: number, height: number, tag = 0): ReportImage {
  const bytes = new Uint8Array(40);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  bytes[39] = tag;
  return reportImage(bytes)!;
}

const WORKBOOK = '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="List" sheetId="1" r:id="rId1"/></sheets></workbook>';
const WORKBOOK_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';
const TYPES = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>';
const rowsXml = [1, 2, 3, 4].map((r) => `<row r="${r}" ht="45" customHeight="1"><c r="B${r}" t="inlineStr"><is><t>Row ${r}</t></is></c></row>`).join('');

async function book(extra: Record<string, string | Uint8Array> = {}, sheetTail = ''): Promise<Workbook> {
  return openWorkbook(
    await makeZip({
      '[Content_Types].xml': TYPES,
      'xl/workbook.xml': WORKBOOK,
      'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cols><col min="2" max="3" width="20" customWidth="1"/></cols><sheetData>${rowsXml}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${sheetTail}</worksheet>`,
      ...extra,
    }),
  );
}

const names = async (b: Workbook) => b.entries.map((e) => e.name);
const text = async (bytes: Uint8Array, part: string) => {
  const reopened = await openWorkbook(bytes.buffer as ArrayBuffer);
  const entry = reopened.entries.find((e) => e.name === part)!;
  return new Response(new Blob([entry.data as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
};

describe('report images', () => {
  it('reads the type and size of PNG and JPEG pictures', () => {
    expect(sniffImageType(png(10, 20).bytes)).toBe('image/png');
    expect(imageSize(png(640, 480).bytes)).toEqual({ width: 640, height: 480 });
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 11, 8, 0x01, 0xe0, 0x02, 0x80, 3, 0, 0, 0]);
    expect(reportImage(jpeg)).toMatchObject({ type: 'image/jpeg', width: 640, height: 480 });
    expect(reportImage(new TextEncoder().encode('not a picture'))).toBeNull();
  });
});

describe('pictures on a list’s rows in Excel', () => {
  const a = png(800, 600, 1);
  const b = png(600, 800, 2);
  const c = png(400, 300, 3);

  it('adds them beside each row, and reads them back on the same rows', async () => {
    const start = await book();
    const result = await saveWorkbook(start, [{ sheet: 'List', cell: 'B2', rows: [['Row 2', 'x'], ['Row 3', 'y']], pictures: [[a, b], [c]] }]);
    expect(result.changed).toBe(2 + 2);
    const reopened = await openWorkbook(result.bytes.buffer as ArrayBuffer);
    const [sheet] = reopened.sheets;
    expect(blockPictures(sheet, 1, 1, 2)).toEqual([[a, b], [c]]);
    // Just right of the two-column block (B:C), so in D onwards.
    expect(sheet.pictures!.every((p) => p.col >= 3)).toBe(true);
    const types = await text(result.bytes, '[Content_Types].xml');
    expect(types).toContain('Extension="png"');
    expect(types).toContain('PartName="/xl/drawings/drawing1.xml"');
    const sheetXml = await text(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(sheetXml).toMatch(/<pageMargins[^>]*\/><drawing xmlns:r="[^"]+" r:id="rIdTlpDrawing1"\/><\/worksheet>/);
    const drawing = await text(result.bytes, 'xl/drawings/drawing1.xml');
    // Scaled to the 45pt (60px) row, keeping the shape: 800×600 → 75×56.
    expect(drawing).toContain(`<xdr:ext cx="${75 * 9525}" cy="${56 * 9525}"/>`);
    expect(drawing).toContain('name="IOL report');
  });

  it('saves nothing when the pictures are as they were, and moves them with their rows without copying the images', async () => {
    const first = await saveWorkbook(await book(), [{ sheet: 'List', cell: 'B2', rows: [['Row 2'], ['Row 3']], pictures: [[a], [b]] }]);
    const saved = await openWorkbook(first.bytes.buffer as ArrayBuffer);
    expect((await saveWorkbook(saved, [{ sheet: 'List', cell: 'B2', rows: [['Row 2'], ['Row 3']], pictures: blockPictures(saved.sheets[0], 1, 1, 2) }])).changed).toBe(0);

    const swapped = await saveWorkbook(saved, [{ sheet: 'List', cell: 'B2', rows: [['Row 3'], ['Row 2']], pictures: [[b], [a]] }]);
    const after = await openWorkbook(swapped.bytes.buffer as ArrayBuffer);
    expect(blockPictures(after.sheets[0], 1, 1, 2)).toEqual([[b], [a]]);
    expect((await names(after)).filter((n) => n.startsWith('xl/media/'))).toHaveLength(2);

    const removed = await saveWorkbook(after, [{ sheet: 'List', cell: 'B2', rows: [['Row 3'], ['Row 2']], pictures: [[b], []] }]);
    const last = await openWorkbook(removed.bytes.buffer as ArrayBuffer);
    expect(blockPictures(last.sheets[0], 1, 1, 2)).toEqual([[b], []]);
    expect((await names(last)).filter((n) => n.startsWith('xl/media/'))).toHaveLength(1);
  });

  it('leaves other pictures on the sheet (like a logo in the heading) alone', async () => {
    const logo = png(100, 50, 9);
    const anchor = (row: number, id: string) =>
      `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="7" name="Logo"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${id}"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
    const start = await book(
      {
        'xl/worksheets/_rels/sheet1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
        'xl/drawings/drawing1.xml': `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchor(0, 'rId1')}${anchor(2, 'rId2')}</xdr:wsDr>`,
        'xl/drawings/_rels/drawing1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/></Relationships>',
        'xl/media/image1.png': logo.bytes,
        'xl/media/image2.png': a.bytes,
      },
      '<drawing r:id="rId1"/>',
    );
    // A picture pasted in Excel on row 3 counts as that patient's.
    expect(blockPictures(start.sheets[0], 1, 1, 3)).toEqual([[], [a], []]);
    const result = await saveWorkbook(start, [{ sheet: 'List', cell: 'B2', rows: [['Row 2'], ['Row 3'], ['Row 4']], pictures: [[a], [], [c]] }]);
    const after = await openWorkbook(result.bytes.buffer as ArrayBuffer);
    expect(after.sheets[0].pictures!.find((p) => p.row === 0)?.image).toEqual(logo);
    expect(blockPictures(after.sheets[0], 1, 1, 3)).toEqual([[a], [], [c]]);
    expect((await names(after)).filter((n) => n.startsWith('xl/drawings/drawing'))).toEqual(['xl/drawings/drawing1.xml']);
  });

  it('keeps pictures with their rows in memory too', () => {
    const sheet = { name: 'List', rows: [['', 'Row 1'], ['', 'Row 2']], pictures: [{ row: 0, col: 3, image: a }, { row: 1, col: 5, image: b }] };
    const edited = withEdit(sheet, { sheet: 'List', cell: 'B1', rows: [['Row 2'], ['Row 1']], pictures: [[b], [a]] });
    expect(blockPictures(edited, 0, 1, 2)).toEqual([[b], [a]]);
    expect(withEdit(sheet, { sheet: 'List', cell: 'B1', rows: [['x'], ['y']] }).pictures).toBe(sheet.pictures);
  });
});

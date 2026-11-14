/** Builds a minimal .xlsx (zip of XML parts) in memory for tests. CRCs are left at 0: the reader ignores them. */
export async function makeZip(files: Record<string, string | Uint8Array>, deflate = true): Promise<ArrayBuffer> {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const raw = typeof content === 'string' ? encoder.encode(content) : content;
    const data = deflate ? new Uint8Array(await new Response(new Blob([raw as BlobPart]).stream().pipeThrough(new CompressionStream('deflate-raw'))).arrayBuffer()) : raw;
    const nameBytes = encoder.encode(name);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(8, deflate ? 8 : 0, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, raw.length, true);
    local.setUint16(26, nameBytes.length, true);
    const dir = new DataView(new ArrayBuffer(46));
    dir.setUint32(0, 0x02014b50, true);
    dir.setUint16(4, 20, true);
    dir.setUint16(6, 20, true);
    dir.setUint16(10, deflate ? 8 : 0, true);
    dir.setUint32(20, data.length, true);
    dir.setUint32(24, raw.length, true);
    dir.setUint16(28, nameBytes.length, true);
    dir.setUint32(42, offset, true);
    chunks.push(new Uint8Array(local.buffer), nameBytes, data);
    central.push(new Uint8Array(dir.buffer), nameBytes);
    offset += 30 + nameBytes.length + data.length;
  }
  const dirSize = central.reduce((n, c) => n + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, Object.keys(files).length, true);
  end.setUint16(10, Object.keys(files).length, true);
  end.setUint32(12, dirSize, true);
  end.setUint32(16, offset, true);
  return new Blob([...chunks, ...central, new Uint8Array(end.buffer)] as BlobPart[]).arrayBuffer();
}

const esc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A workbook whose sheets hold the given rows (strings via shared strings, numbers as numbers). */
export async function makeWorkbook(sheets: Record<string, (string | number | null)[][]>, startRow = 1, startCol = 0): Promise<ArrayBuffer> {
  const shared: string[] = [];
  const col = (i: number) => String.fromCharCode(65 + i);
  const files: Record<string, string> = {};
  const names = Object.keys(sheets);
  files['xl/workbook.xml'] = `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`;
  files['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.map((_, i) => `<Relationship Id="rId${i + 1}" Type="worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}</Relationships>`;
  names.forEach((name, i) => {
    const rows = sheets[name].map((cells, r) => {
      const xml = cells
        .map((v, c) => {
          const ref = `${col(startCol + c)}${startRow + r}`;
          if (v === null || v === '') return `<c r="${ref}" s="1"/>`;
          if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`;
          shared.push(v);
          return `<c r="${ref}" t="s"><v>${shared.length - 1}</v></c>`;
        })
        .join('');
      return `<row r="${startRow + r}">${xml}</row>`;
    });
    files[`xl/worksheets/sheet${i + 1}.xml`] = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join('')}</sheetData></worksheet>`;
  });
  files['xl/sharedStrings.xml'] = `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${shared.map((t) => `<si><t xml:space="preserve">${esc(t)}</t></si>`).join('')}</sst>`;
  return makeZip(files);
}

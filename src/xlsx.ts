/**
 * A small reader for Excel workbooks (.xlsx / .xlsm), entirely in the browser: the file is never
 * uploaded. An .xlsx file is a zip of XML parts; this unzips the parts it needs with the browser's
 * own DecompressionStream and reads each sheet's cell values (as shown, not formulas).
 */

export interface SheetData {
  name: string;
  /** Cell text by row and column, from A1; empty cells are ''. */
  rows: string[][];
}

/** An opened workbook: its sheets' values, plus the file's parts so changes can be saved into it. */
export interface Workbook {
  sheets: SheetData[];
  entries: ZipEntry[];
  /** Sheet name → its part in the file, e.g. "xl/worksheets/sheet2.xml". */
  paths: Map<string, string>;
}

export class WorkbookError extends Error {}

export async function readWorkbook(data: ArrayBuffer): Promise<SheetData[]> {
  return (await openWorkbook(data)).sheets;
}

export async function openWorkbook(data: ArrayBuffer): Promise<Workbook> {
  const entries = readZip(new Uint8Array(data));
  const byName = new Map(entries.map((e) => [e.name, e]));
  const text = (path: string) => {
    const entry = byName.get(path.replace(/^\//, ''));
    return entry ? entryText(entry) : Promise.resolve(undefined);
  };
  const workbook = await text('xl/workbook.xml');
  if (!workbook) throw new WorkbookError('This isn’t an Excel workbook (.xlsx).');

  const shared = parseSharedStrings((await text('xl/sharedStrings.xml')) ?? '');
  const targets = new Map<string, string>();
  for (const rel of tags((await text('xl/_rels/workbook.xml.rels')) ?? '', 'Relationship')) {
    const id = attr(rel, 'Id');
    const target = attr(rel, 'Target');
    if (id && target) targets.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`);
  }

  const sheets: SheetData[] = [];
  const paths = new Map<string, string>();
  for (const sheet of tags(workbook, 'sheet')) {
    const name = decode(attr(sheet, 'name') ?? '');
    const id = attr(sheet, 'r:id') ?? attr(sheet, 'id');
    const path = id ? targets.get(id) : undefined;
    const xml = path ? await text(path) : undefined;
    if (path && xml) {
      sheets.push({ name, rows: parseSheet(xml, shared) });
      paths.set(name, path);
    }
  }
  return { sheets, entries, paths };
}

// ---- Saving changes --------------------------------------------------------------

/** A block of cells to write, e.g. a list's rows from B5. */
export interface SheetEdit {
  sheet: string;
  cell: string;
  rows: string[][];
}

export interface SaveResult {
  bytes: Uint8Array;
  /** Cells whose value changed. */
  changed: number;
  /** Cells left alone because they hold a formula. */
  formulas: number;
}

/**
 * Writes the edits into a copy of the workbook. Only cells whose value differs are touched: their
 * formatting is kept, new cells take the formatting of the cell above, and formula cells are never
 * overwritten. Every other part of the file (other sheets, styles, column widths) is copied as it was.
 */
export async function saveWorkbook(book: Workbook, edits: SheetEdit[]): Promise<SaveResult> {
  let changed = 0;
  let formulas = 0;
  const replaced = new Map<string, string>();
  for (const edit of edits) {
    const path = book.paths.get(edit.sheet);
    const sheet = book.sheets.find((s) => s.name === edit.sheet);
    if (!path || !sheet) throw new WorkbookError(`There’s no sheet “${edit.sheet}” in this workbook any more.`);
    const entry = book.entries.find((e) => e.name === path)!;
    const xml = replaced.get(path) ?? (await entryText(entry));
    const startRow = Number(edit.cell.match(/\d+/)?.[0] ?? 1);
    const startCol = columnIndex(edit.cell);
    const changes: Change[] = [];
    edit.rows.forEach((cells, i) =>
      cells.forEach((value, j) => {
        const current = sheet.rows[startRow - 1 + i]?.[startCol + j] ?? '';
        if (value !== current) changes.push({ row: startRow + i, col: startCol + j, value, blockStart: startRow });
      }),
    );
    if (!changes.length) continue;
    const patched = patchSheet(xml, changes);
    changed += patched.changed;
    formulas += patched.formulas;
    replaced.set(path, patched.xml);
  }
  const entries = await Promise.all(
    book.entries.map(async (e) => (replaced.has(e.name) ? deflateEntry(e, replaced.get(e.name)!) : e)),
  );
  return { bytes: writeZip(entries), changed, formulas };
}

/** The edit applied to the sheet's values in memory (no file is written). */
export function withEdit(sheet: SheetData, edit: SheetEdit): SheetData {
  const startRow = Number(edit.cell.match(/\d+/)?.[0] ?? 1) - 1;
  const startCol = columnIndex(edit.cell);
  const rows = sheet.rows.map((r) => [...r]);
  const width = Math.max(rows[0]?.length ?? 0, startCol + Math.max(0, ...edit.rows.map((r) => r.length)));
  edit.rows.forEach((cells, i) => {
    while (rows.length <= startRow + i) rows.push([]);
    const row = rows[startRow + i];
    while (row.length < width) row.push('');
    cells.forEach((value, j) => (row[startCol + j] = value));
  });
  for (const row of rows) while (row.length < width) row.push('');
  return { name: sheet.name, rows };
}

/** Text in cells below the old block that a longer block would write over. */
export function overwrittenBelow(sheet: SheetData, edit: SheetEdit, oldLength: number): string[] {
  const startRow = Number(edit.cell.match(/\d+/)?.[0] ?? 1) - 1;
  const startCol = columnIndex(edit.cell);
  const lost: string[] = [];
  edit.rows.slice(oldLength).forEach((cells, k) =>
    cells.forEach((value, j) => {
      const current = sheet.rows[startRow + oldLength + k]?.[startCol + j] ?? '';
      if (current.trim() && current !== value) lost.push(current);
    }),
  );
  return lost;
}

interface Change {
  /** 1-based, as in "B5". */
  row: number;
  col: number;
  value: string;
  /** First row of the block, for borrowing the formatting of the cell above. */
  blockStart: number;
}

interface SheetRow {
  r: number;
  open: string;
  prefix: string;
  cells: { col: number; xml: string }[];
}

function patchSheet(xml: string, changes: Change[]): { xml: string; changed: number; formulas: number } {
  const match = xml.match(/<((?:\w+:)?)sheetData\b([^>]*?)(?:\/>|>([\s\S]*?)<\/\1sheetData>)/);
  if (!match) throw new WorkbookError('Couldn’t find the cells in that sheet.');
  const prefix = match[1];
  const rows = new Map<number, SheetRow>();
  let nextRow = 1;
  for (const rowXml of tags(match[3] ?? '', 'row')) {
    const r = Number(attr(rowXml, 'r')) || nextRow;
    nextRow = r + 1;
    const open = rowXml.match(/^<[^>]*>/)![0];
    const selfClosing = open.endsWith('/>');
    const body = selfClosing ? '' : rowXml.slice(open.length, rowXml.lastIndexOf('</'));
    let nextCol = 0;
    const cells = tags(body, 'c').map((c) => {
      const ref = attr(c, 'r');
      const col = ref ? columnIndex(ref) : nextCol;
      nextCol = col + 1;
      return { col, xml: ref ? c : c.replace(/^<((?:\w+:)?)c\b/, `<$1c r="${columnName(col)}${r}"`) };
    });
    // "spans" is only a hint for Excel, and may no longer be right once cells are added.
    const tag = (selfClosing ? open.replace(/\s*\/>$/, '>') : open).replace(/\sspans="[^"]*"/, '');
    rows.set(r, { r, open: tag, prefix, cells });
  }

  let changed = 0;
  let formulas = 0;
  const styleAbove = (row: number, col: number, blockStart: number) => {
    for (let r = row - 1; r >= blockStart - 1 && r >= 1; r--) {
      const style = rows.get(r)?.cells.find((c) => c.col === col)?.xml;
      const s = style ? attr(style, 's') : undefined;
      if (s) return s;
    }
    return undefined;
  };
  for (const change of changes) {
    let row = rows.get(change.row);
    if (!row) {
      row = { r: change.row, open: `<${prefix}row r="${change.row}">`, prefix, cells: [] };
      rows.set(change.row, row);
    }
    const at = row.cells.findIndex((c) => c.col === change.col);
    const old = at >= 0 ? row.cells[at].xml : null;
    if (old && /<(?:\w+:)?f[\s>/]/.test(old)) {
      formulas++;
      continue;
    }
    const style = (old ? attr(old, 's') : undefined) ?? styleAbove(change.row, change.col, change.blockStart);
    const wasNumber = Boolean(old && /<(?:\w+:)?v>/.test(old) && (attr(old, 't') ?? 'n') === 'n');
    const cell = cellXml(prefix, `${columnName(change.col)}${change.row}`, style, change.value, wasNumber);
    if (cell === null) {
      if (at >= 0) row.cells.splice(at, 1);
    } else if (at >= 0) {
      row.cells[at] = { col: change.col, xml: cell };
    } else {
      const before = row.cells.findIndex((c) => c.col > change.col);
      row.cells.splice(before === -1 ? row.cells.length : before, 0, { col: change.col, xml: cell });
    }
    changed++;
  }

  const body = [...rows.values()]
    .sort((a, b) => a.r - b.r)
    .map((r) => `${r.open}${r.cells.map((c) => c.xml).join('')}</${r.prefix}row>`)
    .join('');
  const sheetData = `<${prefix}sheetData${match[2]}>${body}</${prefix}sheetData>`;
  return { xml: xml.replace(match[0], () => sheetData), changed, formulas };
}

function cellXml(p: string, ref: string, style: string | undefined, value: string, asNumber: boolean): string | null {
  const s = style ? ` s="${style}"` : '';
  if (value === '') return style ? `<${p}c r="${ref}"${s}/>` : null;
  if (asNumber && /^-?\d+(\.\d+)?$/.test(value)) return `<${p}c r="${ref}"${s}><${p}v>${value}</${p}v></${p}c>`;
  return `<${p}c r="${ref}"${s} t="inlineStr"><${p}is><${p}t xml:space="preserve">${escapeXml(value)}</${p}t></${p}is></${p}c>`;
}

function escapeXml(text: string): string {
  return text
    .replace(/\r\n?/g, '\n')
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ---- Zip -------------------------------------------------------------------------

interface ZipEntry {
  name: string;
  nameBytes: Uint8Array;
  method: number;
  flags: number;
  time: number;
  date: number;
  crc: number;
  size: number;
  /** The part as stored in the file (compressed when method is 8). */
  data: Uint8Array;
  external: number;
}

/** Reads a zip's parts (still compressed: each is inflated only when it's needed). */
function readZip(bytes: Uint8Array): ZipEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end === -1) throw new WorkbookError('This isn’t an Excel workbook (.xlsx). Old .xls files need saving as .xlsx first.');
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const entries: ZipEntry[] = [];
  for (let n = 0; n < count && view.getUint32(at, true) === 0x02014b50; n++) {
    const nameLength = view.getUint16(at + 28, true);
    const nameBytes = bytes.slice(at + 46, at + 46 + nameLength);
    const compressed = view.getUint32(at + 20, true);
    const local = view.getUint32(at + 42, true);
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    entries.push({
      name: new TextDecoder().decode(nameBytes),
      nameBytes,
      method: view.getUint16(at + 10, true),
      flags: view.getUint16(at + 8, true) & ~0x0008,
      time: view.getUint16(at + 12, true),
      date: view.getUint16(at + 14, true),
      crc: view.getUint32(at + 16, true),
      size: view.getUint32(at + 24, true),
      data: bytes.subarray(start, start + compressed),
      external: view.getUint32(at + 38, true),
    });
    at += 46 + nameLength + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
  }
  return entries;
}

async function entryText(entry: ZipEntry): Promise<string> {
  if (entry.method === 0) return new TextDecoder().decode(entry.data);
  if (entry.method !== 8) throw new WorkbookError('This workbook uses a compression the app can’t read.');
  const stream = new Blob([entry.data as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Response(stream).text();
}

async function deflateEntry(entry: ZipEntry, text: string): Promise<ZipEntry> {
  const raw = new TextEncoder().encode(text);
  const stream = new Blob([raw as BlobPart]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  const data = new Uint8Array(await new Response(stream).arrayBuffer());
  return { ...entry, method: 8, crc: crc32(raw), size: raw.length, data };
}

function writeZip(entries: ZipEntry[]): Uint8Array {
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;
  for (const e of entries) {
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, e.flags, true);
    local.setUint16(8, e.method, true);
    local.setUint16(10, e.time, true);
    local.setUint16(12, e.date, true);
    local.setUint32(14, e.crc, true);
    local.setUint32(18, e.data.length, true);
    local.setUint32(22, e.size, true);
    local.setUint16(26, e.nameBytes.length, true);
    const central = new DataView(new ArrayBuffer(46));
    central.setUint32(0, 0x02014b50, true);
    central.setUint16(4, 20, true);
    central.setUint16(6, 20, true);
    central.setUint16(8, e.flags, true);
    central.setUint16(10, e.method, true);
    central.setUint16(12, e.time, true);
    central.setUint16(14, e.date, true);
    central.setUint32(16, e.crc, true);
    central.setUint32(20, e.data.length, true);
    central.setUint32(24, e.size, true);
    central.setUint16(28, e.nameBytes.length, true);
    central.setUint32(38, e.external, true);
    central.setUint32(42, offset, true);
    locals.push(new Uint8Array(local.buffer), e.nameBytes, e.data);
    centrals.push(new Uint8Array(central.buffer), e.nameBytes);
    offset += 30 + e.nameBytes.length + e.data.length;
  }
  const dirSize = centrals.reduce((n, c) => n + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, entries.length, true);
  end.setUint16(10, entries.length, true);
  end.setUint32(12, dirSize, true);
  end.setUint32(16, offset, true);
  const parts = [...locals, ...centrals, new Uint8Array(end.buffer)];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

let crcTable: Uint32Array | null = null;

function crc32(bytes: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

// ---- Sheet XML -------------------------------------------------------------------

/** Elements named `name` (with or without a namespace prefix), as their full source text. */
function tags(xml: string, name: string): string[] {
  const re = new RegExp(`<(?:\\w+:)?${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</(?:\\w+:)?${name}>)`, 'g');
  return xml.match(re) ?? [];
}

function attr(element: string, name: string): string | undefined {
  const open = element.match(/^<[^>]*>/)?.[0] ?? '';
  return open.match(new RegExp(`\\s${name.replace(':', '\\:')}="([^"]*)"`))?.[1];
}

/** Text of all <t> runs, skipping phonetic guides (<rPh>). */
function runText(xml: string): string {
  return tags(xml.replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g, ''), 't')
    .map((t) => decode(t.replace(/^<[^>]*>/, '').replace(/<\/[^>]*>$/, '')))
    .join('');
}

function parseSharedStrings(xml: string): string[] {
  return tags(xml, 'si').map(runText);
}

function parseSheet(xml: string, shared: string[]): string[][] {
  const rows: string[][] = [];
  let nextRow = 0;
  for (const row of tags(xml, 'row')) {
    const r = Number(attr(row, 'r'));
    const rowIndex = r > 0 ? r - 1 : nextRow;
    nextRow = rowIndex + 1;
    const cells: string[] = [];
    let nextCol = 0;
    for (const cell of tags(row, 'c')) {
      const ref = attr(cell, 'r');
      const col = ref ? columnIndex(ref) : nextCol;
      nextCol = col + 1;
      const value = cellValue(cell, shared);
      if (value === '') continue;
      while (cells.length < col) cells.push('');
      cells[col] = value;
    }
    while (rows.length < rowIndex) rows.push([]);
    rows[rowIndex] = cells;
  }
  const width = Math.max(0, ...rows.map((r) => r.length));
  return rows.map((r) => Array.from({ length: width }, (_, i) => r[i] ?? ''));
}

function cellValue(cell: string, shared: string[]): string {
  const type = attr(cell, 't') ?? 'n';
  if (type === 'inlineStr') return runText(cell).replace(/\r\n?/g, '\n');
  const v = cell.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/)?.[1];
  if (v === undefined) return '';
  const raw = decode(v);
  if (type === 's') return (shared[Number(raw)] ?? '').replace(/\r\n?/g, '\n');
  if (type === 'b') return raw === '1' ? 'TRUE' : 'FALSE';
  if (type === 'n' && /^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(raw)) {
    // Plain numbers as Excel shows them: 1234567 rather than 1.234567E6, and no float noise.
    const n = Number(raw);
    return Number.isInteger(n) ? String(n) : String(Number(n.toPrecision(12)));
  }
  return raw.replace(/\r\n?/g, '\n');
}


/** "B5" → 1 (column B, counting from 0). */
export function columnIndex(ref: string): number {
  const letters = ref.match(/^[A-Z]+/i)?.[0].toUpperCase() ?? 'A';
  return [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

/** 1 → "B". */
export function columnName(index: number): string {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function decode(text: string): string {
  return text
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/_x000D_/g, '');
}

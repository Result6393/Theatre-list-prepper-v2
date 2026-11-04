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

export class WorkbookError extends Error {}

export async function readWorkbook(data: ArrayBuffer): Promise<SheetData[]> {
  const part = await unzip(new Uint8Array(data));
  const text = (path: string) => part(path.replace(/^\//, ''));
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
  for (const sheet of tags(workbook, 'sheet')) {
    const name = decode(attr(sheet, 'name') ?? '');
    const id = attr(sheet, 'r:id') ?? attr(sheet, 'id');
    const path = id ? targets.get(id) : undefined;
    const xml = path ? await text(path) : undefined;
    if (xml) sheets.push({ name, rows: parseSheet(xml, shared) });
  }
  return sheets;
}

// ---- Zip -------------------------------------------------------------------------

/** Reads a zip's directory; each part is inflated only when it's asked for. */
async function unzip(bytes: Uint8Array): Promise<(path: string) => Promise<string> | undefined> {
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
  const entries = new Map<string, { local: number; size: number; method: number }>();
  for (let n = 0; n < count && view.getUint32(at, true) === 0x02014b50; n++) {
    const nameLength = view.getUint16(at + 28, true);
    const name = new TextDecoder().decode(bytes.subarray(at + 46, at + 46 + nameLength));
    entries.set(name, { method: view.getUint16(at + 10, true), size: view.getUint32(at + 20, true), local: view.getUint32(at + 42, true) });
    at += 46 + nameLength + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
  }
  const cache = new Map<string, Promise<string>>();
  return (path) => {
    const entry = entries.get(path);
    if (!entry) return undefined;
    if (!cache.has(path)) {
      const start = entry.local + 30 + view.getUint16(entry.local + 26, true) + view.getUint16(entry.local + 28, true);
      cache.set(path, inflate(bytes.subarray(start, start + entry.size), entry.method));
    }
    return cache.get(path);
  };
}

async function inflate(data: Uint8Array, method: number): Promise<string> {
  if (method === 0) return new TextDecoder().decode(data);
  if (method !== 8) throw new WorkbookError('This workbook uses a compression the app can’t read.');
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Response(stream).text();
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

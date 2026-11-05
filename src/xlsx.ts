/**
 * A small reader for Excel workbooks (.xlsx / .xlsm), entirely in the browser: the file is never
 * uploaded. An .xlsx file is a zip of XML parts; this unzips the parts it needs with the browser's
 * own DecompressionStream and reads each sheet's cell values (as shown, not formulas).
 */

export interface SheetData {
  name: string;
  /** Cell text by row and column, from A1; empty cells are ''. */
  rows: string[][];
  /** Merged ranges, e.g. "B1:C1". */
  merges?: string[];
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
      const merges = tags(xml, 'mergeCell').map((m) => attr(m, 'ref')).filter((r): r is string => Boolean(r));
      sheets.push({ name, rows: parseSheet(xml, shared), merges });
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
  /** Write whole numbers as numbers even in cells that were empty (e.g. the OR number). */
  numbers?: boolean;
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
export interface SaveOptions {
  /** For a new file: keep only these sheets (in their order); the rest are removed. */
  keepOnly?: string[];
  /** The sheet the file opens on (and the only selected tab). */
  activeSheet?: string;
}

export async function saveWorkbook(book: Workbook, edits: SheetEdit[], added: NewSheet[] = [], options: SaveOptions = {}): Promise<SaveResult> {
  let changed = 0;
  let formulas = 0;
  const replaced = new Map<string, string>();
  const entries = [...book.entries];
  const paths = new Map(book.paths);
  const sheets = [...book.sheets];
  const current = async (path: string) => replaced.get(path) ?? (await entryText(entries.find((e) => e.name === path)!));

  for (const add of added) {
    const problem = sheetNameProblem(add.name, sheets.map((s) => s.name));
    if (problem) throw new WorkbookError(problem);
    const templatePath = paths.get(add.copyOf);
    const template = sheets.find((s) => s.name === add.copyOf);
    if (!templatePath || !template) throw new WorkbookError(`There’s no sheet “${add.copyOf}” to copy.`);
    const taken = new Set(entries.map((e) => e.name));
    let n = 1;
    while (taken.has(`xl/worksheets/sheet${n}.xml`)) n++;
    const path = `xl/worksheets/sheet${n}.xml`;
    const templateEntry = entries.find((e) => e.name === templatePath)!;
    entries.push({ ...templateEntry, name: path, nameBytes: new TextEncoder().encode(path) });
    replaced.set(path, copyableSheet(await current(templatePath)));
    paths.set(add.name, path);
    sheets.push({ name: add.name, rows: template.rows.map((r) => [...r]), merges: template.merges });

    const relsPath = 'xl/_rels/workbook.xml.rels';
    const rels = await current(relsPath);
    const ids = new Set(tags(rels, 'Relationship').map((r) => attr(r, 'Id')));
    let rid = `rIdTlp${n}`;
    while (ids.has(rid)) rid += 'x';
    replaced.set(relsPath, appendChild(rels, 'Relationships', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`));

    const workbook = await current('xl/workbook.xml');
    const existing = tags(workbook, 'sheet');
    const prefix = existing[0]?.match(/^<(\w+:)?/)?.[1] ?? '';
    const relPrefix = existing[0]?.match(/\s(\w+):id="/)?.[1] ?? 'r';
    const sheetId = Math.max(0, ...existing.map((t) => Number(attr(t, 'sheetId')) || 0)) + 1;
    // The relationships namespace is declared on the element too: some files declare it per sheet, not once.
    const sheetTag = `<${prefix}sheet xmlns:${relPrefix}="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="${escapeXml(add.name).replace(/"/g, '&quot;')}" sheetId="${sheetId}" ${relPrefix}:id="${rid}"/>`;
    replaced.set('xl/workbook.xml', asFirstSheet(workbook, sheetTag));
    sheets.unshift(sheets.pop()!);

    const types = await current('[Content_Types].xml');
    replaced.set('[Content_Types].xml', appendChild(types, 'Types', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`));
  }

  for (const edit of edits) {
    const path = paths.get(edit.sheet);
    const sheet = sheets.find((s) => s.name === edit.sheet);
    if (!path || !sheet) throw new WorkbookError(`There’s no sheet “${edit.sheet}” in this workbook any more.`);
    const entry = entries.find((e) => e.name === path)!;
    const xml = replaced.get(path) ?? (await entryText(entry));
    const startRow = Number(edit.cell.match(/\d+/)?.[0] ?? 1);
    const startCol = columnIndex(edit.cell);
    const changes: Change[] = [];
    edit.rows.forEach((cells, i) =>
      cells.forEach((value, j) => {
        const current = sheet.rows[startRow - 1 + i]?.[startCol + j] ?? '';
        if (value !== current) changes.push({ row: startRow + i, col: startCol + j, value, blockStart: startRow, number: edit.numbers });
      }),
    );
    if (!changes.length) continue;
    const patched = patchSheet(xml, changes);
    changed += patched.changed;
    formulas += patched.formulas;
    // A list that runs past the old one takes its colours, dropdowns and print area along.
    const lastRow = startRow + edit.rows.reduce((last, cells, i) => (cells.some((v) => v !== '') ? i : last), -1);
    const isList = edit.rows.length > 1;
    replaced.set(path, isList ? extendRanges(patched.xml, startRow, lastRow) : patched.xml);
    if (isList) {
      const index = sheetOrder(await current('xl/workbook.xml')).indexOf(edit.sheet);
      replaced.set('xl/workbook.xml', extendPrintArea(await current('xl/workbook.xml'), edit.sheet, index, startRow, lastRow));
    }
  }
  let kept = entries;
  if (options.keepOnly) kept = await keepOnlySheets(entries, paths, replaced, current, options.keepOnly, options.activeSheet ?? options.keepOnly[0]);
  const written = await Promise.all(kept.map(async (e) => (replaced.has(e.name) ? deflateEntry(e, replaced.get(e.name)!) : e)));
  return { bytes: writeZip(written), changed, formulas };
}

/** "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels" */
function relsPathOf(path: string): string {
  const at = path.lastIndexOf('/');
  return `${path.slice(0, at + 1)}_rels/${path.slice(at + 1)}.rels`;
}

/** A relationship's target as a part name, e.g. "../comments1.xml" from a sheet → "xl/comments1.xml". */
function resolveTarget(from: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = from.split('/').slice(0, -1);
  for (const piece of target.split('/')) {
    if (piece === '..') parts.pop();
    else if (piece !== '.') parts.push(piece);
  }
  return parts.join('/');
}

/**
 * Removes every sheet but `keep` from the workbook, with what belongs only to them (their notes,
 * drawings and so on), their print areas and other per-sheet names, and the calculation chain
 * (Excel rebuilds it). The file opens on `active`, the only selected tab.
 */
async function keepOnlySheets(
  entries: ZipEntry[],
  paths: Map<string, string>,
  replaced: Map<string, string>,
  current: (path: string) => Promise<string>,
  keep: string[],
  active: string,
): Promise<ZipEntry[]> {
  const workbookPath = 'xl/workbook.xml';
  const relsPath = 'xl/_rels/workbook.xml.rels';
  let workbook = await current(workbookPath);
  let rels = await current(relsPath);
  const names = new Set(entries.map((e) => e.name));
  const order = sheetOrder(workbook);
  const removed = order.filter((n) => !keep.includes(n));
  const keptOrder = order.filter((n) => keep.includes(n));
  const drop = new Set<string>();

  // Parts each removed sheet uses (notes, drawings…), unless something we keep uses them too.
  const usedBy = new Map<string, Set<string>>();
  for (const name of names) {
    if (!name.endsWith('.rels')) continue;
    const source = name.replace(/_rels\/([^/]+)\.rels$/, '$1');
    for (const rel of tags(await current(name), 'Relationship')) {
      if (attr(rel, 'TargetMode') === 'External') continue;
      const target = resolveTarget(source, attr(rel, 'Target') ?? '');
      if (!usedBy.has(target)) usedBy.set(target, new Set());
      usedBy.get(target)!.add(source);
    }
  }
  const dropPart = (part: string) => {
    if (drop.has(part) || !names.has(part)) return;
    drop.add(part);
    const partRels = relsPathOf(part);
    if (names.has(partRels)) {
      drop.add(partRels);
      for (const [target, sources] of usedBy) {
        if (sources.has(part) && [...sources].every((s) => s === part || drop.has(s))) dropPart(target);
      }
    }
  };

  for (const sheet of removed) {
    const path = paths.get(sheet);
    const tag = tags(workbook, 'sheet').find((t) => decode(attr(t, 'name') ?? '') === sheet);
    if (!tag) continue;
    const rid = tag.match(/\s\w+:id="([^"]+)"/)?.[1];
    workbook = workbook.replace(tag, '');
    if (rid) rels = rels.replace(new RegExp(`<(?:\\w+:)?Relationship\\b[^>]*\\sId="${rid}"[^>]*/>`), '');
    if (path) dropPart(path);
  }

  // Per-sheet names follow their sheet to its new position; those of removed sheets go.
  const quotedRemoved = removed.map((n) => [`'${n.replace(/'/g, "''")}'!`, `${n}!`]).flat();
  workbook = workbook.replace(/<(?:\w+:)?definedName\b[^>]*>[^<]*<\/(?:\w+:)?definedName>/g, (whole) => {
    const local = whole.match(/localSheetId="(\d+)"/)?.[1];
    if (local !== undefined) {
      const at = keptOrder.indexOf(order[Number(local)]);
      return at === -1 ? '' : whole.replace(/localSheetId="\d+"/, `localSheetId="${at}"`);
    }
    return quotedRemoved.some((q) => whole.includes(q)) ? '' : whole;
  });
  workbook = workbook.replace(/<((?:\w+:)?)definedNames\b[^>]*>\s*<\/\1definedNames>/, '');
  const activeAt = Math.max(0, keptOrder.indexOf(active));
  workbook = workbook.replace(/<((?:\w+:)?)workbookView\b([^>]*?)(\/?)>/g, (_, p: string, attrs: string, close: string) => {
    const a = attrs.replace(/\s(activeTab|firstSheet)="\d+"/g, '');
    return `<${p}workbookView${a} activeTab="${activeAt}"${close}>`;
  });

  // Excel rebuilds the calculation chain; an old one listing removed sheets would need repairing.
  for (const rel of tags(rels, 'Relationship')) {
    if (/calcChain/.test(attr(rel, 'Type') ?? '')) {
      rels = rels.replace(rel, '');
      drop.add(resolveTarget(workbookPath, attr(rel, 'Target') ?? ''));
    }
  }
  replaced.set(workbookPath, workbook);
  replaced.set(relsPath, rels);

  // Only the sheet the file opens on is a selected tab.
  for (const sheet of keptOrder) {
    const path = paths.get(sheet);
    if (!path) continue;
    let xml = (await current(path)).replace(/\stabSelected="(?:1|true)"/g, '');
    if (sheet === active) xml = xml.replace(/<((?:\w+:)?)sheetView\b/, (m) => `${m} tabSelected="1"`);
    replaced.set(path, xml);
  }

  const typesPath = '[Content_Types].xml';
  let types = await current(typesPath);
  for (const part of drop) types = types.replace(new RegExp(`<(?:\\w+:)?Override\\b[^>]*PartName="/${part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), '');
  replaced.set(typesPath, types);
  for (const part of drop) replaced.delete(part);
  return entries.filter((e) => !drop.has(e.name));
}

/** A sheet to add, starting as a copy of another (its headings, formatting and formulas). */
export interface NewSheet {
  name: string;
  copyOf: string;
}

/** Why Excel wouldn't accept this sheet name, or null if it's fine. */
export function sheetNameProblem(name: string, existing: string[]): string | null {
  const trimmed = name.trim();
  if (!trimmed) return 'Give the new sheet a name.';
  if (trimmed.length > 31) return 'Sheet names can be at most 31 characters.';
  if (/[\\/?*[\]:]/.test(trimmed)) return 'Sheet names can’t contain \\ / ? * [ ] or :.';
  if (/^'|'$/.test(trimmed)) return 'Sheet names can’t start or end with an apostrophe.';
  if (trimmed.toLowerCase() === 'history') return '“History” is reserved by Excel.';
  if (existing.some((e) => e.toLowerCase() === trimmed.toLowerCase())) return `There’s already a sheet called “${trimmed}”.`;
  return null;
}

/**
 * A sheet's XML made safe to copy: parts it shares with the original through relationships
 * (drawings, comments, tables, links, add-on extensions) are left out, and it isn't the selected tab.
 */
function copyableSheet(xml: string): string {
  let out = removeElements(xml, ['tableParts', 'hyperlinks', 'oleObjects', 'controls', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture']);
  // Add-on extensions are kept unless they link to other parts (e.g. slicers).
  out = removeElements(out, ['extLst'], (element) => /\s\w+:id="/.test(element));
  return out.replace(/\s\w+:id="[^"]*"/g, '').replace(/\stabSelected="(?:1|true)"/g, '');
}

/** Removes whole elements by name (with any namespace prefix), matching nested ones correctly. */
function removeElements(xml: string, names: string[], only: (element: string) => boolean = () => true): string {
  const tag = new RegExp(`<(/?)(?:\\w+:)?(${names.join('|')})\\b[^>]*?(/?)>`, 'g');
  const cuts: [number, number][] = [];
  const open: { name: string; start: number }[] = [];
  for (const m of xml.matchAll(tag)) {
    const [text, closing, name, selfClosing] = m;
    const at = m.index!;
    if (selfClosing) {
      if (!open.length) cuts.push([at, at + text.length]);
    } else if (!closing) {
      open.push({ name, start: at });
    } else {
      const top = open.pop();
      if (top && !open.length) cuts.push([top.start, at + text.length]);
    }
  }
  let out = xml;
  for (const [start, end] of cuts.reverse()) if (only(xml.slice(start, end))) out = out.slice(0, start) + out.slice(end);
  return out;
}

/**
 * Puts a new sheet first, as the leftmost tab. Things that point at sheets by position move along
 * with them: per-sheet names such as print areas (localSheetId), and the tab Excel opens on
 * (activeTab), so the sheet you last had open (and selected) is still the one shown.
 */
function asFirstSheet(workbook: string, sheetTag: string): string {
  const sheets = workbook.match(/<((?:\w+:)?)sheets\b[^>]*>/);
  if (!sheets) return appendChild(workbook, 'sheets', sheetTag);
  const at = sheets.index! + sheets[0].length;
  let out = workbook.slice(0, at) + sheetTag + workbook.slice(at);
  out = out.replace(/(<(?:\w+:)?definedName\b[^>]*?\slocalSheetId=")(\d+)"/g, (_, open: string, n: string) => `${open}${Number(n) + 1}"`);
  out = out.replace(/<((?:\w+:)?)workbookView\b([^>]*?)(\/?)>/g, (_, p: string, attrs: string, close: string) => {
    let a = /\sactiveTab="\d+"/.test(attrs) ? attrs.replace(/\sactiveTab="(\d+)"/, (_m, n: string) => ` activeTab="${Number(n) + 1}"`) : `${attrs} activeTab="1"`;
    // Scroll the tab bar back to the start, so the new tab is in view.
    a = a.replace(/\sfirstSheet="\d+"/, '');
    return `<${p}workbookView${a}${close}>`;
  });
  return out;
}

/** Adds an element at the end of <parent> (which may be written as an empty <parent/>). */
function appendChild(xml: string, parent: string, child: string): string {
  const close = new RegExp(`</((?:\\w+:)?)${parent}>`);
  if (close.test(xml)) return xml.replace(close, (m) => `${child}${m}`);
  const empty = new RegExp(`<((?:\\w+:)?)${parent}\\b([^>]*?)\\s*/>`);
  if (empty.test(xml)) return xml.replace(empty, (_, p: string, attrs: string) => `<${p}${parent}${attrs}>${child}</${p}${parent}>`);
  throw new WorkbookError(`This workbook is missing its ${parent} list.`);
}

/**
 * One cell range, or a space-separated list of them, with those on the list's rows carried down to
 * `lastRow`. With `fromAbove`, a range that starts higher up (like a print area from A1) counts too.
 */
function extendRefs(refs: string, blockStart: number, lastRow: number, fromAbove = false): string {
  return refs
    .split(/\s+/)
    .map((ref) => {
      const m = ref.match(/^(\$?[A-Z]{1,3}\$?)(\d+):(\$?[A-Z]{1,3}\$?)(\d+)$/i);
      if (!m) return ref;
      const top = Number(m[2]);
      const bottom = Number(m[4]);
      const onList = (fromAbove || top >= blockStart) && bottom >= blockStart && bottom < lastRow;
      return onList ? `${m[1]}${top}:${m[3]}${lastRow}` : ref;
    })
    .join(' ');
}

/** Conditional formatting and dropdowns (data validation) on the list's rows reach its new last row. */
function extendRanges(xml: string, blockStart: number, lastRow: number): string {
  return xml
    .replace(/(<(?:\w+:)?(?:conditionalFormatting|dataValidation)\b[^>]*?\ssqref=")([^"]*)"/g, (_, open: string, refs: string) => `${open}${extendRefs(refs, blockStart, lastRow)}"`)
    .replace(/(<(\w+:)?sqref>)([^<]*)(<\/\2?sqref>)/g, (_, open: string, _p: string, refs: string, close: string) => `${open}${extendRefs(refs, blockStart, lastRow)}${close}`);
}

/** The sheets' names in tab order. */
function sheetOrder(workbook: string): string[] {
  return tags(workbook, 'sheet').map((t) => decode(attr(t, 'name') ?? ''));
}

/** The sheet's print area covers the list's new last row. */
function extendPrintArea(workbook: string, sheet: string, index: number, blockStart: number, lastRow: number): string {
  return workbook.replace(/(<(?:\w+:)?definedName\b[^>]*>)([^<]*)(<\/(?:\w+:)?definedName>)/g, (whole, open: string, text: string, close: string) => {
    if (!/name="_xlnm\.Print_Area"/.test(open)) return whole;
    const local = open.match(/localSheetId="(\d+)"/)?.[1];
    const quoted = `'${sheet.replace(/'/g, "''")}'`;
    const forThisSheet = local !== undefined ? Number(local) === index : text.startsWith(`${quoted}!`) || text.startsWith(`${sheet}!`);
    if (!forThisSheet) return whole;
    const updated = text.replace(/!(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)/gi, (_m, range: string) => `!${extendRefs(range, blockStart, lastRow, true)}`);
    return `${open}${updated}${close}`;
  });
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
  return { name: sheet.name, rows, merges: sheet.merges };
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
  number?: boolean;
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
  /**
   * A row below the old list starts as a copy of the row above it: its height, its formatting, and
   * formulas that don't point at other cells (like the "#" column's =ROW()-4), so a longer list
   * looks and numbers like the rest.
   */
  const newRow = (r: number, blockStart: number): SheetRow => {
    let above: SheetRow | undefined;
    for (let a = r - 1; a >= blockStart && !above; a--) above = rows.get(a);
    if (!above) return { r, open: `<${prefix}row r="${r}">`, prefix, cells: [] };
    const open = above.open.replace(/\sr="\d+"/, ` r="${r}"`);
    const cells = above.cells.flatMap(({ col, xml: c }) => {
      const ref = `${columnName(col)}${r}`;
      const formula = c.match(/<((?:\w+:)?)f\b([^>]*)>([^<]*)<\/\1f>/);
      if (formula && !/\bt="shared"/.test(formula[2]) && !/\$?[A-Z]{1,3}\$?\d+/.test(formula[3])) {
        return [{ col, xml: `<${prefix}c r="${ref}"${attr(c, 's') ? ` s="${attr(c, 's')}"` : ''}><${formula[1]}f${formula[2]}>${formula[3]}</${formula[1]}f></${prefix}c>` }];
      }
      const style = attr(c, 's');
      return style ? [{ col, xml: `<${prefix}c r="${ref}" s="${style}"/>` }] : [];
    });
    return { r, open, prefix, cells };
  };
  for (const change of changes) {
    let row = rows.get(change.row);
    if (!row) {
      row = newRow(change.row, change.blockStart);
      rows.set(change.row, row);
    }
    const at = row.cells.findIndex((c) => c.col === change.col);
    const old = at >= 0 ? row.cells[at].xml : null;
    if (old && /<(?:\w+:)?f[\s>/]/.test(old)) {
      formulas++;
      continue;
    }
    const style = (old ? attr(old, 's') : undefined) ?? styleAbove(change.row, change.col, change.blockStart);
    const wasNumber = Boolean(change.number || (old && /<(?:\w+:)?v>/.test(old) && (attr(old, 't') ?? 'n') === 'n'));
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

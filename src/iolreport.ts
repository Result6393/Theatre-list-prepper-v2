/**
 * Reading an IOL calculation report (laid out like the IOLMaster 700's) from OCR'd words, and checking
 * the lens you picked against it. Only ever a cross-check: nothing here changes what you picked.
 */

import { srktPower } from './srkt';

export interface OcrWord {
  text: string;
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

export interface LensRow {
  /** For a toric calculation, what's printed before the power: the model's number (e.g. "100") or "DCB". */
  model: string;
  power: number;
  /** Predicted refraction (spherical equivalent). */
  ref: number;
  cyl?: number;
  axis?: number;
  /** Predicted residual cylinder (toric calculation). */
  refCyl?: number;
  /** Where it is on the screenshot (for highlighting). */
  words?: OcrWord[];
}

export interface LensTable {
  /** The lens heading as read, e.g. "Johnson & Johnson Vision Tecnis DCB00 Preload". */
  title: string;
  toric: boolean;
  rows: LensRow[];
  titleWords?: OcrWord[];
}

export interface EyeReading {
  /** Target refraction the report was calculated for. */
  target: number | null;
  /** Axial length (mm) and mean keratometry (SE, in D), if they could be read. */
  axialLength?: number;
  k?: number;
  tables: LensTable[];
  /** Where the eye's heading ("OD right") and ΔK are on the screenshot. */
  headingWords?: OcrWord[];
  deltaKWords?: OcrWord[];
}

export interface ReportReading {
  /** As printed beside Patient and Patient ID, if they could be read. */
  name: string;
  id: string;
  /** The report's own warnings, e.g. a difference in axial length between the eyes. */
  warnings: string[];
  right?: EyeReading;
  left?: EyeReading;
  nameWords?: OcrWord[];
  idWords?: OcrWord[];
  dobWords?: OcrWord[];
}

interface Row {
  words: OcrWord[];
  y: number;
}

const centre = (w: OcrWord) => (w.x0 + w.x1) / 2;
const median = (values: number[]) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)] ?? 10;
const text = (words: OcrWord[]) => words.map((w) => w.text).join(' ');

/** Words grouped into rows by their height on the page (OCR's own lines can run across columns or not). */
function rowsOf(words: OcrWord[]): Row[] {
  const tolerance = median(words.map((w) => w.y1 - w.y0)) * 0.55;
  const rows: Row[] = [];
  for (const w of [...words].sort((a, b) => (a.y0 + a.y1) / 2 - (b.y0 + b.y1) / 2)) {
    const y = (w.y0 + w.y1) / 2;
    const row = rows.length ? rows[rows.length - 1] : null;
    if (row && Math.abs(y - row.y) <= tolerance) {
      row.words.push(w);
      row.y = (row.y * (row.words.length - 1) + y) / row.words.length;
    } else rows.push({ words: [w], y });
  }
  for (const row of rows) row.words.sort((a, b) => a.x0 - b.x0);
  return rows;
}

/**
 * A number as printed with two decimals ("+20.50", "-0.33"), allowing for what OCR does to them:
 * a lost decimal point ("+2050", "-046"), a comma, a "+" read as "4" ("421.00") and stray marks.
 */
export function readNumber(token: string): number | null {
  const t = token.replace(/[^\d.,+\-−–]/g, '').replace(/[−–]/g, '-').replace(',', '.');
  let m = t.match(/^([+-]?)(\d{1,2})\.(\d{2})$/);
  if (!m) m = t.match(/^([+-]?)(\d{1,2})(\d{2})$/);
  if (!m) {
    const plus = t.match(/^4(\d{2})\.(\d{2})$/);
    if (plus && Number(plus[1]) >= 10) m = ['', '+', plus[1], plus[2]] as unknown as RegExpMatchArray;
  }
  if (!m) return null;
  const value = Number(`${m[2]}.${m[3]}`);
  return m[1] === '-' ? -value : value;
}

/** An axis in degrees ("3°", "167°"; OCR may give "3*" or "2\""). */
function readAxis(token: string): number | null {
  const m = token.match(/^(\d{1,3})\D?$/);
  const n = m ? Number(m[1]) : NaN;
  return n >= 0 && n <= 180 ? n : null;
}

/** Strips the column rules OCR picks up as characters ("]", "|", "[") from a row's tokens. */
function tokens(words: OcrWord[]): string[] {
  return text(words)
    .split(/\s+/)
    .map((t) => t.replace(/^[\][|{}()]+|[\][|{}()]+$/g, ''))
    .filter(Boolean);
}

function standardRow(words: OcrWord[]): LensRow | null {
  const numbers = tokens(words).map(readNumber).filter((n): n is number => n !== null);
  if (numbers.length < 2) return null;
  const [power, ref] = numbers;
  return power > -15 && power < 45 && Math.abs(ref) < 15 ? { model: '', power, ref, words } : null;
}

/** "100 +20.50 +1.00 3° -0.33 -0.46 +0.27 3°": model, IOL SE, cyl, axis, then Ref SE, sph, cyl, axis. */
function toricRow(words: OcrWord[]): LensRow | null {
  const t = tokens(words);
  const start = t.findIndex((token) => /^[+-]/.test(token.replace(/[−–]/g, '-')) && readNumber(token) !== null);
  if (start === -1 || t.length < start + 4) return null;
  const before = t[start - 1] ?? '';
  const model = /^([A-Z]{2,4}|\d{3})$/i.test(before) ? before.toUpperCase() : '';
  const power = readNumber(t[start]);
  const cyl = readNumber(t[start + 1]);
  const axis = readAxis(t[start + 2]);
  const ref = readNumber(t[start + 3]);
  const refCyl = t[start + 5] !== undefined ? readNumber(t[start + 5]) : null;
  if (power === null || cyl === null || ref === null || power <= -15 || power >= 45) return null;
  return { model, power, ref, cyl, axis: axis ?? undefined, refCyl: refCyl ?? undefined, words };
}

function readEye(rows: Row[], left: number, right: number): EyeReading {
  const inside = (w: OcrWord) => centre(w) >= left && centre(w) < right;
  const half = rows.map((r) => ({ y: r.y, words: r.words.filter(inside) })).filter((r) => r.words.length);
  let target: number | null = null;
  for (const r of half) {
    const m = text(r.words).match(/target\s*re[a-z]?\.?\s*:?\s*([+\-−–]?\s?\d+[.,]\d+)/i);
    if (m) {
      const value = Number(m[1].replace(/\s/g, '').replace(/[−–]/g, '-').replace(',', '.'));
      if (Number.isFinite(value)) target = value;
      break;
    }
  }

  // "AL: 23.53 mm" and "SE: 44.73 D" (not TSE, the total keratometry), allowing for a lost decimal point.
  const measure = (label: RegExp, min: number, max: number) => {
    for (const r of half) {
      const m = text(r.words).match(label);
      const value = m ? readNumber(m[1]) : null;
      if (value !== null && value >= min && value <= max) return value;
    }
    return undefined;
  };
  const axialLength = measure(/\bAL\s*[:.]?\s*(\d{2}[.,]?\d{2})\s*mm/i, 15, 40);
  const k = measure(/(?<![A-Z])SE\s*[:.]?\s*(\d{2}[.,]?\d{2})\s*D/, 30, 60);

  const tables: LensTable[] = [];
  half.forEach((r, i) => {
    const isBarrett = (w: OcrWord) => /^barr?ett$/i.test(w.text.replace(/[^A-Za-z]/g, ''));
    if (!r.words.some(isBarrett)) return;
    const toric = /toric/i.test(text(r.words));
    // The formula line ("- Barrett Universal II TK -") is centred over each lens's table, so it marks the tables' columns.
    const height = median(r.words.map((w) => w.y1 - w.y0));
    const phrases: OcrWord[][] = [];
    for (const w of r.words) {
      const last = phrases[phrases.length - 1];
      if (last && w.x0 - last[last.length - 1].x1 < height * 1.5) last.push(w);
      else phrases.push([w]);
    }
    const centres = phrases.filter((p) => p.some(isBarrett)).map((p) => (p[0].x0 + p[p.length - 1].x1) / 2);
    centres.forEach((c, k) => {
      const from = k === 0 ? left : (centres[k - 1] + c) / 2;
      const to = k === centres.length - 1 ? right : (c + centres[k + 1]) / 2;
      const within = (row: Row) => row.words.filter((w) => centre(w) >= from && centre(w) < to);
      // The lens's name is on the line or two above.
      const title: string[] = [];
      const titleWords: OcrWord[] = [];
      for (let j = i - 1; j >= Math.max(0, i - 3); j--) {
        const words = within(half[j]).filter((w) => !/^(TK|[A-Z]?\W*)$/.test(w.text) || /&/.test(w.text));
        const line = text(words);
        if (!line || /emmetropia|\d[.,]\d{2}/i.test(line) || /^(IOL|Ref|LF)/i.test(line)) break;
        title.unshift(line);
        titleWords.push(...words);
        if (title.length === 2 || /johnson|alcon|zeiss|bausch|hoya|rayner|vision|medical/i.test(line)) break;
      }
      const tableRows: LensRow[] = [];
      for (let j = i + 1; j < Math.min(half.length, i + 12); j++) {
        const words = within(half[j]);
        const line = text(words);
        if (/emmetropia|barr?ett/i.test(line)) break;
        if (!line || /\b(LF|DF|IOL|Ref)\b/i.test(line)) continue;
        const row = toric ? toricRow(words) : standardRow(words);
        if (row) tableRows.push(row);
      }
      if (tableRows.length) tables.push({ title: title.join(' ').replace(/(\w)\s+-\s+(\w)/g, '$1-$2'), toric, rows: tableRows, titleWords });
    });
  });
  // ΔK is the start of the line with K2 on it (OCR reads "Δ" all sorts of ways).
  let deltaKWords: OcrWord[] | undefined;
  for (const r of half) {
    const at = r.words.findIndex((w) => /^[KkX][2zZ]\b/.test(w.text) || /^[KkX][2zZ][:.]/.test(w.text));
    if (at > 0 && /@/.test(text(r.words.slice(0, at)))) {
      deltaKWords = r.words.slice(0, at);
      break;
    }
  }
  return { target, axialLength, k, tables, deltaKWords };
}

/** The report's patient, warnings and each eye's lens tables, or null if it isn't an IOL calculation. */
export function readReport(words: OcrWord[]): ReportReading | null {
  const rows = rowsOf(words.filter((w) => w.text.trim()));
  // OD and OS head the two eyes' halves, in large type on the same line (OCR may read them as "0D", "O5"…).
  const eyeWord = (w: OcrWord) => w.text.replace(/[^A-Za-z0-9]/g, '').toUpperCase().replace(/[0Q]/g, 'O').replace(/5/g, 'S');
  let headings: { row: number; od?: OcrWord; os?: OcrWord; middle: number } | null = null;
  rows.forEach((r, i) => {
    const od = r.words.find((w) => eyeWord(w) === 'OD');
    const os = r.words.find((w) => eyeWord(w) === 'OS');
    if (od && os && od.x1 < os.x0 && !headings) headings = { row: i, od, os, middle: (od.x1 + os.x0) / 2 };
  });
  if (!headings) {
    // Without them, the formula lines ("- Barrett … -"), centred over each lens's table, split the eyes:
    // four tables (two an eye) or a toric calculation each.
    const isBarrett = (w: OcrWord) => /^barr?ett$/i.test(w.text.replace(/[^A-Za-z]/g, ''));
    const at = rows.findIndex((r) => r.words.filter(isBarrett).length >= 2);
    if (at === -1) return null;
    const centres = rows[at].words.filter(isBarrett).map(centre).sort((a, b) => a - b);
    const middle = (centres[Math.floor(centres.length / 2) - 1] + centres[Math.floor(centres.length / 2)]) / 2;
    // The calculation starts at its title, or else a few lines above the first lens heading.
    const title = rows.findIndex((r) => /IOL\s*calculation|eye\s*status/i.test(text(r.words)));
    headings = { row: title !== -1 && title < at ? title : Math.max(0, at - 18), middle };
  }
  const { row: headingRow, od, os, middle } = headings as { row: number; od?: OcrWord; os?: OcrWord; middle: number };

  const header = rows.slice(0, headingRow);
  const after = (label: RegExp, row: Row) => {
    const at = row.words.findIndex((_w, k) => label.test(text(row.words.slice(0, k + 1))));
    return at === -1 ? [] : row.words.slice(at + 1).filter((w) => centre(w) < middle * 0.9);
  };
  let nameWords: OcrWord[] = [];
  let idWords: OcrWord[] = [];
  let dobWords: OcrWord[] = [];
  for (const r of header) {
    const line = text(r.words);
    if (/date\s*of\s*birth/i.test(line)) dobWords = after(/birth$/i, r).filter((w) => /\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/.test(w.text)).slice(0, 1);
    if (/patient\s*id\b/i.test(line)) idWords = after(/patient\s*id$/i, r);
    else if (/^patient\b/i.test(line) && !nameWords.length) nameWords = after(/^patient$/i, r);
  }
  const name = text(nameWords).trim();
  const id = text(idWords).trim();
  // Warnings sit in the box between the measurement details and the calculation.
  const measured = header.findIndex((r) => /date\s*of\s*measurement/i.test(text(r.words)));
  const warnings = header
    .slice(measured + 1)
    .map((r) => text(r.words).replace(/^[\W_]+/, '').trim())
    .filter((line) => (line.match(/[A-Za-z]{3,}/g) ?? []).length >= 3 && !/IOL calculation|Physician|Operator/i.test(line));

  const body = rows.slice(headingRow + 1);
  const minX = Math.min(...words.map((w) => w.x0));
  const maxX = Math.max(...words.map((w) => w.x1));
  const right = readEye(body, minX - 1, middle);
  const left = readEye(body, middle, maxX + 1);
  // "OD" / "OS" with "right" / "left" under them.
  const under = (big: OcrWord, word: RegExp) =>
    body.slice(0, 2).flatMap((r) => r.words).find((w) => word.test(w.text) && w.x0 < big.x1 + (big.x1 - big.x0) && w.x1 > big.x0 - (big.x1 - big.x0));
  if (od) right.headingWords = [od, under(od, /^r[il1]ght$/i)].filter((w): w is OcrWord => Boolean(w));
  if (os) left.headingWords = [os, under(os, /^[l1I]eft$/i)].filter((w): w is OcrWord => Boolean(w));
  return {
    name,
    id: id.replace(/[^A-Za-z0-9]/g, ''),
    nameWords: nameWords.length ? nameWords : undefined,
    idWords: idWords.length ? idWords : undefined,
    dobWords: dobWords.length ? dobWords : undefined,
    warnings,
    right: right.tables.length || right.target !== null ? right : undefined,
    left: left.tables.length || left.target !== null ? left : undefined,
  };
}

// ---- Checking a lens against the report ----------------------------------------------

export interface PickedLens {
  /** RIGHT, LEFT, BOTH or '' (not set). */
  eye: string;
  /** e.g. "DCB00", "ZCU150", "AR40". */
  code: string;
  /** For a lens with a number (like "ZCUx"), the part before it and the number: "ZCU", "150". */
  family?: { base: string; number: string };
  power: number | null;
  /** Target refraction you picked (plano is 0), or null if none. */
  target: number | null;
  /** The patient as on your list, to compare with the report's. */
  mrn: string;
  surname: string;
}

export type CheckLevel = 'ok' | 'warn' | 'error';

/** A lens's A-constant (SRK/T), e.g. { pattern: 'CNA0Tx', value: 119.1 }; "x" stands for a typed number. */
export interface LensConstant {
  pattern: string;
  value: number;
}

/** "DCB00 119.3" per line, as in Settings. */
export function parseConstants(lines: string[]): LensConstant[] {
  return lines.flatMap((line) => {
    const m = line.trim().match(/^(\S+)\s*[:=]?\s*(\d{2,3}(?:[.,]\d+)?)$/);
    return m ? [{ pattern: m[1], value: Number(m[2].replace(',', '.')) }] : [];
  });
}

/** For a lens that isn't on the report: its numbers worked out from one that is. */
export interface LensEstimate {
  side: 'right' | 'left';
  /** The lens on the report it's worked out from, e.g. "DCB00". */
  reference: string;
  /** Power change needed for the same result (negative: a lower power). */
  shift: number;
  /** The working, in words, to check at a glance. */
  working: string[];
  /** Estimated predicted refraction for each power of the lens. */
  table: { power: number; ref: number }[];
  /** The report's rows for the lens it's worked out from. */
  referenceRows: LensRow[];
}

/** A box to draw on the screenshot: red for the right eye, blue for the left, purple for both or unknown. */
export interface Highlight {
  eye: 'right' | 'left' | 'both';
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  /** For something looked at but not a match: the reference lens for an estimate, or the other eye's power. */
  dashed?: boolean;
}

function box(words: OcrWord[] | undefined, eye: Highlight['eye'], dashed = false): Highlight[] {
  if (!words?.length) return [];
  return [{ eye, x0: Math.min(...words.map((w) => w.x0)), y0: Math.min(...words.map((w) => w.y0)), x1: Math.max(...words.map((w) => w.x1)), y1: Math.max(...words.map((w) => w.y1)), dashed: dashed || undefined }];
}

export interface LensCheck {
  /** Worst of the flags; 'ok' only when the lens was found on the report for the right eye. */
  level: CheckLevel | 'none';
  /** What the report says for your lens, e.g. "Right eye: DCB00 +20.50 → predicted −0.33". */
  facts: string[];
  flags: { level: CheckLevel; text: string }[];
  notes: string[];
  estimate?: LensEstimate;
  /** What to mark on the screenshot. */
  highlights: Highlight[];
}

const EYE_NAME: Record<'right' | 'left', string> = { right: 'right', left: 'left' };

/** Upper case without spaces or punctuation, with the letter O read as zero (OCR mixes them up in codes). */
function canonical(s: string): string {
  return s.toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/O/g, '0');
}

export function signed(n: number): string {
  return `${n < 0 ? '−' : '+'}${Math.abs(n).toFixed(2)}`;
}

/** Several screenshots of one patient's report read as one (e.g. one per eye, or two pages). */
export function combineReadings(readings: ReportReading[]): ReportReading | null {
  if (!readings.length) return null;
  const eye = (side: 'right' | 'left'): EyeReading | undefined => {
    const all = readings.map((r) => r[side]).filter((e): e is EyeReading => Boolean(e));
    if (!all.length) return undefined;
    return {
      target: all.find((e) => e.target !== null)?.target ?? null,
      axialLength: all.find((e) => e.axialLength !== undefined)?.axialLength,
      k: all.find((e) => e.k !== undefined)?.k,
      tables: all.flatMap((e) => e.tables),
      headingWords: all.find((e) => e.headingWords?.length)?.headingWords,
      deltaKWords: all.find((e) => e.deltaKWords?.length)?.deltaKWords,
    };
  };
  return {
    name: readings.find((r) => r.name)?.name ?? '',
    id: readings.find((r) => r.id)?.id ?? '',
    nameWords: readings.find((r) => r.nameWords)?.nameWords,
    idWords: readings.find((r) => r.idWords)?.idWords,
    dobWords: readings.find((r) => r.dobWords)?.dobWords,
    warnings: [...new Set(readings.flatMap((r) => r.warnings))],
    right: eye('right'),
    left: eye('left'),
  };
}

function editDistance(a: string, b: string): number {
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array<number>(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++)
    for (let j = 1; j <= b.length; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return d[a.length][b.length];
}

function tablesFor(eye: EyeReading | undefined, lens: PickedLens): LensTable[] {
  if (!eye) return [];
  const code = canonical(lens.family?.base ?? lens.code);
  const named = eye.tables.filter((t) => canonical(t.title).includes(code));
  // A lens with a number (a toric) is read from the toric calculation; others from their own table first.
  if (lens.family) return named.filter((t) => t.toric).concat(named.filter((t) => !t.toric));
  const plain = named.filter((t) => !t.toric);
  if (plain.length) return plain;
  // A plain lens recommended in the toric calculation is listed there by its first letters (e.g. "DCB").
  return eye.tables.filter((t) => t.toric && t.rows.some((r) => r.model && code.startsWith(canonical(r.model))));
}

function rowFor(tables: LensTable[], lens: PickedLens): { table: LensTable; row: LensRow } | null {
  if (lens.power === null) return null;
  for (const table of tables) {
    const row = table.rows.find((r) => {
      if (Math.abs(r.power - lens.power!) > 0.01) return false;
      if (!table.toric || !r.model) return true;
      if (lens.family) return canonical(r.model) === canonical(lens.family.number);
      return canonical(lens.code).startsWith(canonical(r.model));
    });
    if (row) return { table, row };
  }
  return null;
}


function constantFor(code: string, constants: LensConstant[]): LensConstant | null {
  const c = canonical(code);
  return (
    constants.find((k) => canonical(k.pattern) === c) ??
    constants.find((k) => {
      const at = k.pattern.indexOf('x');
      if (at === -1) return false;
      return new RegExp(`^${canonical(k.pattern.slice(0, at))}[A-Z0-9]*${canonical(k.pattern.slice(at + 1))}$`).test(c);
    }) ??
    null
  );
}

/** Predicted refraction at a power, between the table's rows (or beyond its ends, along its slope). */
function refractionAt(rows: LensRow[], power: number): { ref: number; low: LensRow; high: LensRow } {
  const sorted = [...rows].sort((a, b) => a.power - b.power);
  let i = sorted.findIndex((r) => r.power >= power);
  if (i <= 0) i = 1;
  const low = sorted[Math.min(i, sorted.length - 1) - 1];
  const high = sorted[Math.min(i, sorted.length - 1)];
  const ref = high.power === low.power ? low.ref : low.ref + ((power - low.power) * (high.ref - low.ref)) / (high.power - low.power);
  return { ref, low, high };
}

const dioptres = (n: number) => `${Math.abs(n).toFixed(2)} D`;
const round2 = (n: number) => Math.round(n * 100) / 100;

/**
 * A lens that isn't on the report, worked out from one that is (on this eye, with an A-constant):
 * the report's Barrett numbers for that lens, moved by the power difference between the two lenses'
 * A-constants in this eye (SRK/T with the report's AL and K; or the A-constant difference itself if
 * those couldn't be read).
 */
export function estimateLens(
  eye: EyeReading,
  side: 'right' | 'left',
  lens: PickedLens,
  constants: LensConstant[],
  target = eye.target ?? lens.target ?? 0,
): LensEstimate | null {
  const own = constantFor(lens.code, constants);
  if (!own) return null;
  let reference: { table: LensTable; code: string; a: number } | null = null;
  for (const table of eye.tables) {
    if (table.toric || table.rows.length < 2) continue;
    const k = constants.find((c) => {
      const base = c.pattern.includes('x') ? c.pattern.slice(0, c.pattern.indexOf('x')) : c.pattern;
      return base.length >= 3 && canonical(table.title).includes(canonical(base)) && !c.pattern.includes('x');
    });
    if (k) {
      reference = { table, code: k.pattern, a: k.value };
      break;
    }
  }
  if (!reference) return null;
  const rows = reference.table.rows;
  const measured = eye.axialLength !== undefined && eye.k !== undefined;
  const shift = round2(
    measured ? srktPower(eye.axialLength!, eye.k!, own.value, target) - srktPower(eye.axialLength!, eye.k!, reference.a, target) : own.value - reference.a,
  );
  const ref = reference.code;
  const code = lens.code;
  const working = [
    `${code} isn’t on the report, so this is worked out from the ${side} eye’s ${ref} numbers (Barrett).`,
    `A-constants used (SRK/T, from Settings): ${code} ${own.value}${own.pattern !== code ? ` (as ${own.pattern})` : ''}, ${ref} ${reference.a}.`,
  ];
  if (Math.abs(shift) < 0.005) {
    working.push(`${code} and ${ref} have the same A-constant (${own.value}), so ${ref}’s numbers apply as they are.`);
  } else {
    const way = shift < 0 ? 'lower' : 'higher';
    working.push(`${code} has a ${way} A-constant than ${ref} (${own.value} vs ${reference.a}), so it needs a ${way} power for the same result.`);
    working.push(
      measured
        ? `SRK/T for this eye (AL ${eye.axialLength!.toFixed(2)} mm, K ${eye.k!.toFixed(2)} D, target ${signed(target)}): ${ref} (A ${reference.a}) ${signed(srktPower(eye.axialLength!, eye.k!, reference.a, target))}, ${code} (A ${own.value}) ${signed(srktPower(eye.axialLength!, eye.k!, own.value, target))}, so ${code} needs ${dioptres(shift)} ${way}.`
        : `The eye’s AL and K couldn’t be read from the report, so the difference is taken as the A-constants’ difference: ${own.value} − ${reference.a} = ${shift > 0 ? '+' : '−'}${dioptres(shift)}.`,
    );
    const at = lens.power ?? rows.reduce((best, r) => (Math.abs(r.ref - target) < Math.abs(best.ref - target) ? r : best)).power;
    const same = refractionAt(rows, at - shift).ref - refractionAt(rows, at).ref;
    working.push(`Keeping the same power with ${code} gives a more ${same < 0 ? 'myopic' : 'hyperopic'} result: about ${dioptres(same)}.`);
  }
  const powers = rows.map((r) => r.power);
  const table: LensEstimate['table'] = [];
  // Highest power first, as on the report.
  for (let p = Math.floor((Math.max(...powers) + shift) * 2) / 2; p >= Math.min(...powers) + shift - 1e-9; p -= 0.5) {
    table.push({ power: p, ref: round2(refractionAt(rows, p - shift).ref) });
  }
  if (lens.power !== null) {
    const equivalent = round2(lens.power - shift);
    const { ref: predicted, low, high } = refractionAt(rows, equivalent);
    const outside = equivalent < Math.min(...powers) - 0.5 || equivalent > Math.max(...powers) + 0.5;
    working.push(
      `So ${code} ${signed(lens.power)} is like ${ref} ${signed(equivalent)} → predicted about ${signed(round2(predicted))}` +
        (outside ? `, well outside the report’s ${ref} powers, so less certain.` : ` (${ref} ${signed(low.power)} → ${signed(low.ref)}, ${signed(high.power)} → ${signed(high.ref)}).`),
    );
  }
  if (table.length) {
    const best = table.reduce((a, b) => (Math.abs(b.ref - target) < Math.abs(a.ref - target) ? b : a));
    working.push(`For a target of ${signed(target)}: ${code} ${signed(best.power)} (about ${signed(best.ref)}).`);
  }
  if (lens.family && lens.family.number.replace(/0/g, '') !== '') working.push('This is the spherical equivalent only: check the cylinder and axis with a toric calculator.');
  return { side, reference: ref, shift, working, table, referenceRows: rows };
}


// ---- Trying another lens ------------------------------------------------------------

export interface SimRow {
  power: number;
  ref: number;
  /** Beyond the powers on the report, so read along the table's slope: less certain. */
  outside?: boolean;
  cyl?: number;
  axis?: number;
  refCyl?: number;
}

export interface Simulation {
  side: 'right' | 'left';
  code: string;
  target: number;
  source: 'report' | 'estimate';
  /** For an estimate: the lens on the report it's worked out from, and the working. */
  reference?: string;
  working: string[];
  notes: string[];
  /** Powers around the best one for the target (and the one chosen). */
  rows: SimRow[];
  best: SimRow;
  selected?: SimRow;
}

/**
 * Any lens and target against one eye of the report: straight from the report's numbers when the lens
 * is on it (its own table, or the toric calculation for a toric), otherwise estimated from a lens that
 * is (see estimateLens). Returns a reason instead when it can't be worked out.
 */
export function simulateLens(reading: ReportReading, side: 'right' | 'left', lens: PickedLens, constants: LensConstant[]): Simulation | string {
  const eye = reading[side];
  if (!eye) return `Couldn’t read the ${side} eye’s side of the report.`;
  if (!lens.code) return 'Choose a lens.';
  if (lens.family && !lens.family.number) return `Choose the number for ${lens.family.base}… (e.g. the toric model).`;
  const target = lens.target ?? eye.target ?? 0;
  const notes: string[] = [];
  let working: string[] = [];
  let reference: string | undefined;
  let source: Simulation['source'] = 'report';
  let base: LensRow[] = [];
  let shift = 0;

  const tables = tablesFor(eye, lens);
  const toric = lens.family ? tables.find((t) => t.toric) : undefined;
  if (toric) {
    const number = canonical(lens.family!.number);
    const models = toric.rows.map((r) => r.model).filter(Boolean);
    const recommended = models.sort((a, b) => models.filter((m) => m === b).length - models.filter((m) => m === a).length)[0] ?? '';
    const axis = toric.rows.find((r) => r.axis !== undefined)?.axis;
    const matching = toric.rows.filter((r) => !r.model || canonical(r.model) === number);
    base = matching.length >= 2 ? matching : toric.rows;
    if (recommended && canonical(recommended) !== number) {
      const what = /^[A-Z]+$/.test(recommended) ? 'a non-toric lens' : `${lens.family!.base}${recommended}`;
      notes.push(
        `The report’s toric calculation (Barrett Toric) chose ${what} for this eye${axis !== undefined ? ` at ${axis}°` : ''}. ` +
          `The spherical equivalents below still apply to ${lens.code}, but its axis and residual astigmatism aren’t calculated here: check them with a toric calculator.`,
      );
      base = toric.rows.map((r) => ({ ...r, cyl: undefined, axis: undefined, refCyl: undefined }));
    } else {
      notes.push(`From the report’s toric calculation (Barrett Toric) for ${lens.code}${axis !== undefined ? `: axis ${axis}°` : ''}.`);
    }
  } else if (tables.some((t) => !t.toric && t.rows.length >= 2)) {
    base = tables.find((t) => !t.toric && t.rows.length >= 2)!.rows;
  } else {
    const estimate = estimateLens(eye, side, lens, constants, target);
    if (!estimate) {
      return constantFor(lens.code, constants)
        ? `${lens.code} isn’t on the ${side} eye’s side of the report, and there’s no lens on it with an A-constant in Settings to work it out from.`
        : `${lens.code} isn’t on the ${side} eye’s side of the report. Add its A-constant in Settings to estimate it from a lens that is.`;
    }
    source = 'estimate';
    reference = estimate.reference;
    working = estimate.working;
    base = estimate.referenceRows;
    shift = estimate.shift;
    if (lens.family && /[1-9]/.test(lens.family.number)) notes.push('Spherical equivalent only: choose the cylinder and axis with a toric calculator.');
  }

  const powers = base.map((r) => r.power + shift);
  const low = Math.min(...powers);
  const high = Math.max(...powers);
  const row = (power: number): SimRow => {
    const exact = source === 'report' ? base.find((r) => Math.abs(r.power - power) < 0.01) : undefined;
    return {
      power,
      ref: round2(exact ? exact.ref : refractionAt(base, power - shift).ref),
      outside: power < low - 0.25 || power > high + 0.25 || undefined,
      cyl: exact?.cyl,
      axis: exact?.axis,
      refCyl: exact?.refCyl,
    };
  };
  // The best power is the half-dioptre step nearest the target (the more myopic of two equally near).
  let best = row(Math.round(low * 2) / 2);
  for (let p = Math.floor((low - 4) * 2) / 2; p <= high + 4; p += 0.5) {
    const candidate = row(p);
    const gap = Math.abs(candidate.ref - target) - Math.abs(best.ref - target);
    if (gap < -0.001 || (Math.abs(gap) <= 0.001 && candidate.ref < best.ref)) best = candidate;
  }
  const shown = new Set([-1, -0.5, 0, 0.5, 1].map((d) => best.power + d));
  const selected = lens.power !== null ? row(lens.power) : undefined;
  if (selected) shown.add(selected.power);
  const rows = [...shown].sort((a, b) => b - a).map(row);
  if (rows.some((r) => r.outside)) notes.push('Powers marked ≈ are beyond those on the report, read along its slope: less certain.');
  return { side, code: lens.code, target, source, reference, working, notes, rows, best, selected };
}

/** Your lens checked against the report: found for the eye you're operating on, or what's wrong. */
export function checkLens(reading: ReportReading, lens: PickedLens, constants: LensConstant[] = []): LensCheck {
  const facts: string[] = [];
  const flags: LensCheck['flags'] = [];
  const notes = [...reading.warnings];
  let estimate: LensEstimate | undefined;
  const operated: Highlight['eye'] = lens.eye === 'RIGHT' ? 'right' : lens.eye === 'LEFT' ? 'left' : 'both';
  const highlights: Highlight[] = [...box(reading.nameWords, operated), ...box(reading.dobWords, operated), ...box(reading.idWords, operated)];

  // Is it this patient's report?
  const reportId = reading.id.replace(/\D/g, '').replace(/^0+/, '');
  const ourId = lens.mrn.replace(/\D/g, '').replace(/^0+/, '');
  if (reportId.length >= 4 && ourId.length >= 4 && reportId !== ourId) {
    flags.push({ level: 'error', text: `The report is for patient ID ${reading.id}, not ${lens.mrn}.` });
  }
  const surname = lens.surname.toUpperCase().replace(/[^A-Z]/g, '');
  const reportWords = reading.name.toUpperCase().split(/[^A-Z]+/).filter((w) => w.length > 1);
  if (surname.length >= 2 && reportWords.length && !reportWords.some((w) => editDistance(w, surname) <= (surname.length > 4 ? 1 : 0))) {
    flags.push({ level: 'error', text: `The report’s name (${reading.name}) doesn’t match ${lens.surname}.` });
  }

  const sides: ('right' | 'left')[] = lens.eye === 'RIGHT' ? ['right'] : lens.eye === 'LEFT' ? ['left'] : ['right', 'left'];
  // The eye's heading and ΔK are marked whether or not a lens has been picked.
  for (const side of sides) highlights.push(...box(reading[side]?.headingWords, side), ...box(reading[side]?.deltaKWords, side));
  if (!lens.eye) flags.push({ level: 'warn', text: 'The eye isn’t set, so both eyes on the report were checked.' });
  if (!lens.code || lens.power === null) {
    flags.push({ level: 'warn', text: 'No lens and power picked yet to check against the report.' });
  } else {
    let found = false;
    for (const side of sides) {
      const eye = reading[side];
      if (!eye) {
        flags.push({ level: 'warn', text: `Couldn’t read the ${EYE_NAME[side]} eye’s side of the report.` });
        continue;
      }
      const tables = tablesFor(eye, lens);
      const hit = rowFor(tables, lens);
      if (hit) {
        found = true;
        highlights.push(...box(hit.table.titleWords, side), ...box(hit.row.words, side));
        const { row } = hit;
        const toric = row.cyl !== undefined && row.axis !== undefined && row.cyl > 0 ? `, axis ${row.axis}°` : '';
        facts.push(`${side === 'right' ? 'Right' : 'Left'} eye: ${lens.code} ${signed(row.power)}${toric} → predicted ${signed(row.ref)}`);
        continue;
      }
      const other = side === 'right' ? 'left' : 'right';
      const elsewhere = sides.length === 1 ? rowFor(tablesFor(reading[other], lens), lens) : null;
      if (elsewhere) {
        highlights.push(...box(elsewhere.table.titleWords, other, true), ...box(elsewhere.row.words, other, true));
        flags.push({ level: 'error', text: `${signed(lens.power)} is in the ${other.toUpperCase()} eye’s column of the report, not the ${side}’s.` });
      } else if (tables.length) {
        const powers = tables.flatMap((t) => t.rows.map((r) => r.power));
        flags.push({
          level: 'warn',
          text: `${lens.code} ${signed(lens.power)} isn’t among the ${side} eye’s powers on the report (${signed(Math.min(...powers))} to ${signed(Math.max(...powers))}).`,
        });
      } else if (eye.tables.length) {
        const worked = estimateLens(eye, side, lens, constants);
        if (worked) {
          found = true;
          estimate ??= worked;
          const reference = eye.tables.find((t) => !t.toric && canonical(t.title).includes(canonical(worked.reference)));
          if (reference && reference.rows.length >= 2) {
            const near = refractionAt(reference.rows, lens.power - worked.shift);
            highlights.push(...box(reference.titleWords, side, true), ...box(near.low.words, side, true), ...box(near.high.words, side, true));
          }
          const predicted = signed(round2(refractionAt(tablesOf(eye, worked.reference), lens.power - worked.shift).ref));
          facts.push(`${side === 'right' ? 'Right' : 'Left'} eye (estimated): ${lens.code} ${signed(lens.power)} → about ${predicted}`);
          flags.push({ level: 'warn', text: `${lens.code} isn’t on the report: estimated from ${worked.reference}. Check the working.` });
        } else {
          flags.push({ level: 'warn', text: `${lens.code} isn’t on the ${side} eye’s side of the report${constantFor(lens.code, constants) ? '' : ' (add its A-constant in Settings to estimate it)'}.` });
        }
      } else {
        flags.push({ level: 'warn', text: `Couldn’t read the lens tables for the ${side} eye.` });
      }
    }
    if (found && lens.target !== null) {
      for (const side of sides) {
        const t = reading[side]?.target;
        if (t !== null && t !== undefined && Math.abs(t - lens.target) > 0.01) {
          flags.push({ level: 'warn', text: `The report was calculated for a target of ${signed(t)} D (${side} eye); you picked ${lens.target === 0 ? 'plano' : signed(lens.target)}.` });
        }
      }
    }
  }
  const level: LensCheck['level'] = flags.some((f) => f.level === 'error') ? 'error' : flags.length ? 'warn' : facts.length ? 'ok' : 'none';
  return { level, facts, flags, notes, estimate, highlights };
}

function tablesOf(eye: EyeReading, code: string): LensRow[] {
  return eye.tables.find((t) => !t.toric && canonical(t.title).includes(canonical(code)))?.rows ?? [];
}

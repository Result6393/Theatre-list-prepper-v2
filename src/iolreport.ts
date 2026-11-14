/**
 * Reading an IOL calculation report (laid out like the IOLMaster 700's) from OCR'd words, and checking
 * the lens you picked against it. Only ever a cross-check: nothing here changes what you picked.
 */

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
}

export interface LensTable {
  /** The lens heading as read, e.g. "Johnson & Johnson Vision Tecnis DCB00 Preload". */
  title: string;
  toric: boolean;
  rows: LensRow[];
}

export interface EyeReading {
  /** Target refraction the report was calculated for. */
  target: number | null;
  tables: LensTable[];
}

export interface ReportReading {
  /** As printed beside Patient and Patient ID, if they could be read. */
  name: string;
  id: string;
  /** The report's own warnings, e.g. a difference in axial length between the eyes. */
  warnings: string[];
  right?: EyeReading;
  left?: EyeReading;
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
  return power > -15 && power < 45 && Math.abs(ref) < 15 ? { model: '', power, ref } : null;
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
  if (power === null || cyl === null || ref === null || power <= -15 || power >= 45) return null;
  return { model, power, ref, cyl, axis: axis ?? undefined };
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
      for (let j = i - 1; j >= Math.max(0, i - 3); j--) {
        const words = within(half[j]).filter((w) => !/^(TK|[A-Z]?\W*)$/.test(w.text) || /&/.test(w.text));
        const line = text(words);
        if (!line || /emmetropia|\d[.,]\d{2}/i.test(line) || /^(IOL|Ref|LF)/i.test(line)) break;
        title.unshift(line);
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
      if (tableRows.length) tables.push({ title: title.join(' ').replace(/(\w)\s+-\s+(\w)/g, '$1-$2'), toric, rows: tableRows });
    });
  });
  return { target, tables };
}

/** The report's patient, warnings and each eye's lens tables, or null if it isn't an IOL calculation. */
export function readReport(words: OcrWord[]): ReportReading | null {
  const rows = rowsOf(words.filter((w) => w.text.trim()));
  // OD and OS head the two eyes' halves, in large type on the same line.
  let headings: { row: number; od: OcrWord; os: OcrWord } | null = null;
  rows.forEach((r, i) => {
    const od = r.words.find((w) => /^[O0Q]D$/i.test(w.text));
    const os = r.words.find((w) => /^[O0Q]S$/i.test(w.text));
    if (od && os && od.x1 < os.x0 && !headings) headings = { row: i, od, os };
  });
  if (!headings) return null;
  const { row: headingRow, od, os } = headings as { row: number; od: OcrWord; os: OcrWord };
  const middle = (od.x1 + os.x0) / 2;

  const header = rows.slice(0, headingRow);
  const after = (label: RegExp, row: Row) => {
    const at = row.words.findIndex((_w, k) => label.test(text(row.words.slice(0, k + 1))));
    return at === -1 ? '' : text(row.words.slice(at + 1).filter((w) => centre(w) < middle * 0.9)).trim();
  };
  let name = '';
  let id = '';
  for (const r of header) {
    const line = text(r.words);
    if (/patient\s*id\b/i.test(line)) id = after(/patient\s*id$/i, r);
    else if (/^patient\b/i.test(line) && !name) name = after(/^patient$/i, r);
  }
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
  return {
    name,
    id: id.replace(/[^A-Za-z0-9]/g, ''),
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

export interface LensCheck {
  /** Worst of the flags; 'ok' only when the lens was found on the report for the right eye. */
  level: CheckLevel | 'none';
  /** What the report says for your lens, e.g. "Right eye: DCB00 +20.50 → predicted −0.33". */
  facts: string[];
  flags: { level: CheckLevel; text: string }[];
  notes: string[];
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
    return { target: all.find((e) => e.target !== null)?.target ?? null, tables: all.flatMap((e) => e.tables) };
  };
  return {
    name: readings.find((r) => r.name)?.name ?? '',
    id: readings.find((r) => r.id)?.id ?? '',
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

/** Your lens checked against the report: found for the eye you're operating on, or what's wrong. */
export function checkLens(reading: ReportReading, lens: PickedLens): LensCheck {
  const facts: string[] = [];
  const flags: LensCheck['flags'] = [];
  const notes = [...reading.warnings];

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
        const { row } = hit;
        const toric = row.cyl !== undefined && row.axis !== undefined && row.cyl > 0 ? `, axis ${row.axis}°` : '';
        facts.push(`${side === 'right' ? 'Right' : 'Left'} eye: ${lens.code} ${signed(row.power)}${toric} → predicted ${signed(row.ref)}`);
        continue;
      }
      const other = side === 'right' ? 'left' : 'right';
      const elsewhere = sides.length === 1 ? rowFor(tablesFor(reading[other], lens), lens) : null;
      if (elsewhere) {
        flags.push({ level: 'error', text: `${signed(lens.power)} is in the ${other.toUpperCase()} eye’s column of the report, not the ${side}’s.` });
      } else if (tables.length) {
        const powers = tables.flatMap((t) => t.rows.map((r) => r.power));
        flags.push({
          level: 'warn',
          text: `${lens.code} ${signed(lens.power)} isn’t among the ${side} eye’s powers on the report (${signed(Math.min(...powers))} to ${signed(Math.max(...powers))}).`,
        });
      } else if (eye.tables.length) {
        flags.push({ level: 'warn', text: `${lens.code} isn’t on the ${side} eye’s side of the report.` });
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
  return { level, facts, flags, notes };
}

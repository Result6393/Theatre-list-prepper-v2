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
  /** The rest of the biometry, for the alerts (see biometryAlerts). */
  biometry?: Biometry;
}

/** Biometry values for the alerts, each with where it is on the screenshot. */
export interface Biometry {
  acd?: Measured;
  lt?: Measured;
  wtw?: Measured;
  al?: Measured;
  k?: Measured;
  /** Corneal astigmatism (ΔK) and total keratometry astigmatism (ΔTK): cylinder and axis. */
  deltaK?: Measured & { axis?: number };
  deltaTK?: Measured & { axis?: number };
  /** "Phakic", "Pseudophakic" or "Aphakic". */
  lensStatus?: string;
  lensStatusWords?: OcrWord[];
  /** Laser vision correction: "Untreated", or a treatment. */
  lvc?: string;
  lvcWords?: OcrWord[];
  /** Lines with a value the IOLMaster marks as borderline "(!)". */
  borderline: { text: string; words: OcrWord[] }[];
}

export interface Measured {
  value: number;
  words: OcrWord[];
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

  /** A labelled value and the words it's in (from the label to the value), e.g. "ACD: 2.83 mm". */
  const labelled = (label: RegExp, value: RegExp, parse: (t: string) => number | null, min: number, max: number): Measured | undefined => {
    for (const r of half) {
      const at = r.words.findIndex((w) => label.test(w.text));
      if (at === -1) continue;
      for (let j = at; j < Math.min(r.words.length, at + 3); j++) {
        const m = text(r.words.slice(at, j + 1)).match(value);
        const n = m ? parse(m[1]) : null;
        if (n !== null && n >= min && n <= max) return { value: n, words: r.words.slice(at, j + 1) };
      }
    }
    return undefined;
  };
  // One decimal (WTW "11.9") or two ("2.83"), allowing for a lost decimal point.
  const oneOrTwo = (t: string) => {
    const c = t.replace(',', '.').replace(/[^\d.]/g, '');
    if (c.includes('.')) return Number(c);
    return c.length === 3 ? Number(`${c.slice(0, 2)}.${c.slice(2)}`) : c.length === 4 ? Number(`${c.slice(0, 2)}.${c.slice(2)}`) : null;
  };
  const decimal = (t: string) => {
    const c = t.replace(',', '.').replace(/[^\d.]/g, '');
    if (c.includes('.')) return Number(c);
    return c.length === 3 ? Number(`${c[0]}.${c.slice(1)}`) : null;
  };
  /** "ΔK: +0.47 D @ 6°" (OCR reads Δ as A, &, 4…): the cylinder and its axis. */
  const astigmatism = (label: RegExp) => {
    for (const r of half) {
      const at = r.words.findIndex((w) => label.test(w.text));
      if (at === -1) continue;
      const words: OcrWord[] = [];
      for (const w of r.words.slice(at)) {
        words.push(w);
        if (/\d\s*°?$/.test(w.text) && /@/.test(text(words))) break;
        if (words.length > 5) break;
      }
      const m = text(words).match(/([+-]?\d+[.,]?\d*)\s*D?\s*@\s*(\d{1,3})/);
      const cyl = m ? readNumber(m[1]) : null;
      if (cyl !== null && Math.abs(cyl) < 15) return { value: Math.abs(cyl), axis: Number(m![2]) % 181, words };
    }
    return undefined;
  };
  const statusOf = (label: RegExp, value: RegExp) => {
    for (const r of half) {
      const line = text(r.words);
      const m = line.match(value);
      if (m && label.test(line)) {
        const at = r.words.findIndex((w) => label.test(w.text));
        const end = r.words.findIndex((w, j) => j > at && new RegExp(m[1], 'i').test(w.text));
        return { text: m[1], words: r.words.slice(Math.max(0, at), end === -1 ? at + 2 : end + 1) };
      }
    }
    return undefined;
  };
  const lens = statusOf(/^LS\b/i, /\bLS\s*[:.]?\s*(Pseudophakic|Aphakic|Phakic)/i);
  const lvc = statusOf(/^LV[CG]\b/i, /\bLV[CG]\s*:\s*(?!mode)([A-Za-z][A-Za-z-]*)/i);
  const biometry: Biometry = {
    al: labelled(/^AL\b/i, /\bAL\s*[:.]?\s*(\d{2}[.,]?\d{2})/i, (t) => readNumber(t), 15, 40),
    k: labelled(/^SE\b/, /\bSE\s*[:.]?\s*(\d{2}[.,]?\d{2})/, (t) => readNumber(t), 30, 60),
    acd: labelled(/^A[CG]D\b/i, /\bA[CG]D\s*[:.]?\s*(\d[.,]?\d{2})/i, decimal, 1, 7),
    lt: labelled(/^LT\b/i, /\bLT\s*[:.]?\s*(\d[.,]?\d{2})/i, decimal, 0.2, 8),
    wtw: labelled(/^W[TW]{1,2}W?\b/i, /\bW[TW]{1,2}W?\s*[:.]?\s*(\d{2}[.,]?\d{1,2})/i, oneOrTwo, 8, 15),
    deltaK: astigmatism(/^[ΔA&4]K[:.]?$/i),
    deltaTK: astigmatism(/^[ΔA&4]TK[:.]?$/i),
    lensStatus: lens?.text.charAt(0).toUpperCase() + (lens?.text.slice(1).toLowerCase() ?? ''),
    lensStatusWords: lens?.words,
    lvc: lvc?.text,
    lvcWords: lvc?.words,
    // "(!)" after a value (OCR may drop the "!"), but not the legend at the foot of the page.
    borderline: half
      .filter((r) => r.words.some((w) => /^\((?:!|I|l|1|\|)?\)$/.test(w.text)) && !/borderline|edited/i.test(text(r.words)) && /\d/.test(text(r.words)))
      .map((r) => {
        const at = r.words.findIndex((w) => /^\((?:!|I|l|1|\|)?\)$/.test(w.text));
        const words = r.words.slice(0, at + 3);
        return { text: text(words).replace(/\s*\((?:!|I|l|1|\|)?\)\s*/g, ' (!) ').trim(), words };
      }),
  };
  if (!lens) {
    biometry.lensStatus = undefined;
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
  return { target, axialLength, k, tables, deltaKWords, biometry };
}

/** The report's patient, warnings and each eye's lens tables, or null if it isn't an IOL calculation. */
export function readReport(words: OcrWord[]): ReportReading | null {
  const rows = rowsOf(words.filter((w) => w.text.trim()));
  // OD and OS head the two eyes' halves, in large type on the same line (OCR may read them as "0D", "O5"…).
  // Large type only: "[OD]" in a warning ("[OD] Attention: …") isn't a heading.
  const typical = median(words.map((w) => w.y1 - w.y0));
  const eyeWord = (w: OcrWord) =>
    w.y1 - w.y0 > typical * 1.6 ? w.text.replace(/[.,:;]+$/, '').toUpperCase().replace(/[0Q]/g, 'O').replace(/5/g, 'S') : '';
  let headings: { row: number; od?: OcrWord; os?: OcrWord; middle: number } | null = null;
  rows.forEach((r, i) => {
    const od = r.words.find((w) => eyeWord(w) === 'OD');
    const os = r.words.find((w) => eyeWord(w) === 'OS');
    if (od && os && od.x1 < os.x0 && !headings) headings = { row: i, od, os, middle: (od.x1 + os.x0) / 2 };
  });
  if (!headings) {
    // Without them, the centred titles ("IOL calculation", "Eye status", "Biometric values") mark the
    // middle: lines of just those words, not a sentence that mentions them ("This IOL calculation contains…").
    // Failing that, the formula lines ("- Barrett … -") centred over the lens tables of both eyes.
    const titleWords = /^(IOL|I0L|calculation|eye|status|biometric|values|[O0Q]D|[O0Q]S)$/i;
    const titles = rows
      .map((r, i) => ({ r, i }))
      .filter(({ r }) => /IOL\s*calculation|eye\s*status|biometric\s*values/i.test(text(r.words)) && r.words.filter((w) => !titleWords.test(w.text.replace(/[^A-Za-z0-9]/g, ''))).length === 0);
    const isBarrett = (w: OcrWord) => /^barr?ett$/i.test(w.text.replace(/[^A-Za-z]/g, ''));
    const formulas = rows.findIndex((r) => r.words.filter(isBarrett).length >= 2);
    if (titles.length) {
      const middles = titles.map(({ r }) => {
        const t = r.words.filter((w) => /^(IOL|I0L|calculation|eye|status|biometric|values)$/i.test(w.text.replace(/[^A-Za-z0-9]/g, '')));
        return (Math.min(...t.map((w) => w.x0)) + Math.max(...t.map((w) => w.x1))) / 2;
      });
      headings = { row: titles[0].i, middle: median(middles) };
    } else if (formulas !== -1) {
      const centres = rows[formulas].words.filter(isBarrett).map(centre).sort((a, b) => a - b);
      headings = { row: Math.max(0, formulas - 18), middle: (centres[Math.floor(centres.length / 2) - 1] + centres[Math.floor(centres.length / 2)]) / 2 };
    } else return null;
  }
  const found = headings as { row: number; od?: OcrWord; os?: OcrWord; middle: number };
  // With one heading read, still mark it on the screenshot.
  if (!found.od || !found.os) {
    for (const w of rows[found.row]?.words ?? []) {
      if (eyeWord(w) === 'OD' && centre(w) < found.middle) found.od = w;
      if (eyeWord(w) === 'OS' && centre(w) > found.middle) found.os = w;
    }
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
    .map((r) => text(r.words).replace(/^[^\w[]+/, '').trim())
    .filter((line) => (line.match(/[A-Za-z]{3,}/g) ?? []).length >= 3 && !/^\W*IOL\s*calculation\W*$/i.test(line));

  const body = rows.slice(headingRow + 1);
  const minX = Math.min(...words.map((w) => w.x0));
  const maxX = Math.max(...words.map((w) => w.x1));
  const right = readEye(body, minX - 1, middle);
  const left = readEye(body, middle, maxX + 1);
  // "OD" / "OS" with "right" / "left" under them.
  const under = (big: OcrWord, word: RegExp) =>
    body.slice(0, 2).flatMap((r) => r.words).find((w) => word.test(w.text) && w.x0 < big.x1 + (big.x1 - big.x0) && w.x1 > big.x0 - (big.x1 - big.x0));
  // Without the big heading (OCR can garble it), the small "right" / "left" under it is marked.
  const small = (word: RegExp, side: 'right' | 'left') =>
    rows.slice(headingRow, headingRow + 3).flatMap((r) => r.words).find((w) => word.test(w.text) && (side === 'right' ? centre(w) < middle : centre(w) > middle));
  right.headingWords = (od ? [od, under(od, /^r[il1]ght$/i)] : [small(/^r[il1]ght$/i, 'right')]).filter((w): w is OcrWord => Boolean(w));
  left.headingWords = (os ? [os, under(os, /^[l1I]eft$/i)] : [small(/^[l1I]eft$/i, 'left')]).filter((w): w is OcrWord => Boolean(w));
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


// ---- Biometry alerts --------------------------------------------------------------

/** Cutoffs for the biometry alerts (Settings → Cataract prompts). */
export interface BiometryLimits {
  alVeryShort: number;
  alShort: number;
  alLong: number;
  alVeryLong: number;
  acdShallow: number;
  ltThick: number;
  kFlat: number;
  kSteep: number;
  /** ΔK at or above this with a non-toric lens: consider a toric. */
  toricCyl: number;
  /** ΔK and ΔTK disagreeing by more than this (D, or degrees): re-check. */
  kAgreeCyl: number;
  kAgreeAxis: number;
  /** Differences between the eyes. */
  alDiff: number;
  kDiff: number;
  wtwSmall: number;
  wtwLarge: number;
}

export function defaultLimits(): BiometryLimits {
  return {
    alVeryShort: 21,
    alShort: 22,
    alLong: 26,
    alVeryLong: 28,
    acdShallow: 2.5,
    ltThick: 5,
    kFlat: 41,
    kSteep: 47,
    toricCyl: 1,
    kAgreeCyl: 0.5,
    kAgreeAxis: 15,
    alDiff: 0.3,
    kDiff: 1,
    wtwSmall: 11,
    wtwLarge: 12.5,
  };
}

export interface BiometryAlert {
  level: 'warn' | 'error';
  /** A few words for a badge, e.g. "Long eye". */
  short: string;
  text: string;
  side: 'right' | 'left' | 'both';
  words?: OcrWord[];
}

const mm = (n: number) => `${n.toFixed(2)} mm`;
const dioptre = (n: number) => `${n.toFixed(2)} D`;

/**
 * What in the biometry is worth a second look for the eye(s) being operated on: a short or long eye,
 * shallow AC, thick lens, flat or steep cornea, astigmatism with a non-toric lens, ΔK and ΔTK not
 * agreeing, a small or large cornea, values the IOLMaster marks borderline, previous laser surgery,
 * and (an error) an eye that already has a lens. Plus differences between the two eyes.
 */
export function biometryAlerts(reading: ReportReading, lens: PickedLens, limits: BiometryLimits): BiometryAlert[] {
  const sides: ('right' | 'left')[] = lens.eye === 'RIGHT' ? ['right'] : lens.eye === 'LEFT' ? ['left'] : ['right', 'left'];
  const alerts: BiometryAlert[] = [];
  const toric = Boolean(lens.family && /[1-9]/.test(lens.family.number));
  for (const side of sides) {
    const b = reading[side]?.biometry;
    if (!b) continue;
    const eye = `${side} eye`;
    const add = (level: BiometryAlert['level'], short: string, text: string, words?: OcrWord[]) => alerts.push({ level, short, text, side, words });
    if (b.lensStatus && /pseudo|aphak/i.test(b.lensStatus)) {
      add('error', `${b.lensStatus} ${side === 'right' ? 'RE' : 'LE'}`, `The report says the ${eye} is ${b.lensStatus.toLowerCase()} (no natural lens): check it’s the right eye.`, b.lensStatusWords);
    }
    if (b.lvc && !/untreated/i.test(b.lvc)) add('warn', 'Prior LVC', `Previous laser vision correction on the ${eye} (LVC: ${b.lvc}): check the formula.`, b.lvcWords);
    const al = b.al?.value ?? reading[side]?.axialLength;
    if (al !== undefined) {
      if (al < limits.alVeryShort) add('warn', 'Very short eye', `Very short eye: AL ${mm(al)} (under ${limits.alVeryShort.toFixed(1)}).`, b.al?.words);
      else if (al < limits.alShort) add('warn', 'Short eye', `Short eye: AL ${mm(al)} (under ${limits.alShort.toFixed(1)}).`, b.al?.words);
      else if (al > limits.alVeryLong) add('warn', 'Very long eye', `Very long eye: AL ${mm(al)} (over ${limits.alVeryLong.toFixed(1)}).`, b.al?.words);
      else if (al > limits.alLong) add('warn', 'Long eye', `Long eye: AL ${mm(al)} (over ${limits.alLong.toFixed(1)}).`, b.al?.words);
    }
    const phakic = !b.lensStatus || /^phakic$/i.test(b.lensStatus);
    if (phakic && b.acd && b.acd.value < limits.acdShallow) add('warn', 'Shallow AC', `Shallow AC: ACD ${mm(b.acd.value)} (under ${limits.acdShallow.toFixed(1)}).`, b.acd.words);
    if (phakic && b.lt && b.lt.value > limits.ltThick) add('warn', 'Thick lens', `Thick lens: LT ${mm(b.lt.value)} (over ${limits.ltThick.toFixed(1)}).`, b.lt.words);
    const k = b.k?.value ?? reading[side]?.k;
    if (k !== undefined) {
      if (k < limits.kFlat) add('warn', 'Flat K', `Flat cornea: K ${dioptre(k)} (under ${limits.kFlat.toFixed(1)}).`, b.k?.words);
      else if (k > limits.kSteep) add('warn', 'Steep K', `Steep cornea: K ${dioptre(k)} (over ${limits.kSteep.toFixed(1)}).`, b.k?.words);
    }
    if (b.deltaK && b.deltaK.value >= limits.toricCyl && !toric) {
      add('warn', 'Astigmatism', `Corneal astigmatism ΔK ${dioptre(b.deltaK.value)} @ ${b.deltaK.axis}°${lens.code ? ` with a non-toric lens (${lens.code})` : ''}: consider a toric.`, b.deltaK.words);
    }
    if (b.deltaK && b.deltaTK) {
      const cyl = Math.abs(b.deltaK.value - b.deltaTK.value);
      const raw = Math.abs((b.deltaK.axis ?? 0) - (b.deltaTK.axis ?? 0)) % 180;
      const axis = Math.min(raw, 180 - raw);
      // An axis means little for a small cylinder.
      const axisCounts = Math.min(b.deltaK.value, b.deltaTK.value) >= 0.75;
      if (cyl > limits.kAgreeCyl || (axisCounts && axis > limits.kAgreeAxis)) {
        add('warn', 'K mismatch', `ΔK (${dioptre(b.deltaK.value)} @ ${b.deltaK.axis}°) and ΔTK (${dioptre(b.deltaTK.value)} @ ${b.deltaTK.axis}°) don’t agree: re-check the keratometry.`, [...b.deltaK.words, ...b.deltaTK.words]);
      }
    }
    if (b.wtw && b.wtw.value < limits.wtwSmall) add('warn', 'Small cornea', `Small cornea: WTW ${b.wtw.value.toFixed(1)} mm (under ${limits.wtwSmall.toFixed(1)}).`, b.wtw.words);
    if (b.wtw && b.wtw.value > limits.wtwLarge) add('warn', 'Large cornea', `Large cornea: WTW ${b.wtw.value.toFixed(1)} mm (over ${limits.wtwLarge.toFixed(1)}).`, b.wtw.words);
    for (const line of b.borderline) add('warn', 'Borderline value', `The report marks a value as borderline (!): ${line.text}`, line.words);
  }
  // Between the eyes (the IOLMaster may warn about the axial length too).
  const [r, l] = [reading.right?.biometry, reading.left?.biometry];
  const alR = r?.al?.value ?? reading.right?.axialLength;
  const alL = l?.al?.value ?? reading.left?.axialLength;
  if (alR !== undefined && alL !== undefined && Math.abs(alR - alL) > limits.alDiff + 1e-9) {
    alerts.push({ level: 'warn', short: 'AL differs', side: 'both', text: `The eyes’ axial lengths differ by ${mm(Math.abs(alR - alL))} (over ${limits.alDiff.toFixed(1)}): re-check.`, words: [...(r?.al?.words ?? []), ...(l?.al?.words ?? [])] });
  }
  const kR = r?.k?.value ?? reading.right?.k;
  const kL = l?.k?.value ?? reading.left?.k;
  if (kR !== undefined && kL !== undefined && Math.abs(kR - kL) > limits.kDiff + 1e-9) {
    alerts.push({ level: 'warn', short: 'K differs', side: 'both', text: `The eyes’ keratometry differs by ${dioptre(Math.abs(kR - kL))} (over ${limits.kDiff.toFixed(1)}): re-check.`, words: [...(r?.k?.words ?? []), ...(l?.k?.words ?? [])] });
  }
  return alerts;
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
  /** First given name, if known. */
  givenName?: string;
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
  /** Set when the report's name isn't this patient's: it must be impossible to miss. */
  nameMismatch?: string;
  /** When the report's ID isn't the list's MRN (often another system's number): a note, not an error. */
  idNote?: string;
  estimate?: LensEstimate;
  /** What to mark on the screenshot. */
  highlights: Highlight[];
  /** Biometry worth a second look (long eye, shallow AC…). */
  alerts: BiometryAlert[];
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
      biometry: all.find((e) => e.biometry)?.biometry,
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

const letters = (text: string) => text.toUpperCase().replace(/[^A-Z]/g, '');

/** Two names the same, allowing one letter misread by OCR in a longer name, and part of a double-barrelled one. */
function similarName(a: string, b: string): boolean {
  const [x, y] = [letters(a), letters(b)];
  if (!x || !y) return false;
  if (x === y) return true;
  if (Math.min(x.length, y.length) >= 3 && (x.includes(y) || y.includes(x))) return true;
  return editDistance(x, y) <= (Math.max(x.length, y.length) > 4 ? 1 : 0);
}

/**
 * Whether the name on the report ("SURNAME, Given …") is this patient's: the surname, and the first
 * given name when both are known (an initial matches its name). Family members can share a surname.
 */
export function sameName(reportName: string, surname: string, givenName: string): boolean {
  const [before, after = ''] = reportName.split(',');
  const hasComma = reportName.includes(',');
  const words = reportName.split(/[\s,]+/).filter((w) => letters(w).length > 0);
  const surnameParts = surname.split(/[\s-]+/).filter((w) => letters(w).length > 1);
  const reportSurnames = hasComma ? before.split(/[\s-]+/) : words;
  const surnameOk = (surnameParts.length ? surnameParts : [surname]).some((part) => reportSurnames.some((w) => similarName(w, part)));
  const given = letters(givenName.split(/\s+/)[0] ?? '');
  if (!given) return surnameOk;
  const reportGivens = hasComma ? after.split(/\s+/).filter((w) => letters(w)) : words;
  if (!reportGivens.length) return surnameOk;
  const givenOk = reportGivens.some((w) => {
    const g = letters(w);
    return similarName(g, given) || (g.length === 1 && g === given[0]) || (given.length === 1 && given === g[0]);
  });
  return surnameOk && givenOk;
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
  if (sorted.length < 2) {
    const only = sorted[0] ?? { model: '', power, ref: NaN };
    return { ref: only.ref, low: only, high: only };
  }
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
    if (table.toric || table.rows.length < 2) continue;  // needs two rows to read between
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



/** Lenses the report's calculation for another lens applies to, e.g. { from: 'DIUx', to: 'ZCUx' } (same platform and A-constant). */
export interface LensEquivalent {
  from: string;
  to: string;
}

/** "DIUx = ZCUx" per line, as in Settings. */
export function parseEquivalents(lines: string[]): LensEquivalent[] {
  return lines.flatMap((line) => {
    const m = line.trim().match(/^(\S+)\s*(?:=|→|->)\s*(\S+)$/);
    return m ? [{ from: m[1], to: m[2] }] : [];
  });
}

/** The lens whose calculation on the report to use instead (DIU150 → ZCU150), if one is set for this lens. */
export function equivalentLens(lens: PickedLens, equivalents: LensEquivalent[]): PickedLens | null {
  for (const e of equivalents) {
    const fromAt = e.from.indexOf('x');
    const toAt = e.to.indexOf('x');
    if (fromAt >= 0 && toAt >= 0 && lens.family && canonical(lens.family.base) === canonical(e.from.slice(0, fromAt))) {
      const base = e.to.slice(0, toAt);
      return { ...lens, code: `${base}${lens.family.number}${e.to.slice(toAt + 1)}`, family: { base, number: lens.family.number } };
    }
    if (fromAt < 0 && canonical(lens.code) === canonical(e.from)) return { ...lens, code: e.to.replace('x', ''), family: undefined };
  }
  return null;
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
export function simulateLens(reading: ReportReading, side: 'right' | 'left', lens: PickedLens, constants: LensConstant[], equivalents: LensEquivalent[] = []): Simulation | string {
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

  let tables = tablesFor(eye, lens);
  // Not on the report, but the same calculation as one that is (e.g. a DIU as its ZCU)?
  let via: PickedLens | null = null;
  if (!tables.length) {
    const alias = equivalentLens(lens, equivalents);
    const aliasTables = alias ? tablesFor(eye, alias) : [];
    if (alias && aliasTables.length) {
      tables = aliasTables;
      via = alias;
      notes.push(`${lens.code} isn’t on the report: read from its ${alias.code} calculation (the same lens platform and A-constant, as set in Settings).`);
    }
  }
  const toric = lens.family ? tables.find((t) => t.toric) : undefined;
  if (toric) {
    const number = canonical(lens.family!.number);
    const models = toric.rows.map((r) => r.model).filter(Boolean);
    const recommended = models.sort((a, b) => models.filter((m) => m === b).length - models.filter((m) => m === a).length)[0] ?? '';
    const axis = toric.rows.find((r) => r.axis !== undefined)?.axis;
    const matching = toric.rows.filter((r) => !r.model || canonical(r.model) === number);
    base = matching.length >= 2 ? matching : toric.rows;
    if (recommended && canonical(recommended) !== number) {
      const what = /^[A-Z]+$/.test(recommended) ? 'a non-toric lens' : `${(via ?? lens).family!.base}${recommended}`;
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
export function checkLens(
  reading: ReportReading,
  lens: PickedLens,
  constants: LensConstant[] = [],
  limits: BiometryLimits = defaultLimits(),
  equivalents: LensEquivalent[] = [],
): LensCheck {
  const facts: string[] = [];
  const flags: LensCheck['flags'] = [];
  const notes = [...reading.warnings];
  let estimate: LensEstimate | undefined;
  const operated: Highlight['eye'] = lens.eye === 'RIGHT' ? 'right' : lens.eye === 'LEFT' ? 'left' : 'both';
  const highlights: Highlight[] = [...box(reading.nameWords, operated), ...box(reading.dobWords, operated), ...box(reading.idWords, operated)];

  // Is it this patient's report? The name must match (surname and given name); the ID often won't,
  // as the IOLMaster may use another system's number, so that's only a note.
  let nameMismatch: string | undefined;
  let idNote: string | undefined;
  const ourName = [lens.surname, lens.givenName].filter(Boolean).join(', ');
  if (/[A-Za-z]{2}/.test(reading.name) && lens.surname.trim()) {
    if (!sameName(reading.name, lens.surname, lens.givenName ?? '')) {
      nameMismatch = `NAME MISMATCH: the report is for “${reading.name}”, not “${ourName}”. Check you have this patient’s report.`;
      flags.push({ level: 'error', text: nameMismatch });
    }
  } else if (lens.surname.trim()) {
    flags.push({ level: 'warn', text: 'Couldn’t read the patient’s name on the report: check it’s this patient’s.' });
  }
  const reportId = reading.id.replace(/\D/g, '').replace(/^0+/, '');
  const ourId = lens.mrn.replace(/\D/g, '').replace(/^0+/, '');
  if (reportId.length >= 4 && ourId.length >= 4 && reportId !== ourId) {
    idNote = `The report’s patient ID (${reading.id}) isn’t the MRN on the list (${lens.mrn}): often another system’s number, so check the name and date of birth.`;
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
      let tables = tablesFor(eye, lens);
      // Not on the report, but the same calculation as one that is (e.g. a DIU as its ZCU)?
      let match = lens;
      if (!tables.length) {
        const alias = equivalentLens(lens, equivalents);
        const aliasTables = alias ? tablesFor(eye, alias) : [];
        if (alias && aliasTables.length) {
          tables = aliasTables;
          match = alias;
        }
      }
      const hit = rowFor(tables, match);
      if (hit) {
        found = true;
        highlights.push(...box(hit.table.titleWords, side), ...box(hit.row.words, side));
        const { row } = hit;
        const toric = row.cyl !== undefined && row.axis !== undefined && row.cyl > 0 ? `, axis ${row.axis}°` : '';
        const from = match !== lens ? ` (from the report’s ${match.code} calculation)` : '';
        facts.push(`${side === 'right' ? 'Right' : 'Left'} eye: ${lens.code} ${signed(row.power)}${toric} → predicted ${signed(row.ref)}${from}`);
        continue;
      }
      const other = side === 'right' ? 'left' : 'right';
      const elsewhere = sides.length === 1 ? rowFor(tablesFor(reading[other], match), match) : null;
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
    // The IOLMaster is set to −0.50 to show the most useful range of lenses, so plano and −0.50 are
    // the usual targets: only another target the report wasn't calculated for is worth a note.
    const usual = lens.target !== null && [0, -0.5].some((t) => Math.abs(t - lens.target!) < 0.01);
    if (found && lens.target !== null && !usual) {
      for (const side of sides) {
        const t = reading[side]?.target;
        if (t !== null && t !== undefined && Math.abs(t - lens.target) > 0.01) {
          flags.push({ level: 'warn', text: `The report was calculated for a target of ${signed(t)} D (${side} eye); you picked ${lens.target === 0 ? 'plano' : signed(lens.target)}.` });
        }
      }
    }
  }
  const level: LensCheck['level'] = flags.some((f) => f.level === 'error') ? 'error' : flags.length ? 'warn' : facts.length ? 'ok' : 'none';
  const alerts = biometryAlerts(reading, lens, limits);
  for (const a of alerts) highlights.push(...box(a.words, a.side === 'both' ? operated : a.side));
  return { level: alerts.some((a) => a.level === 'error') ? 'error' : level, facts, flags, notes, estimate, highlights, nameMismatch, idNote, alerts };
}

function tablesOf(eye: EyeReading, code: string): LensRow[] {
  return eye.tables.find((t) => !t.toric && canonical(t.title).includes(canonical(code)))?.rows ?? [];
}

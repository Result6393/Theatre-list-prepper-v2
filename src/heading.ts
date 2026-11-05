import type { Layout } from './layout';
import type { SheetData } from './xlsx';

/**
 * The heading of a cheat sheet: cells above the list labelled Consultant, Date, OR and AM/PM, with
 * the value in the cell to their right (or after a colon in the same cell, e.g. "Consultant: Smith").
 * Saving a list fills these in from the run sheet, so a new sheet copied from another doesn't keep
 * the old list's consultant and date.
 */

export type HeadingField = 'consultant' | 'date' | 'or' | 'ampm';

export interface HeadingValues {
  consultant: string;
  /** "14/10/2026" */
  date: string;
  /** "OR 4" */
  or: string;
  /** "AM", "PM" or "AM + PM" */
  ampm: string;
}

export interface HeadingCell {
  field: HeadingField;
  /** The label as written on the sheet, e.g. "Consultant". */
  label: string;
  /** e.g. "C1" */
  cell: string;
  /** What's written into the cell (a date can be Excel's number for it). */
  value: string;
  /** The same, as you'd read it, e.g. "14/10/2026". */
  shown: string;
}

const LABELS: [HeadingField, RegExp][] = [
  ['consultant', /^(consultant|surgeon)$/i],
  ['date', /^(date|list date|day)$/i],
  ['or', /^(or|theatre|theater|operating room)$/i],
  ['ampm', /^(am\s*\/\s*pm|am\s*-\s*pm|am\s+or\s+pm|session)$/i],
];

function fieldOf(text: string): HeadingField | null {
  const label = text.trim().replace(/\s*:$/, '').replace(/\s+/g, ' ');
  return LABELS.find(([, re]) => re.test(label))?.[0] ?? null;
}

/** Excel keeps dates as numbers of days since 30/12/1899: 14/10/2026 is 46309. */
function dateSerial(date: string): string | null {
  const m = date.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const days = (Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1])) - Date.UTC(1899, 11, 30)) / 86400000;
  return String(days);
}

function columnName(index: number): string {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/**
 * Where the heading's values go on this sheet, and what to write there. Only rows above the list
 * (`listRow`, 0-based) are looked at; the list's column headings and other labels are never
 * written over.
 */
export function headingCells(sheet: SheetData, listRow: number, values: HeadingValues, layout: Layout): HeadingCell[] {
  const headings = new Set(layout.columns.map((c) => c.heading.trim().toLowerCase()).filter(Boolean));
  const isHeading = (text: string) => headings.has(text.trim().toLowerCase());
  const isHeadingRow = (row: string[]) => row.filter(isHeading).length >= 3;
  const merges = (sheet.merges ?? []).map(parseRange).filter((m): m is Range => m !== null);
  const found: HeadingCell[] = [];
  const seen = new Set<HeadingField>();
  for (let r = 0; r < Math.min(listRow, sheet.rows.length); r++) {
    const row = sheet.rows[r];
    if (isHeadingRow(row)) continue;
    for (let c = 0; c < row.length; c++) {
      const text = row[c].trim();
      if (!text) continue;
      const sameCell = text.match(/^([^:]{1,20}):\s*(\S.*)$/);
      const field = fieldOf(sameCell ? sameCell[1] : text);
      if (!field || seen.has(field)) continue;
      if (sameCell) {
        const value = valueFor(field, values, sameCell[2]);
        if (!value) continue;
        found.push({ field, label: sameCell[1].trim(), cell: `${columnName(c)}${r + 1}`, value: `${sameCell[1].trim()}: ${value.text}`, shown: value.text });
      } else {
        // The value goes just right of the label (after it, if the label is merged across cells),
        // unless that's another label or one of the list's column headings.
        const merged = merges.find((m) => m.top <= r && r <= m.bottom && m.left <= c && c <= m.right);
        const v = (merged ? merged.right : c) + 1;
        const next = row[v] ?? '';
        if (fieldOf(next) || isHeading(next)) continue;
        const value = valueFor(field, values, next);
        if (!value) continue;
        found.push({ field, label: text.replace(/\s*:$/, ''), cell: `${columnName(v)}${r + 1}`, value: value.cell, shown: value.text });
      }
      seen.add(field);
    }
  }
  return found;
}

interface Range {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

/** "B1:C2" → rows 0–1, columns 1–2. */
function parseRange(ref: string): Range | null {
  const m = ref.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!m) return null;
  const col = (letters: string) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { left: col(m[1]), top: Number(m[2]) - 1, right: col(m[3] ?? m[1]), bottom: Number(m[4] ?? m[2]) - 1 };
}

/** The value to write, matching how the sheet already writes it where it can (e.g. "4" or "OR 4"). */
function valueFor(field: HeadingField, values: HeadingValues, current: string): { cell: string; text: string } | null {
  const now = current.trim();
  if (field === 'consultant') return values.consultant ? { cell: values.consultant, text: values.consultant } : null;
  if (field === 'ampm') return values.ampm ? { cell: values.ampm, text: values.ampm } : null;
  if (field === 'or') {
    if (!values.or) return null;
    // A sheet that writes just the number next to "OR" (or "Theatre") gets just the number.
    const bare = values.or.replace(/^OR\s*/i, '');
    const value = now === '' || /^\d/.test(now) ? bare : values.or;
    return { cell: value, text: value };
  }
  if (!values.date) return null;
  // A real Excel date stays a date (its format shows it as before); otherwise it's written as text.
  const serial = dateSerial(values.date);
  return { cell: serial && /^\d{5}(\.\d+)?$/.test(now) ? serial : values.date, text: values.date };
}

import Papa from 'papaparse';
import type { Patient, Session } from './extract';
import { buildRow, cellValue, type Layout } from './layout';
import { columnIndex, columnName, type SheetData } from './xlsx';

/**
 * Keeping your prep when the list changes on the day: rows copied back out of the cheat sheet are
 * matched by MRN to the latest run sheet, so notes stay with their patients, new patients are added,
 * and patients who have dropped off are kept at the bottom, labelled as cancelled.
 */

/** 'sheet': a row from your prepped Excel file, before any run sheet is loaded. */
export type RowStatus = 'fresh' | 'kept' | 'new' | 'cancelled' | 'sheet';

export interface GridRow {
  /** The run-sheet booking; null for a row that's only on your sheet (cancelled, or no run sheet yet). */
  patient: Patient | null;
  /** Identifies a row without a booking, e.g. for keeping focus while editing. */
  id?: string;
  cells: string[];
  included: boolean;
  status: RowStatus;
  /** Where the patient was on your sheet (1 = first), when the update moved them. */
  movedFrom?: number;
  /** Was in your cancelled rows and is back on the run sheet. */
  reinstated?: boolean;
  /** Run-sheet values (by column) that differ from what's on your sheet, for you to check. */
  changes: Record<number, string>;
  /** For a cancelled row: the list the patient is on now, if they moved to another one. */
  note?: string;
}

export interface MergeSummary {
  kept: number;
  added: number;
  cancelled: number;
  moved: number;
  changed: number;
}

export interface MergeResult {
  rows: GridRow[];
  /** Rows on your sheet before; the copy is padded with blank rows to this length to clear leftovers. */
  padTo: number;
  summary: MergeSummary;
}

export class MergeError extends Error {}

/** First cell of the row that heads the cancelled patients in the copy. */
export const CANCELLED_LABEL = 'CANCELLED';
export const CANCELLED_NOTE = 'Not on the latest run sheet';

export function freshRows(patients: Patient[], layout: Layout): GridRow[] {
  return patients.map((patient) => ({
    patient,
    cells: buildRow(patient, layout),
    included: true,
    status: 'fresh',
    changes: {},
  }));
}

/** Rows pasted from Excel (tab-separated, multi-line cells in quotes), without trailing empty rows. */
export function parsePastedRows(text: string): string[][] {
  // Excel on Windows ends rows with \r\n, so a bare \n is a line break inside a cell (quoted or not).
  // Set it explicitly: guessing from the first line break goes wrong when a cell's comes first.
  const clean = text.replace(/^﻿/, '');
  const newline = clean.includes('\r\n') ? '\r\n' : '\n';
  const result = Papa.parse<string[]>(clean, { delimiter: '\t', newline, skipEmptyLines: false });
  return withoutTrailingBlankRows(result.data.map((row) => row.map((cell) => cell.replace(/\r\n?/g, '\n'))));
}

export function withoutTrailingBlankRows(rows: string[][]): string[][] {
  const out = [...rows];
  while (out.length && isEmptyRow(out[out.length - 1])) out.pop();
  return out;
}

/** True when every row has the same number of cells, as a clean copy from Excel does. */
export function isRectangular(rows: string[][]): boolean {
  return rows.every((row) => row.length === rows[0].length);
}

/**
 * A list straight from your prepped sheet, before the run sheet is loaded: its block of patient rows
 * on the sheet (in the column with most MRNs), from the first MRN to the last.
 */
export function findListBlock(sheet: SheetData, layout: Layout): FoundRows | null {
  const key = keyColumn(layout);
  const width = layout.columns.length;
  const preferred = columnIndex(layout.pasteStart || 'B5');
  const widest = Math.max(0, ...sheet.rows.map((r) => r.length));
  const starts = [preferred, ...Array.from({ length: Math.min(widest, 12) }, (_, i) => i).filter((i) => i !== preferred)];
  let best: FoundRows | null = null;
  for (const start of starts) {
    const isPatient = sheet.rows.map((r) => key.isPatient(r, start));
    const first = isPatient.indexOf(true);
    if (first === -1) continue;
    const last = isPatient.lastIndexOf(true);
    const patients = isPatient.slice(first, last + 1).filter(Boolean).length;
    if (patients > (best?.matches ?? 0)) {
      best = {
        sheet: sheet.name,
        rows: sheet.rows.slice(first, last + 1).map((r) => Array.from({ length: width }, (_, i) => r[start + i] ?? '')),
        cell: `${columnName(start)}${first + 1}`,
        matches: patients,
        others: 0,
      };
    }
  }
  return best;
}

/** Editable rows from a block on your sheet: patients in order, then any under the CANCELLED row. */
export function sheetRows(block: string[][], idPrefix: string): GridRow[] {
  let cancelled = false;
  const rows: GridRow[] = [];
  block.forEach((cells, i) => {
    if (isEmptyRow(cells)) return;
    const first = cells.find((c) => c.trim()) ?? '';
    if (new RegExp(`^${CANCELLED_LABEL}\\b`, 'i').test(first.trim()) && !cells.slice(1).some((c) => MRN_LIKE.test(normaliseMrn(c.split(/\s+/)[0] ?? '')))) {
      cancelled = true;
      return;
    }
    rows.push({ patient: null, id: `${idPrefix}:${i}`, cells: [...cells], included: true, status: cancelled ? 'cancelled' : 'sheet', changes: {} });
  });
  return rows;
}

/** "0001234" and "1234 " match: Excel can drop leading zeros from an MRN-only column. */
export function normaliseMrn(value: string): string {
  const text = value.trim().toUpperCase().replace(/\s+/g, '');
  return /^\d+$/.test(text) ? text.replace(/^0+(?=\d)/, '') : text;
}

/** Which column holds the MRN, and how to read it from that cell. */
interface KeyColumn {
  index: number;
  read: (cell: string) => string;
  /**
   * Whether a sheet row (its block starting at `start`) looks like a patient: an MRN, plus a name
   * after it in a Patient cell, or other text beside an MRN-only column. A number on its own, such
   * as a date in the sheet's heading (Excel keeps dates as numbers like 46310), isn't a patient.
   */
  isPatient: (row: string[], start: number) => boolean;
}

function keyColumn(layout: Layout): KeyColumn {
  const width = layout.columns.length;
  const exact = layout.columns.findIndex((c) => c.source === 'field:mrn' || c.source === 'raw:MRN');
  if (exact !== -1) {
    return {
      index: exact,
      read: normaliseMrn,
      isPatient: (row, start) =>
        MRN_LIKE.test(normaliseMrn(row[start + exact] ?? '')) &&
        row.slice(start, start + width).some((cell, i) => i !== exact && /[A-Za-z]/.test(cell)),
    };
  }
  const patient = layout.columns.findIndex((c) => c.source === 'field:patient');
  if (patient !== -1) {
    const read = (cell: string) => normaliseMrn(cell.trim().split(/\s+/)[0] ?? '');
    return {
      index: patient,
      read,
      isPatient: (row, start) => {
        const cell = row[start + patient] ?? '';
        return MRN_LIKE.test(read(cell)) && /^\s*\S+\s+.*[A-Za-z]/.test(cell);
      },
    };
  }
  throw new MergeError('To match patients, your columns need a Patient or MRN column (Settings → Columns).');
}

export interface FoundRows {
  sheet: string;
  /** Your rows, from the first patient to the last (cancelled ones included), as if copied. */
  rows: string[][];
  /** Where they start in the sheet, e.g. "B5": paste the update back there. */
  cell: string;
  /** How many of them are on this list. */
  matches: number;
  /** Patients on the sheet who aren't on this list (cancelled, or another list's). */
  others: number;
}

/**
 * How well a sheet fits a list: patients on this list count for it, others against it, so the
 * sheet made for this list beats another list's sheet that a moved patient also appears on.
 */
function fit(found: FoundRows | null): number {
  return found && found.matches > 0 ? 2 * found.matches - found.others : -Infinity;
}

const MRN_LIKE = /^[A-Z]{0,3}\d{4,}[A-Z]?$/;

/**
 * Finds your prepped rows in a workbook: the sheet (and column) with the most patients from this
 * list, from its first MRN down to its last. Headings above and notes below are left out.
 */
export function findPreppedRows(sheets: SheetData[], session: Session, layout: Layout): FoundRows | null {
  let best: FoundRows | null = null;
  for (const sheet of sheets) {
    const found = findInSheet(sheet, session, layout);
    if (fit(found) > fit(best)) best = found;
  }
  return best;
}

/** Your rows for this list on one sheet, or null if none of its patients are there. */
export function findInSheet(sheet: SheetData, session: Session, layout: Layout): FoundRows | null {
  const key = keyColumn(layout);
  const width = layout.columns.length;
  const listMrns = new Set(session.patients.map((p) => normaliseMrn(p.mrn)));
  const preferred = columnIndex(layout.pasteStart || 'B5');
  const widest = Math.max(0, ...sheet.rows.map((r) => r.length));
  // Your paste column first, then the others in case the sheet starts somewhere else.
  const starts = [preferred, ...Array.from({ length: Math.min(widest, 12) }, (_, i) => i).filter((i) => i !== preferred)];
  let best: FoundRows | null = null;
  for (const start of starts) {
    const isPatient = sheet.rows.map((r) => key.isPatient(r, start));
    const first = isPatient.indexOf(true);
    if (first === -1) continue;
    const last = isPatient.lastIndexOf(true);
    const keys = sheet.rows.map((r) => key.read(r[start + key.index] ?? ''));
    let matches = 0;
    let others = 0;
    for (let i = first; i <= last; i++) {
      if (!isPatient[i]) continue;
      if (listMrns.has(keys[i])) matches++;
      else others++;
    }
    const found: FoundRows = {
      sheet: sheet.name,
      rows: sheet.rows.slice(first, last + 1).map((r) => Array.from({ length: width }, (_, i) => r[start + i] ?? '')),
      cell: `${columnName(start)}${first + 1}`,
      matches,
      others,
    };
    if (fit(found) > fit(best)) best = found;
  }
  return best;
}

export function mergePrepped(
  pasted: string[][],
  session: Session,
  allSessions: Session[],
  layout: Layout,
  describe: (session: Session) => string,
): MergeResult {
  const width = layout.columns.length;
  const key = keyColumn(layout);
  const listMrns = new Set(session.patients.map((p) => normaliseMrn(p.mrn)));
  const elsewhere = new Map<string, Session>();
  for (const other of allSessions) {
    if (other.id === session.id) continue;
    for (const p of other.patients) if (!elsewhere.has(normaliseMrn(p.mrn))) elsewhere.set(normaliseMrn(p.mrn), other);
  }

  // If an extra column was copied on the left (e.g. column A's row numbers), skip it.
  const longest = Math.max(0, ...pasted.map((r) => r.length));
  let offset = 0;
  let bestMatches = -1;
  for (let o = 0; o <= Math.min(3, Math.max(0, longest - width)); o++) {
    const matches = pasted.filter((r) => listMrns.has(key.read(r[o + key.index] ?? ''))).length;
    if (matches > bestMatches) {
      offset = o;
      bestMatches = matches;
    }
  }

  const headings = layout.columns.map((c) => c.heading.trim().toLowerCase());
  const waiting = new Map<string, Patient[]>();
  for (const p of session.patients) {
    const mrn = normaliseMrn(p.mrn);
    waiting.set(mrn, [...(waiting.get(mrn) ?? []), p]);
  }

  const matched = new Map<Patient, { cells: string[]; oldPos: number; wasCancelled: boolean }>();
  const cancelled: GridRow[] = [];
  let inCancelled = false;
  let position = 0;
  let padTo = 0;

  for (const raw of pasted) {
    const cells = Array.from({ length: width }, (_, i) => raw[offset + i] ?? '');
    if (cells.some(Boolean) && cells.every((c, i) => c.trim().toLowerCase() === headings[i])) continue;
    padTo += 1;
    if (isEmptyRow(cells)) continue;
    const mrn = key.read(cells[key.index]);
    const patient = waiting.get(mrn)?.shift();
    if (patient) {
      if (!inCancelled) position += 1;
      matched.set(patient, { cells, oldPos: position, wasCancelled: inCancelled });
      continue;
    }
    const first = cells.find((c) => c.trim()) ?? '';
    if (new RegExp(`^${CANCELLED_LABEL}\\b`, 'i').test(first.trim())) {
      inCancelled = true;
      continue;
    }
    if (!inCancelled) position += 1;
    const now = elsewhere.get(mrn);
    cancelled.push({
      patient: null,
      cells,
      included: true,
      status: 'cancelled',
      changes: {},
      note: now ? `Now on ${describe(now)}` : undefined,
    });
  }

  if (matched.size === 0 && cancelled.length > 0) {
    throw new MergeError(
      `None of the pasted patients are on this list (${describe(session)}). Check you picked the right list and copied the rows from your sheet.`,
    );
  }
  if (matched.size === 0) throw new MergeError('No patient rows found in what you pasted.');

  const rows: GridRow[] = session.patients.map((patient) => {
    const prepped = matched.get(patient);
    if (!prepped) {
      return { patient, cells: buildRow(patient, layout), included: true, status: 'new', changes: {} };
    }
    const changes: Record<number, string> = {};
    layout.columns.forEach((column, i) => {
      if (column.source === 'blank') return;
      const latest = cellValue(patient, column.source);
      if (latest.trim() && !squash(prepped.cells[i]).includes(squash(latest))) changes[i] = latest;
    });
    return {
      patient,
      cells: prepped.cells,
      included: true,
      status: 'kept',
      reinstated: prepped.wasCancelled || undefined,
      changes,
    };
  });

  // Patients outside the longest run still in their old order are the ones that moved.
  const kept = rows.filter((r) => r.status === 'kept' && !r.reinstated);
  const oldOrder = kept.map((r) => matched.get(r.patient!)!.oldPos);
  const stayed = inEveryLongestRun(oldOrder);
  kept.forEach((r, i) => {
    if (!stayed.has(i)) r.movedFrom = oldOrder[i];
  });

  return {
    rows: [...rows, ...cancelled],
    padTo,
    summary: {
      kept: rows.filter((r) => r.status === 'kept').length,
      added: rows.filter((r) => r.status === 'new').length,
      cancelled: cancelled.length,
      moved: kept.filter((r) => r.movedFrom !== undefined).length,
      changed: rows.filter((r) => Object.keys(r.changes).length > 0).length,
    },
  };
}

/**
 * The rows to paste back over your sheet: the list, then the ticked cancelled patients under a
 * label row, then blank rows to clear anything left over from the old, longer block.
 */
export function copyBlock(rows: GridRow[], layout: Layout, padTo = 0): string[][] {
  const width = layout.columns.length;
  const blank = () => Array.from({ length: width }, () => '');
  const out = rows.filter((r) => r.status !== 'cancelled' && r.included).map((r) => r.cells);
  const cancelled = rows.filter((r) => r.status === 'cancelled' && r.included);
  if (cancelled.length) {
    const label = blank();
    label[0] = CANCELLED_LABEL;
    let noteAt = layout.columns.findIndex((c) => c.source === 'field:patient' || c.source === 'field:mrn');
    if (noteAt <= 0) noteAt = width > 1 ? 1 : -1;
    if (noteAt > 0) label[noteAt] = CANCELLED_NOTE;
    out.push(blank(), label, ...cancelled.map((r) => r.cells));
  }
  while (out.length < padTo) out.push(blank());
  return out;
}

function isEmptyRow(row: string[]): boolean {
  return row.every((cell) => !cell.trim());
}

function squash(text: string): string {
  return text.toLowerCase().replace(/\s+/g, ' ').trim();
}

/**
 * Patients who certainly kept their place: those in every longest run still in the old order.
 * When two swap, neither is certain, so both show as moved; when one jumps, only that one does.
 */
function inEveryLongestRun(values: number[]): Set<number> {
  const n = values.length;
  const ending = values.map(() => 1);
  const starting = values.map(() => 1);
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < i; j++) if (values[j] < values[i]) ending[i] = Math.max(ending[i], ending[j] + 1);
  }
  for (let i = n - 1; i >= 0; i--) {
    for (let j = i + 1; j < n; j++) if (values[j] > values[i]) starting[i] = Math.max(starting[i], starting[j] + 1);
  }
  const longest = Math.max(0, ...ending);
  const onSome = values.map((_, i) => ending[i] + starting[i] - 1 === longest);
  const atStep = new Map<number, number>();
  values.forEach((_, i) => {
    if (onSome[i]) atStep.set(ending[i], (atStep.get(ending[i]) ?? 0) + 1);
  });
  return new Set(values.map((_, i) => i).filter((i) => onSome[i] && atStep.get(ending[i]) === 1));
}

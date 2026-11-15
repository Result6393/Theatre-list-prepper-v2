import type { Column, Layout } from './layout';
import type { SheetData } from './xlsx';

/**
 * Your own part in each case, for your logbook and audit: one of a few levels, and a free-text note
 * (e.g. "did phaco only"). Each has its own column on the sheet, right of the consultant's columns,
 * so the levels can be filtered and counted in Excel. Like every other pick, it's just the cell's text.
 */

export const INVOLVEMENT_LEVELS = ['Observed', 'Supervised', 'Assisted', 'Unsupervised'];

export const INVOLVEMENT_HEADING = 'Involvement';
export const INVOLVEMENT_NOTES_HEADING = 'Involvement notes';

const NOTES = /\b(involvement|role|audit)\b.*\b(notes?|comments?|details?)\b/i;
const LEVEL = /\b(involvement|my role)\b/i;

/** Which columns hold the level and the note, or -1 if that column isn't on your sheet. */
export function involvementColumns(layout: Layout): { level: number; notes: number } {
  const notes = layout.columns.findIndex((c) => NOTES.test(c.heading));
  const level = layout.columns.findIndex((c, i) => i !== notes && LEVEL.test(c.heading));
  return { level, notes };
}

/** The level a cell holds, as listed ("supervised " → "Supervised"); '' if it's empty or something else. */
export function readInvolvement(cell: string): string {
  const text = cell.trim().toLowerCase();
  return INVOLVEMENT_LEVELS.find((level) => level.toLowerCase() === text) ?? '';
}

/** Headings of the cheat sheet before the involvement columns were added. */
const BEFORE_INVOLVEMENT = ['eye', 'patient', 'procedure', 'diagnosis & relevant history', 'exam (va / iop)', 'drops / allergies', 'implants / to order', 'ordered', 'alerts', 'plan & notes'];

/**
 * Saved columns still exactly as they were by default before involvement was added get the two new
 * columns on the end. Columns you've changed yourself are left alone: add them in Settings → Columns.
 */
export function withInvolvementColumns(columns: Column[]): Column[] {
  const same = columns.length === BEFORE_INVOLVEMENT.length && columns.every((c, i) => c.heading.trim().toLowerCase() === BEFORE_INVOLVEMENT[i]);
  if (!same) return columns;
  return [...columns, { heading: INVOLVEMENT_HEADING, source: 'blank' }, { heading: INVOLVEMENT_NOTES_HEADING, source: 'blank' }];
}

/**
 * A sheet made before a column was added to your layout has no heading for it. Given where the list
 * starts (e.g. "B5"), returns the headings to add to the row above it: only where that row already
 * has your other headings in place, and only into empty cells.
 */
export function missingHeadings(sheet: SheetData, listCell: string, layout: Layout, columns: number[]): { cell: string; heading: string }[] {
  const row = Number(listCell.match(/\d+/)?.[0] ?? 1) - 2;
  const letters = listCell.match(/^[A-Za-z]+/)?.[0] ?? 'A';
  const start = letters.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  if (row < 0) return [];
  const cells = sheet.rows[row] ?? [];
  const norm = (text: string) => text.trim().toLowerCase();
  const inPlace = layout.columns.filter((c, i) => c.heading.trim() && norm(cells[start + i] ?? '') === norm(c.heading)).length;
  if (inPlace < 3) return [];
  return columns
    .filter((i) => i >= 0 && layout.columns[i]?.heading.trim() && !(cells[start + i] ?? '').trim())
    .map((i) => ({ cell: `${letterOf(start + i)}${row + 1}`, heading: layout.columns[i].heading }));
}

function letterOf(index: number): string {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

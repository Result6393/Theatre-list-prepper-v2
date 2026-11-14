import { formatPower, lensCode, readLens, readTarget } from './prompts';

/**
 * The lens email for the theatre coordinator: a table of each ticked cataract patient with the lens
 * picked in the Implants box (model, power and target), then how many of each lens are needed.
 */

export interface LensItem {
  /** e.g. "1000001 SURNAME, Given names". */
  who: string;
  /** RIGHT, LEFT or BOTH, as in the Eye box. */
  eye: string;
  /** What's in the Implants box. */
  cell: string;
}

export interface LensLine {
  eye: string;
  who: string;
  /** Lens code, e.g. "CNA0T4", or what you typed when it isn't one of your listed lenses. */
  model: string;
  /** e.g. "+20.0D"; empty for a typed lens. */
  power: string;
  /** e.g. "plano" or "-0.50"; empty if none set. */
  target: string;
}

export interface LensCount {
  model: string;
  power: string;
  count: number;
}

const SHORT_EYE: Record<string, string> = { RIGHT: 'RE', LEFT: 'LE', BOTH: 'BE' };

export function shortEye(eye: string): string {
  return SHORT_EYE[eye.trim().toUpperCase()] ?? '';
}

/** The lens picked, or else whatever you typed in the box (minus the target line). */
function lensOf(cell: string, lenses: string[]): { model: string; power: string } {
  const pick = readLens(cell, lenses);
  if (pick) return { model: lensCode(pick), power: formatPower(pick.power) };
  const typed = cell
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line && !/^target\b/i.test(line))
    .join('; ');
  return { model: typed, power: '' };
}

/** Lines for patients with something in the Implants box, and those without (left out of the list). */
export function lensLines(items: LensItem[], lenses: string[]): { lines: LensLine[]; missing: LensItem[] } {
  const lines: LensLine[] = [];
  const missing: LensItem[] = [];
  for (const item of items) {
    const { model, power } = lensOf(item.cell, lenses);
    if (!model) {
      missing.push(item);
      continue;
    }
    lines.push({ eye: shortEye(item.eye), who: item.who, model, power, target: readTarget(item.cell) ?? '' });
  }
  return { lines, missing };
}

function powerValue(power: string): number {
  const n = parseFloat(power.replace(/[^\d.+-]/g, ''));
  return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY;
}

/** How many of each lens: by model, then power (lowest first). */
export function lensSummary(lines: LensLine[]): LensCount[] {
  const counts = new Map<string, LensCount>();
  for (const { model, power } of lines) {
    const key = `${model}\u0000${power}`;
    const found = counts.get(key);
    if (found) found.count += 1;
    else counts.set(key, { model, power, count: 1 });
  }
  return [...counts.values()].sort(
    (a, b) =>
      a.model.localeCompare(b.model, undefined, { numeric: true, sensitivity: 'base' }) ||
      powerValue(a.power) - powerValue(b.power) ||
      a.power.localeCompare(b.power),
  );
}

export interface LensEmailInfo {
  /** The list, e.g. the consultant's name. */
  list: string;
  /** e.g. "Tuesday 13/10". */
  day: string;
  /** Your name for the sign-off; left off if empty. */
  name: string;
}

const PATIENT_HEAD = ['Eye', 'Patient', 'Lens', 'Power', 'Target'];
const SUMMARY_HEAD = ['Lens', 'Power', 'Number'];

function intro({ list, day }: LensEmailInfo): string {
  const which = list ? `the ${list} list` : 'the list';
  return `Here is the list of lenses for ${which}${day ? ` on ${day}` : ''}.`;
}

function signOff(info: LensEmailInfo): string[] {
  return ['All the best,', ...(info.name.trim() ? [info.name.trim()] : [])];
}

const patientRow = (l: LensLine) => [l.eye, l.who, l.model, l.power, l.target];
const summaryRow = (c: LensCount) => [c.model, c.power, String(c.count)];

/** The whole email as plain text, for apps that don't take formatted text. */
export function lensEmailText(lines: LensLine[], info: LensEmailInfo): string {
  const join = (cells: string[]) => cells.filter(Boolean).join(' · ');
  return [
    'Hi All,',
    '',
    intro(info),
    '',
    ...lines.map((l) => join([l.eye, l.who, l.model, l.power, l.target ? `Target ${l.target}` : ''])),
    '',
    'In summary, here are the lenses needed:',
    '',
    ...lensSummary(lines).map((c) => `${join([c.model, c.power])} × ${c.count}`),
    '',
    ...signOff(info),
  ].join('\n');
}

const escapeHtml = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const CELL = 'border:1px solid #999;padding:3px 8px;text-align:left';

function table(head: string[], rows: string[][]): string {
  const cells = (tag: string, row: string[]) => row.map((c) => `<${tag} style="${CELL}">${escapeHtml(c)}</${tag}>`).join('');
  return (
    '<table style="border-collapse:collapse" cellpadding="0" cellspacing="0">' +
    `<thead><tr>${cells('th', head)}</tr></thead>` +
    `<tbody>${rows.map((r) => `<tr>${cells('td', r)}</tr>`).join('')}</tbody>` +
    '</table>'
  );
}

/**
 * The whole email as HTML, with the patients and the summary as tables, for pasting into Outlook.
 * Only the tables are styled (borders), so the text keeps your default font.
 */
export function lensEmailHtml(lines: LensLine[], info: LensEmailInfo): string {
  const line = (t: string) => `<div>${escapeHtml(t)}</div>`;
  const gap = '<div><br></div>';
  return [
    line('Hi All,'),
    gap,
    line(intro(info)),
    gap,
    table(PATIENT_HEAD, lines.map(patientRow)),
    gap,
    line('In summary, here are the lenses needed:'),
    gap,
    table(SUMMARY_HEAD, lensSummary(lines).map(summaryRow)),
    gap,
    ...signOff(info).map(line),
  ].join('');
}

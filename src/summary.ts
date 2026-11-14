/**
 * A summary of the prepped list to send on (e.g. to the consultant): as text and a table to paste into
 * an email, or as a page to print or save as a PDF. Only what's on your list; nothing is sent anywhere.
 */

export interface SummaryRow {
  /** Booked time, if known. */
  time: string;
  /** The list's columns, in order. */
  cells: string[];
  /** What the IOL report check found, if there's a report. */
  check: string;
}

export interface Summary {
  /** e.g. "OR 4 · AM · Smith · Tuesday 14/10". */
  title: string;
  /** Column headings, in order. */
  headings: string[];
  /** The columns that head each patient in the text version (e.g. eye and patient). */
  nameColumns: number[];
  rows: SummaryRow[];
}

const escape = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const oneLine = (t: string) =>
  t
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean)
    .join(' / ');

/** Columns with something in them for at least one patient (empty ones are left out). */
function usedColumns(s: Summary): number[] {
  return s.headings.map((_, i) => i).filter((i) => s.rows.some((r) => (r.cells[i] ?? '').trim()));
}

function countLine(s: Summary): string {
  return `${s.rows.length} patient${s.rows.length === 1 ? '' : 's'}`;
}

/** Plain text: a numbered block per patient, each column on its own line ("Heading: value"). */
export function summaryText(s: Summary): string {
  const columns = usedColumns(s).filter((i) => !s.nameColumns.includes(i));
  const blocks = s.rows.map((r, n) => {
    const head = [r.time, ...s.nameColumns.map((i) => oneLine(r.cells[i] ?? ''))].filter(Boolean).join(' · ');
    const lines = columns.filter((i) => (r.cells[i] ?? '').trim()).map((i) => `   ${s.headings[i]}: ${oneLine(r.cells[i])}`);
    if (r.check) lines.push(`   IOL report: ${r.check}`);
    return [`${n + 1}. ${head}`, ...lines].join('\n');
  });
  return [`${s.title}: ${countLine(s)}`, '', blocks.join('\n\n')].join('\n');
}

function table(s: Summary, cellStyle: string): string {
  const columns = usedColumns(s);
  const hasTime = s.rows.some((r) => r.time);
  const hasCheck = s.rows.some((r) => r.check);
  const cell = (tag: 'th' | 'td', text: string) => `<${tag} style="${cellStyle}">${escape(text).replace(/\n/g, '<br>')}</${tag}>`;
  const head = ['#', ...(hasTime ? ['Time'] : []), ...columns.map((i) => s.headings[i]), ...(hasCheck ? ['IOL report'] : [])];
  const rows = s.rows.map((r, n) =>
    [String(n + 1), ...(hasTime ? [r.time] : []), ...columns.map((i) => (r.cells[i] ?? '').trim()), ...(hasCheck ? [r.check] : [])],
  );
  return (
    '<table style="border-collapse:collapse" cellpadding="0" cellspacing="0">' +
    `<thead><tr>${head.map((h) => cell('th', h)).join('')}</tr></thead>` +
    `<tbody>${rows.map((row) => `<tr>${row.map((c) => cell('td', c)).join('')}</tr>`).join('')}</tbody></table>`
  );
}

/** For pasting into an email: a heading line and a table (no fonts set, so the email's own is used). */
export function summaryHtml(s: Summary): string {
  return `<div>${escape(`${s.title}: ${countLine(s)}`)}</div><div><br></div>${table(s, 'border:1px solid #999;padding:3px 6px;text-align:left;vertical-align:top')}`;
}

/** A page to print, or save as a PDF from the print dialog: landscape, with the list as a table. */
export function summaryPage(s: Summary, made: string): string {
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>${escape(s.title)}</title><style>
@page { size: A4 landscape; margin: 12mm; }
body { font: 9.5pt/1.35 system-ui, -apple-system, "Segoe UI", sans-serif; color: #111; margin: 0; }
h1 { font-size: 13pt; margin: 0 0 2pt; }
p { margin: 0 0 8pt; color: #555; }
table { width: 100%; border-collapse: collapse; }
th, td { border: 1px solid #999; padding: 3pt 5pt; text-align: left; vertical-align: top; }
th { background: #eef1f5; }
tr { break-inside: avoid; }
</style></head><body><h1>${escape(s.title)}</h1><p>${escape(countLine(s))} · ${escape(made)}</p>${table(s, '')}</body></html>`;
}

import type { Patient } from './extract';
import type { Layout } from './layout';

/**
 * Cataract prompts: one-click findings, lens and target under each lens patient. Each click edits the
 * cell text itself, so nothing is added unless you pick it, and the sheet stays the only record:
 * picks show as selected again when your rows come back in through "Update from my sheet".
 */

export interface PromptSettings {
  enabled: boolean;
  /** One per entry. "A | B | C" is a pick-one group, e.g. good, moderate or poor dilation. */
  findings: string[];
  /** Lens codes. An "x" stands for a number you type, e.g. the toric power in "CNA0Tx". */
  lenses: string[];
  /**
   * Each lens's SRK/T A-constant, "DCB00 119.3" per line, to estimate a lens that isn't on the IOL
   * report from one that is.
   */
  constants: string[];
  /** Column headings the picks go into; '' picks automatically. */
  findingsColumn: string;
  lensColumn: string;
}

export function defaultPrompts(): PromptSettings {
  return {
    enabled: true,
    findings: ['PXF', 'Guttata', 'Good dilation | Mod dilation | Poor dilation', 'Phacodonesis', 'Tamsulosin'],
    lenses: ['DCB00', 'DIB00', 'ZCUx', 'CNA0Tx'],
    constants: ['DCB00 119.3', 'CNA0Tx 119.1'],
    findingsColumn: '',
    lensColumn: '',
  };
}

export interface FindingGroup {
  /** Shown before a pick-one group, e.g. "Dilation". */
  label: string;
  options: { token: string; short: string }[];
}

/** "Good dilation | Mod dilation | Poor dilation" → Dilation: Good / Mod / Poor. */
export function findingGroups(findings: string[]): FindingGroup[] {
  return findings
    .map((f) => f.split('|').map((t) => t.trim()).filter(Boolean))
    .filter((tokens) => tokens.length > 0)
    .map((tokens) => {
      if (tokens.length === 1) return { label: '', options: [{ token: tokens[0], short: tokens[0] }] };
      const last = (t: string) => t.split(/\s+/).pop()!.toLowerCase();
      const shared = tokens.every((t) => t.includes(' ') && last(t) === last(tokens[0])) ? last(tokens[0]) : '';
      return {
        label: shared ? shared[0].toUpperCase() + shared.slice(1) : '',
        options: tokens.map((token) => ({ token, short: shared ? token.slice(0, token.length - shared.length).trim() : token })),
      };
    });
}

/** Lens surgery, where the lens and target prompts make sense. */
export function isLensSurgery(patient: Patient): boolean {
  return /cataract|phaco|\b(pc)?iol\b|intra\s*-?\s*ocular|\blens/i.test(patient.procedureRaw);
}

/** Which columns the picks go into, or -1 if that column isn't on your sheet. */
export function promptColumns(layout: Layout, prompts: PromptSettings): { findings: number; lens: number } {
  const byHeading = (heading: string) => layout.columns.findIndex((c) => norm(c.heading) === norm(heading));
  const pick = (chosen: string, auto: () => number) => (chosen.trim() ? byHeading(chosen) : auto());
  return {
    findings: pick(prompts.findingsColumn, () => {
      const history = layout.columns.findIndex((c) => c.source === 'field:history');
      return history !== -1 ? history : layout.columns.findIndex((c) => /diagnosis|history/i.test(c.heading));
    }),
    lens: pick(prompts.lensColumn, () => layout.columns.findIndex((c) => /implant|lens|\biol\b/i.test(c.heading))),
  };
}

// ---- Findings -------------------------------------------------------------------

const SEPARATORS = /^[\s,;]+|[\s,;]+$/g;

function tokenPattern(token: string, flags = 'i'): RegExp {
  return new RegExp(`(^|[^A-Za-z0-9])${escape(token)}(?=$|[^A-Za-z0-9])`, flags);
}

export function hasToken(cell: string, token: string): boolean {
  return tokenPattern(token).test(cell);
}

/** Adds a finding to the line that already has your findings, or on a new line. */
export function addToken(cell: string, token: string, findings: string[]): string {
  if (hasToken(cell, token)) return cell;
  const all = findingGroups(findings).flatMap((g) => g.options.map((o) => o.token));
  const lines = cell ? cell.split('\n') : [];
  const at = lines.findIndex((line) => all.some((t) => hasToken(line, t)));
  if (at === -1) return [...lines.filter((l, i) => l.trim() || i < lines.length - 1), token].join('\n');
  lines[at] = `${lines[at].replace(SEPARATORS, '')}, ${token}`;
  return lines.join('\n');
}

export function removeToken(cell: string, token: string): string {
  return cell
    .split('\n')
    .map((line) => {
      if (!hasToken(line, token)) return line;
      const cut = line
        .replace(tokenPattern(token, 'gi'), '$1')
        .replace(/\s*([,;])(\s*[,;])+/g, '$1')
        .replace(/\s+([,;])/g, '$1')
        .replace(SEPARATORS, '')
        .replace(/ {2,}/g, ' ');
      return cut.trim() ? cut : null;
    })
    .filter((line): line is string => line !== null)
    .join('\n');
}

/** Picks one option in a group (clearing the others), or clears the group when it's already picked. */
export function toggleFinding(cell: string, group: FindingGroup, token: string, findings: string[]): string {
  const wasOn = hasToken(cell, token);
  let out = group.options.reduce((text, o) => removeToken(text, o.token), cell);
  if (!wasOn) out = addToken(out, token, findings);
  return out;
}

// ---- First or second eye --------------------------------------------------------

export type EyeOrder = 'first' | 'second';

const EYE_ORDER = /\b(1st|first|2nd|second)\s+eye\b\.?[ \t]*/gi;

/** "1st eye" or "2nd eye" anywhere in the cell (e.g. "NS2. 1st eye. Macula ok."), also as "first"/"second". */
export function readEyeOrder(cell: string): EyeOrder | null {
  const m = /\b(1st|first|2nd|second)\s+eye\b/i.exec(cell);
  if (!m) return null;
  return /^(1st|first)$/i.test(m[1]) ? 'first' : 'second';
}

/** Puts "1st eye" or "2nd eye" on its own line at the top (replacing any old mention), or removes it. */
export function writeEyeOrder(cell: string, order: EyeOrder | null): string {
  const lines = (cell ? cell.split('\n') : []).flatMap((line) => {
    if (readEyeOrder(line) === null) return [line];
    // Take the old mention out; a line that was only that mention goes altogether.
    const cut = line.replace(EYE_ORDER, '').replace(/([,;])(\s*[,;])+/g, '$1').replace(/\s+([,;])/g, '$1').replace(/^[\s,;.]+/, '').replace(/[ \t]{2,}/g, ' ').replace(/[\s,;]+$/, '');
    return cut.trim() ? [cut] : [];
  });
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  if (order) lines.unshift(order === 'first' ? '1st eye' : '2nd eye');
  return lines.join('\n');
}

// ---- Lens and target ------------------------------------------------------------

export interface LensPick {
  /** The lens as listed in settings, e.g. "CNA0Tx". */
  lens: string;
  /** What replaces the "x", e.g. "4" for CNA0T4. */
  extra: string;
  power: string;
}

function lensPattern(lens: string): RegExp {
  const at = lens.indexOf('x');
  return at === -1
    ? new RegExp(`^${escape(lens)}$`, 'i')
    : new RegExp(`^${escape(lens.slice(0, at))}([A-Za-z0-9.]*)${escape(lens.slice(at + 1))}$`, 'i');
}

function findLensLine(lines: string[], lenses: string[]): { index: number; pick: LensPick } | null {
  for (const [index, line] of lines.entries()) {
    const [first = '', ...rest] = line.trim().split(/\s+/);
    // Exact codes first, so "DCB00" isn't read as a typed-number lens.
    const ordered = [...lenses.filter((l) => !l.includes('x')), ...lenses.filter((l) => l.includes('x'))];
    for (const lens of ordered) {
      const m = first.match(lensPattern(lens));
      if (m) return { index, pick: { lens, extra: m[1] ?? '', power: rest.join(' ') } };
    }
  }
  return null;
}

export function readLens(cell: string, lenses: string[]): LensPick | null {
  return findLensLine(cell.split('\n'), lenses)?.pick ?? null;
}

export function lensCode(pick: LensPick): string {
  return pick.lens.replace('x', pick.extra.replace(/[^A-Za-z0-9.]/g, ''));
}

/** Puts the lens line at the top of the cell (replacing any old one), or removes it. */
export function writeLens(cell: string, lenses: string[], pick: LensPick | null): string {
  const lines = cell ? cell.split('\n') : [];
  const found = findLensLine(lines, lenses);
  if (found) lines.splice(found.index, 1);
  if (pick) lines.unshift([lensCode(pick), formatPower(pick.power)].filter(Boolean).join(' '));
  return lines.join('\n');
}

/** "21.5" → "+21.5D", "-1" → "-1.0D"; anything else is kept as typed. */
export function formatPower(text: string): string {
  const t = text.trim();
  const m = t.match(/^([+-]?)(\d+(?:\.\d+)?)\s*d?$/i);
  if (!m) return t;
  const n = Number(m[2]);
  return `${m[1] === '-' && n !== 0 ? '-' : '+'}${n.toFixed(1)}D`;
}

const TARGET_LINE = /^\s*target\b[:\s]*(.*)$/i;

export function readTarget(cell: string): string | null {
  for (const line of cell.split('\n')) {
    const m = line.match(TARGET_LINE);
    if (m) return m[1].trim();
  }
  return null;
}

/** Puts "Target …" under the lens line (replacing any old one), or removes it. */
export function writeTarget(cell: string, lenses: string[], target: string | null): string {
  const lines = (cell ? cell.split('\n') : []).filter((line) => !TARGET_LINE.test(line));
  const value = target === null ? '' : formatTarget(target);
  if (value) {
    const lens = findLensLine(lines, lenses);
    lines.splice(lens ? lens.index + 1 : 0, 0, `Target ${value}`);
  }
  return lines.join('\n');
}

/** "0" → "plano", "-0.5" → "-0.50", "+0.25" → "+0.25". Without a sign it's kept as typed, never guessed. */
export function formatTarget(text: string): string {
  const t = text.trim();
  if (/^(plano|pl)$/i.test(t)) return 'plano';
  const m = t.match(/^([+-]?)(\d*(?:\.\d+)?)\s*d?$/i);
  if (!m || !m[2]) return t;
  const n = Number(m[2]);
  if (n === 0) return 'plano';
  return m[1] ? `${m[1]}${n.toFixed(2)}` : t;
}

function norm(text: string): string {
  return text.trim().toLowerCase();
}

function escape(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

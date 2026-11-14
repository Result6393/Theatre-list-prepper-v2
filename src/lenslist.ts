import { lensCode, readLens, readTarget } from './prompts';

/**
 * The lens list for the theatre coordinator: one line per ticked cataract patient with the lens picked
 * in the Implants box (code and power, plus the target), ready to paste into an email.
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
  lens: string;
  target: string;
}

const SHORT_EYE: Record<string, string> = { RIGHT: 'RE', LEFT: 'LE', BOTH: 'BE' };

export function shortEye(eye: string): string {
  return SHORT_EYE[eye.trim().toUpperCase()] ?? '';
}

/** What to tell the coordinator: the lens picked, or else whatever you typed in the box (minus the target line). */
function lensText(cell: string, lenses: string[]): string {
  const pick = readLens(cell, lenses);
  if (pick) return [lensCode(pick), pick.power.trim()].filter(Boolean).join(' ');
  return cell
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line && !/^target\b/i.test(line))
    .join('; ');
}

/** Lines for patients with something in the Implants box, and those without (left out of the list). */
export function lensLines(items: LensItem[], lenses: string[]): { lines: LensLine[]; missing: LensItem[] } {
  const lines: LensLine[] = [];
  const missing: LensItem[] = [];
  for (const item of items) {
    const lens = lensText(item.cell, lenses);
    if (!lens) {
      missing.push(item);
      continue;
    }
    const target = readTarget(item.cell);
    lines.push({ eye: shortEye(item.eye), who: item.who, lens, target: target ? `Target ${target}` : '' });
  }
  return { lines, missing };
}

/** "RE · 1000001 SURNAME, Given · CNA0T4 +20.0D · Target -0.50", one patient per line under a heading. */
export function lensListText(lines: LensLine[], heading: string): string {
  const body = lines.map((l) => [l.eye, l.who, l.lens, l.target].filter(Boolean).join(' · '));
  return [heading, '', ...body].join('\n');
}

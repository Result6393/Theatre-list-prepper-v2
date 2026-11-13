import type { Table } from './parse';

export type Eye = 'RIGHT' | 'LEFT' | 'BOTH' | '';

export interface Patient {
  id: string;
  raw: Record<string, string>;
  time: string;
  mrn: string;
  name: string;
  ageSex: string;
  patient: string;
  eye: Eye;
  procedure: string;
  procedureRaw: string;
  interpreter: string;
  diabetic: string;
  history: string;
  allergy: string;
  /** The run sheet says the patient has an allergy (Y) but not what it is. */
  allergyPending: boolean;
  alerts: string;
  hospitalFlags: string[];
  planNotes: string;
  comment: string;
  consultant: string;
  consultantSurname: string;
  ward: string;
  admissionType: string;
  arrival: string;
  priority: string;
}

export interface Session {
  id: string;
  theatre: string;
  or: string;
  date: string;
  start: string;
  end: string;
  /** AM or PM; "AM + PM" for lists combined with combineSessions. */
  ampm: string;
  slotName: string;
  /** Surname of the consultant the session belongs to, from the slot name. */
  slotOwner: string;
  /** Slot owner, or the most common booking consultant for unnamed slots. */
  consultant: string;
  patients: Patient[];
}

export interface ExtractOptions {
  /** Hospitals to flag in Plan & notes when they appear anywhere on a row. */
  flagHospitals: string[];
  /** Hospital codes dropped from comments when they appear on their own. */
  hospitalCodes: string[];
}

/** Groups run-sheet rows into theatre sessions with cleaned-up patients. */
export function extractSessions(table: Table, options: ExtractOptions): Session[] {
  const sessions = new Map<string, Session>();

  for (const row of table.rows) {
    const type = (row.TYPE ?? '').toLowerCase();
    if (type !== 'slot' && type !== 'booking') continue;
    const key = [row.THEATRE, row.SLOT_BEG, row.SLOT_NAME].join('|');
    let session = sessions.get(key);
    if (!session) {
      session = newSession(key, row);
      sessions.set(key, session);
    }
    if (type === 'booking' && row.PT_NAME) {
      session.patients.push(extractPatient(row, `${key}|${session.patients.length}`, options));
    }
  }

  const list = [...sessions.values()];
  for (const session of list) {
    session.patients.sort((a, b) => a.time.localeCompare(b.time));
    session.consultant = session.slotOwner || mostCommon(session.patients.map((p) => p.consultantSurname));
  }
  return list.sort(
    (a, b) =>
      dateKey(a.date).localeCompare(dateKey(b.date)) ||
      a.or.localeCompare(b.or, undefined, { numeric: true }) ||
      a.start.localeCompare(b.start),
  );
}

/** "13/10/2026" → "20261013", so dates sort correctly as text. Unparseable dates sort last. */
export function dateKey(date: string): string {
  const m = date.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? `${m[3]}${m[2].padStart(2, '0')}${m[1].padStart(2, '0')}` : '99999999';
}

function newSession(id: string, row: Record<string, string>): Session {
  const begin = splitDateTime(row.SLOT_BEG ?? '');
  const end = splitDateTime(row.SLOT_END ?? '');
  const slotName = cleanSpaces(row.SLOT_NAME ?? '');
  return {
    id,
    theatre: row.THEATRE ?? '',
    or: theatreToOr(row.THEATRE ?? ''),
    date: begin.date || expandShortDate(row.SCHED_DATE ?? ''),
    start: begin.time,
    end: end.time,
    ampm: begin.time && begin.time < '12:00' ? 'AM' : 'PM',
    slotName,
    slotOwner: slotName.includes(',') ? surnameOf(slotName) : '',
    consultant: '',
    patients: [],
  };
}

function extractPatient(row: Record<string, string>, id: string, options: ExtractOptions): Patient {
  const mrn = row.MRN ?? '';
  const name = cleanSpaces(row.PT_NAME ?? '');
  const ageSex = formatAgeSex(row.PT_AGE ?? '', row.PT_SEX ?? '');
  const procedureRaw = cleanSpaces(row.PLANNED_PROCEDURE ?? '');
  const eye = detectEye(procedureRaw);
  const comment = cleanSpaces(stripCommentTags(row.GENERAL_COMMENT ?? ''));
  const notes = parseComment(comment, options.hospitalCodes);
  const hospitalFlags = options.flagHospitals.filter((h) => mentionsWord(Object.values(row), h));
  const consultant = cleanSpaces(row.CONSULTANT ?? '');

  return {
    id,
    raw: row,
    time: normaliseTime(row.SCHED_TM ?? ''),
    mrn,
    name,
    ageSex,
    patient: [mrn, name, ageSex].filter(Boolean).join(' '),
    eye,
    procedure: abbreviateProcedure(procedureRaw, eye),
    procedureRaw,
    interpreter: notes.interpreter,
    diabetic: notes.diabetic,
    history: [notes.interpreter, notes.diabetic].filter(Boolean).join('\n'),
    allergy: formatAllergy(row.ALLERGY ?? ''),
    allergyPending: /^y(es)?$/i.test((row.ALLERGY ?? '').trim()),
    alerts: formatAlerts(row.CLINICAL_ALERT ?? ''),
    hospitalFlags,
    planNotes: [...hospitalFlags.map((h) => `${h} patient`), ...notes.leftover].join('\n'),
    comment,
    consultant,
    consultantSurname: surnameOf(consultant),
    ward: row.CURRENT_WARD ?? '',
    admissionType: cleanSpaces((row.INTEND_ADM_TYPE ?? '').replace(/\(.*?\)/g, '')),
    arrival: splitDateTime(row.TCI_TM || row.PLANNED_ADM_TM || '').time,
    priority: row.CLINICAL_PRIORITY ?? '',
  };
}

// ---- Eye and procedure -------------------------------------------------

export function detectEye(procedure: string): Eye {
  const text = procedure.toLowerCase();
  if (/\b(bilateral|bilat|both)\b/.test(text) || /\b(BE|OU)\b/.test(procedure)) return 'BOTH';
  const right = /\b(right|rt)\b/.test(text) || /\b(RE|OD)\b/.test(procedure);
  const left = /\b(left|lt)\b/.test(text) || /\b(LE|OS)\b/.test(procedure);
  if (right && left) return 'BOTH';
  if (right) return 'RIGHT';
  if (left) return 'LEFT';
  return '';
}

const EYE_PREFIX: Record<Eye, string> = { RIGHT: 'RE', LEFT: 'LE', BOTH: 'BE', '': '' };

interface ProcedureRule {
  test: RegExp;
  label: (text: string) => string;
  /** Words this rule accounts for, so they aren't repeated as "other" parts of the procedure. */
  words: string[];
}

/**
 * Each matching rule contributes a short label. Any words no rule accounts
 * for are kept and added on the end, so nothing on the booking is lost.
 */
export const PROCEDURE_RULES: ProcedureRule[] = [
  {
    test: /cataract|phaco/,
    label: (t) => {
      const lens = /toric/.test(t) ? 'toric IOL' : 'IOL';
      return /lens|\b(pc)?iol\b|intra\s*-?\s*ocular|toric/.test(t) ? `CE/${lens}` : 'CE';
    },
    words: ['cataract', 'cataracts', 'extraction', 'phaco', 'phacoemulsification', 'intraocular', 'intra', 'ocular',
      'lens', 'lense', 'implant', 'iol', 'pciol', 'toric'],
  },
  {
    test: /i\s*-?\s*stent/,
    label: (t) => (/inject/.test(t) ? 'iStent inject' : 'iStent'),
    words: ['istent', 'stent', 'stents', 'inject', 'trabecular', 'micro', 'bypass'],
  },
  { test: /hydrus/, label: () => 'Hydrus', words: ['hydrus', 'microstent', 'stent'] },
  {
    test: /trabeculectomy|\btrab\b/,
    label: (t) => (/mitom[iy]c[iy]n|\bmmc\b/.test(t) ? 'trabeculectomy + MMC' : 'trabeculectomy'),
    words: ['trabeculectomy', 'trab', 'mitomycin', 'mitomicyn', 'mitomycn', 'mitomicin', 'mmc'],
  },
  { test: /vitrectomy|\bppv\b/, label: () => 'PPV', words: ['vitrectomy', 'pars', 'plana', 'ppv'] },
  { test: /dacryocystorhinostomy|\bdcr\b/, label: () => 'DCR', words: ['dacryocystorhinostomy', 'dcr'] },
  { test: /ptosis/, label: () => 'ptosis repair', words: ['ptosis', 'repair', 'correction'] },
  {
    test: /pterygium/,
    label: (t) => (/graft/.test(t) ? 'pterygium excision + autograft' : 'pterygium excision'),
    words: ['pterygium', 'excision', 'autograft', 'graft', 'conjunctival', 'conj'],
  },
];

// Laterality and joining words, never worth repeating.
const FILLER_WORDS = ['left', 'right', 'lt', 'rt', 're', 'le', 'be', 'ou', 'od', 'os', 'eye', 'eyes', 'bilateral',
  'bilat', 'both', 'and', 'of', 'an', 'a', 'the', 'with', 'plus', 'insertion', 'insert', 'inserted', 'implantation'];

export function abbreviateProcedure(procedure: string, eye: Eye): string {
  const text = procedure.toLowerCase();
  const matched = PROCEDURE_RULES.filter((r) => r.test.test(text));
  if (matched.length === 0) return procedure;

  const known = new Set([...FILLER_WORDS, ...matched.flatMap((r) => r.words)]);
  const isKnown = (word: string) => word.length === 1 || /^\d+$/.test(word) || known.has(word.toLowerCase());
  const others = unexplainedPhrases(procedure, isKnown);

  const labels = [...matched.map((r) => r.label(text)), ...others];
  return [EYE_PREFIX[eye], labels.join(' + ')].filter(Boolean).join(' ');
}

/** Runs of consecutive words not accounted for, e.g. "under GA" from "... lens under GA". */
function unexplainedPhrases(procedure: string, isKnown: (word: string) => boolean): string[] {
  const phrases: string[] = [];
  let current: string[] = [];
  const flush = () => {
    if (current.length) phrases.push(current.join(' '));
    current = [];
  };
  // Separators like "+", "," or "/" also end a phrase.
  for (const token of procedure.split(/(\s+|[+,/;&()]+)/)) {
    if (!token.trim()) continue;
    if (/^[+,/;&()]+$/.test(token)) {
      flush();
      continue;
    }
    for (const word of token.split('-')) {
      if (!word) continue;
      if (isKnown(word)) flush();
      else current.push(word);
    }
  }
  flush();
  return phrases;
}

// ---- Comments ----------------------------------------------------------

const INTERPRETER = /\b(?:interpreter|interpretor|intepreter|interp|intep)\b/i;
const DIABETES = /diabet|daibet|\b(?:t[12]\s*dm|n?iddm|dm)\b/i;
const GLP1 = /\b(ozempic|semaglutide|mounjaro|tirzepatide|trulicity|dulaglutide|wegovy|saxenda)\b/i;
const NOT_A_LANGUAGE = new Set(['need', 'needs', 'require', 'requires', 'req', 'a', 'an', 'with', 'and', 'book', 'booked']);

export interface CommentNotes {
  interpreter: string;
  diabetic: string;
  leftover: string[];
}

/**
 * Splits a run-sheet comment into pieces and sorts them: interpreter and
 * diabetes notes go to the history column, bare hospital codes are dropped
 * (they are flagged separately), anything else is kept for Plan & notes.
 */
export function parseComment(comment: string, hospitalCodes: string[] = []): CommentNotes {
  const codes = new Set(hospitalCodes.map((c) => c.toUpperCase()));
  let interpreter = '';
  let diabetes = false;
  const meds: string[] = [];
  const leftover: string[] = [];

  for (const piece of comment.split(/[,;/\n]|\.(?=\s|$)/)) {
    const text = piece.trim();
    if (!text) continue;
    if (INTERPRETER.test(text)) {
      interpreter ||= describeInterpreter(text);
      continue;
    }
    const med = text.match(GLP1);
    if (DIABETES.test(text) || med) {
      if (DIABETES.test(text)) diabetes = true;
      if (med) meds.push(capitalise(med[1]));
      continue;
    }
    if (codes.has(text.replace(/\s+(pt|pts|patient)$/i, '').toUpperCase())) continue;
    leftover.push(text);
  }

  const diabetic = [diabetes ? 'Diabetic' : '', meds.length ? `(${meds.join(', ')})` : '']
    .filter(Boolean)
    .join(' ');
  return { interpreter, diabetic, leftover };
}

function describeInterpreter(text: string): string {
  const before = text.match(/([A-Za-z]+)\s+(?:interpreter|interpretor|intepreter|interp|intep)\b/i);
  const after = text.match(/(?:interpreter|interp)\s*[:-]\s*([A-Za-z]+)/i);
  const language = [before?.[1], after?.[1]].find((w) => w && !NOT_A_LANGUAGE.has(w.toLowerCase()));
  return language ? `${capitalise(language)} interpreter` : 'Interpreter';
}

function stripCommentTags(comment: string): string {
  return comment.replace(/\[[A-Z]\]\s*:\s*/g, ' ');
}

// ---- Other fields ------------------------------------------------------

export function formatAllergy(value: string): string {
  const v = value.trim();
  if (!v || /not\s*recorded/i.test(v)) return 'Allergy not recorded';
  if (/^(nka|nkda|nil( known)?|no known.*)$/i.test(v)) return 'NKA';
  if (/^y(es)?$/i.test(v)) return 'Allergy: Y';
  return `Allergy: ${v}`;
}

export function formatAlerts(value: string): string {
  const v = value.trim();
  if (/^(no|nil|none|n)?$/i.test(v)) return '';
  return v
    .split(';')
    .map((s) => s.trim())
    .filter(Boolean)
    .join('\n');
}

export function formatAgeSex(age: string, sex: string): string {
  const m = age.match(/(\d+)\s*([A-Za-z]*)/);
  const unit = (m?.[2] ?? '').toUpperCase();
  const suffix: Record<string, string> = { M: 'mo', W: 'wk', D: 'd' };
  const ageText = m ? m[1] + (suffix[unit.charAt(0)] ?? '') : '';
  return ageText + sex.trim().charAt(0).toUpperCase();
}

/** "Smith, J" → "Smith"; "Smith, J (Max)  Eye - General" → "Smith". */
export function surnameOf(name: string): string {
  return name.split(',')[0].trim();
}

export function theatreToOr(theatre: string): string {
  const m = theatre.match(/\bOR\s*0*(\d+)/i);
  return m ? `OR ${m[1]}` : theatre.trim();
}

function splitDateTime(value: string): { date: string; time: string } {
  const m = value.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}[:.]\d{2}))?/);
  if (!m) return { date: '', time: normaliseTime(value) };
  const year = m[3].length === 2 ? `20${m[3]}` : m[3];
  return { date: `${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}/${year}`, time: normaliseTime(m[4] ?? '') };
}

function expandShortDate(value: string): string {
  return splitDateTime(value).date;
}

function normaliseTime(value: string): string {
  const m = value.trim().match(/(\d{1,2})[:.](\d{2})/);
  return m ? `${m[1].padStart(2, '0')}:${m[2]}` : '';
}

function mentionsWord(values: string[], word: string): boolean {
  const pattern = new RegExp(`\\b${word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
  return values.some((v) => pattern.test(v));
}

function mostCommon(values: string[]): string {
  const counts = new Map<string, number>();
  for (const v of values) if (v) counts.set(v, (counts.get(v) ?? 0) + 1);
  let best = '';
  let bestCount = 0;
  for (const [v, n] of counts) {
    if (n > bestCount) [best, bestCount] = [v, n];
  }
  return best;
}

function cleanSpaces(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

function capitalise(word: string): string {
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}

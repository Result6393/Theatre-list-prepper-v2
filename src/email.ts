import type { Patient, Session } from './extract';

export interface EmailSettings {
  to: string[];
  cc: string[];
  /** Name used to sign off the message. */
  name: string;
  subject: string;
  message: string;
  /** Outlook on the web address that hosts the compose link. */
  outlookUrl: string;
}

export interface Email {
  to: string[];
  cc: string[];
  subject: string;
  body: string;
}

export const DEFAULT_SUBJECT = 'Files for surgery on {day}';

export const DEFAULT_MESSAGE = `Hello,

Could I please collect the files for the following patients having surgery on {day}?

{patients}

Thank you!
-{name}`;

export const DEFAULT_OUTLOOK_URL = 'https://outlook.cloud.microsoft/mail/';

/** Longest compose link we'll open; longer ones fall back to copying the text. */
export const MAX_URL_LENGTH = 1800;

export function defaultEmailSettings(): EmailSettings {
  return {
    to: [],
    cc: [],
    name: '',
    subject: DEFAULT_SUBJECT,
    message: DEFAULT_MESSAGE,
    outlookUrl: DEFAULT_OUTLOOK_URL,
  };
}

/** Loose check: something@something, no spaces. */
export function isEmailAddress(value: string): boolean {
  return /^[^\s@]+@[^\s@]+$/.test(value.trim());
}

export function validAddresses(list: string[]): string[] {
  return list.map((a) => a.trim()).filter(isEmailAddress);
}

/** "13/10/2026" → "Tuesday 13/10". Returns the input unchanged if it isn't a date. */
export function weekdayLabel(date: string): string {
  const m = date.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return date;
  const [day, month, year] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const when = new Date(Date.UTC(year, month - 1, day));
  if (when.getUTCDate() !== day || when.getUTCMonth() !== month - 1) return date;
  const weekday = when.toLocaleDateString('en-AU', { weekday: 'long', timeZone: 'UTC' });
  return `${weekday} ${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}`;
}

/** Replaces {placeholders}; unknown ones are left as typed. */
export function fillTemplate(template: string, values: Record<string, string>): string {
  return template.replace(/\{(\w+)\}/g, (whole, key: string) => values[key] ?? whole);
}

/** "A", "A and B", "A, B and C". */
export function joinList(items: string[]): string {
  return items.length < 2 ? (items[0] ?? '') : `${items.slice(0, -1).join(', ')} and ${items[items.length - 1]}`;
}

/** A theatre list and the patients from it to ask for. */
export interface EmailGroup {
  session: Session;
  patients: Patient[];
}

function unique(items: string[]): string[] {
  return [...new Set(items)];
}

const patientLine = (p: Patient) => `- ${[p.mrn, p.name].filter(Boolean).join(' ')}`;

/**
 * One email for one or more lists. With a single list the patients are listed plainly;
 * with several they're grouped under a heading per list (day shown only if the days differ).
 */
export function buildEmail(settings: EmailSettings, allGroups: EmailGroup[]): Email {
  const groups = allGroups.filter((g) => g.patients.length > 0);
  const dates = unique(groups.map((g) => g.session.date));
  const heading = (g: EmailGroup) => {
    const { session } = g;
    const day = dates.length > 1 ? `${weekdayLabel(session.date)} · ` : '';
    return `${day}${session.or} ${session.ampm}${session.consultant ? ` (${session.consultant})` : ''}`;
  };

  const values = {
    day: joinList(dates.map(weekdayLabel)),
    date: joinList(dates),
    theatre: joinList(unique(groups.map((g) => `${g.session.or} ${g.session.ampm}`))),
    count: String(groups.reduce((n, g) => n + g.patients.length, 0)),
    name: settings.name,
    patients:
      groups.length > 1
        ? groups.map((g) => [heading(g), ...g.patients.map(patientLine)].join('\n')).join('\n\n')
        : groups.flatMap((g) => g.patients.map(patientLine)).join('\n'),
  };
  return {
    to: validAddresses(settings.to),
    cc: validAddresses(settings.cc),
    subject: fillTemplate(settings.subject, values),
    body: fillTemplate(settings.message, values),
  };
}

/** Link that opens a draft in Outlook on the web. Nothing is sent until the user presses Send. */
export function outlookComposeUrl(baseUrl: string, email: Email): string {
  // Accept the address with or without a trailing "/mail/", as people copy it from the browser.
  const base = (baseUrl.trim() || DEFAULT_OUTLOOK_URL).replace(/\/+$/, '').replace(/\/mail$/i, '');
  const params = [
    `to=${encodeURIComponent(email.to.join(';'))}`,
    email.cc.length ? `cc=${encodeURIComponent(email.cc.join(';'))}` : '',
    `subject=${encodeURIComponent(email.subject)}`,
    `body=${encodeURIComponent(email.body.replace(/\r?\n/g, '\r\n'))}`,
  ].filter(Boolean);
  return `${base}/mail/deeplink/compose?${params.join('&')}`;
}

/** `mailto:` link that opens a draft in the computer's default email app (e.g. installed Outlook). */
export function mailtoUrl(email: Email): string {
  const params = [
    email.cc.length ? `cc=${encodeURIComponent(email.cc.join(','))}` : '',
    `subject=${encodeURIComponent(email.subject)}`,
    `body=${encodeURIComponent(email.body.replace(/\r?\n/g, '\r\n'))}`,
  ].filter(Boolean);
  return `mailto:${email.to.map((a) => encodeURIComponent(a).replace(/%40/g, '@')).join(',')}?${params.join('&')}`;
}

/** Just the message, for pasting into an email you've started yourself (no recipients or subject). */
export function emailBodyText(email: Email): string {
  return email.body.replace(/\r\n?/g, '\n');
}

/**
 * The message as HTML with one block per line and an empty block for each blank line, so a
 * paste into Outlook keeps full blank lines. No styles, so Outlook uses your default font.
 */
export function emailBodyHtml(email: Email): string {
  const escape = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return emailBodyText(email)
    .split('\n')
    .map((line) => (line.trim() ? `<div>${escape(line)}</div>` : '<div><br></div>'))
    .join('');
}

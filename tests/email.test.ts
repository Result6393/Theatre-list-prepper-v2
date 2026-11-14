import { describe, expect, it } from 'vitest';
import {
  buildEmail, defaultEmailSettings, emailBodyHtml, emailBodyText, fillTemplate, isEmailAddress, joinList, mailtoUrl, outlookComposeUrl,
  weekdayLabel,
} from '../src/email';
import { extractSessions, type Patient, type Session } from '../src/extract';
import { parseTable } from '../src/parse';
import { defaultSettings, exportSettings, importSettings } from '../src/settings';
import { nextDayRows, ROWS, runSheetCsv } from './fixtures';

const session = extractSessions(parseTable(runSheetCsv()), { flagHospitals: [], hospitalCodes: [] }).find(
  (s) => s.or === 'OR 1' && s.ampm === 'AM',
)!;

const settings = { ...defaultEmailSettings(), to: ['records@example.org'], name: 'Sam' };

describe('weekdayLabel', () => {
  it('adds the weekday', () => {
    expect(weekdayLabel('13/10/2026')).toBe('Tuesday 13/10');
    expect(weekdayLabel('1/11/2026')).toBe('Sunday 01/11');
  });
  it('leaves non-dates and impossible dates alone', () => {
    expect(weekdayLabel('soon')).toBe('soon');
    expect(weekdayLabel('31/02/2026')).toBe('31/02/2026');
  });
});

describe('buildEmail', () => {
  it('fills the default wording with the given patients, without age or sex', () => {
    const email = buildEmail(settings, [{ session, patients: session.patients }]);
    expect(email.to).toEqual(['records@example.org']);
    expect(email.subject).toBe('Files for surgery on Wednesday 14/10');
    expect(email.body).toBe(
      `Hello,

Could I please collect the files for the following patients having surgery on Wednesday 14/10?

- 1000002 O'SAMPLE, Bob
- 1000001 TESTER, Alice Jane

Thank you!
-Sam`,
    );
  });

  it('only includes the patients passed in, and ignores invalid addresses', () => {
    const email = buildEmail({ ...settings, to: ['bad address', 'ok@example.org'], cc: ['x'] }, [{ session, patients: [session.patients[1]] }]);
    expect(email.body).toContain('- 1000001 TESTER, Alice Jane');
    expect(email.body).not.toContain('1000002');
    expect(email.to).toEqual(['ok@example.org']);
    expect(email.cc).toEqual([]);
  });

  it('supports the other placeholders and leaves unknown ones alone', () => {
    expect(fillTemplate('{theatre} {count} {date} {nope}', { theatre: 'OR 1 AM', count: '2', date: '14/10/2026' })).toBe(
      'OR 1 AM 2 14/10/2026 {nope}',
    );
  });
});

describe('several lists', () => {
  const all = extractSessions(parseTable(runSheetCsv([...ROWS, ...nextDayRows()])), { flagHospitals: [], hospitalCodes: [] });
  const find = (date: string, or: string, ampm: string) => all.find((s) => s.date === date && s.or === or && s.ampm === ampm)!;
  const sameDayA = find('14/10/2026', 'OR 1', 'AM');
  const sameDayB = find('14/10/2026', 'OR 2', 'AM');
  const nextDay = find('15/10/2026', 'OR 1', 'AM');

  it('keeps the same list on different days separate and in date order', () => {
    expect(all.filter((s) => s.or === 'OR 1' && s.ampm === 'AM').map((s) => s.date)).toEqual(['14/10/2026', '15/10/2026']);
    expect(all[all.length - 1].date).toBe('15/10/2026');
    expect(all.filter((s) => s.date === '14/10/2026').length).toBe(all.length - 1);
  });

  it('groups patients under a heading per list on the same day', () => {
    const email = buildEmail(settings, [
      { session: sameDayA, patients: sameDayA.patients },
      { session: sameDayB, patients: sameDayB.patients },
    ]);
    expect(email.body).toBe(
      `Hello,

Could I please collect the files for the following patients having surgery on Wednesday 14/10?

OR 1 AM (Cedar)
- 1000002 O'SAMPLE, Bob
- 1000001 TESTER, Alice Jane

OR 2 AM (Alder)
- 1000005 PLACEHOLDER, Eve

Thank you!
-Sam`,
    );
  });

  it('names every day and adds the day to headings when days differ', () => {
    const email = buildEmail(settings, [
      { session: sameDayA, patients: [sameDayA.patients[0]] },
      { session: nextDay, patients: nextDay.patients },
    ]);
    expect(email.subject).toBe('Files for surgery on Wednesday 14/10 and Thursday 15/10');
    expect(email.body).toContain('Wednesday 14/10 · OR 1 AM (Cedar)\n- 1000002 O\'SAMPLE, Bob');
    expect(email.body).toContain('Thursday 15/10 · OR 1 AM (Cedar)\n- 2000001 NEXTDAY, Nina');
  });

  it('skips lists with nothing ticked, leaving a plain single-list email', () => {
    const email = buildEmail(settings, [
      { session: sameDayA, patients: [sameDayA.patients[0]] },
      { session: sameDayB, patients: [] },
    ]);
    expect(email.body).toContain('\n\n- 1000002 O\'SAMPLE, Bob\n\n');
    expect(email.body).not.toContain('OR 2');
  });

  it('joins lists of days', () => {
    expect([[], ['a'], ['a', 'b'], ['a', 'b', 'c']].map(joinList)).toEqual(['', 'a', 'a and b', 'a, b and c']);
  });
});

describe('outlookComposeUrl', () => {
  const email = { to: ['a@x.org', 'b@x.org'], cc: ['c@x.org'], subject: 'R&D #1', body: "Line 1\nO'Brien & co" };

  it('encodes recipients, subject and line breaks', () => {
    const url = new URL(outlookComposeUrl('https://outlook.office.com/', email));
    expect(url.origin + url.pathname).toBe('https://outlook.office.com/mail/deeplink/compose');
    for (const base of ['https://outlook.cloud.microsoft/mail/', 'https://outlook.cloud.microsoft/mail', 'https://outlook.cloud.microsoft/']) {
      expect(outlookComposeUrl(base, email).startsWith('https://outlook.cloud.microsoft/mail/deeplink/compose?')).toBe(true);
    }
    expect(url.searchParams.get('to')).toBe('a@x.org;b@x.org');
    expect(url.searchParams.get('cc')).toBe('c@x.org');
    expect(url.searchParams.get('subject')).toBe('R&D #1');
    expect(url.searchParams.get('body')).toBe("Line 1\r\nO'Brien & co");
  });

  it('omits cc when empty and falls back to the default address', () => {
    const url = outlookComposeUrl('', { ...email, cc: [] });
    expect(url.startsWith('https://outlook.cloud.microsoft/mail/deeplink/compose?')).toBe(true);
    expect(url).not.toContain('cc=');
  });
});

describe('mailtoUrl', () => {
  it('encodes recipients, subject and line breaks', () => {
    const url = mailtoUrl({ to: ['a@x.org', 'b@x.org'], cc: ['c@x.org'], subject: 'R&D #1', body: "L1\nO'Brien" });
    expect(url.startsWith('mailto:a@x.org,b@x.org?')).toBe(true);
    const params = new URLSearchParams(url.slice(url.indexOf('?') + 1));
    expect(params.get('cc')).toBe('c@x.org');
    expect(params.get('subject')).toBe('R&D #1');
    expect(params.get('body')).toBe("L1\r\nO'Brien");
  });
});

describe('misc', () => {
  it('copies only the message: no recipients or subject', () => {
    const email = { to: ['a@x.org'], cc: ['c@x.org'], subject: 'Subj', body: 'Hello,\r\n\r\nBody' };
    expect(emailBodyText(email)).toBe('Hello,\n\nBody');
  });
  it('keeps a full blank line after Hello when pasted as HTML', () => {
    const email = buildEmail(settings, [{ session, patients: session.patients }]);
    const html = emailBodyHtml(email);
    expect(html.startsWith('<div>Hello,</div><div><br></div><div>Could I please collect')).toBe(true);
    expect(html).not.toContain('records@example.org');
    expect(html).not.toContain('Files for surgery');
    expect(emailBodyHtml({ ...email, body: 'a <b> & c' })).toBe('<div>a &lt;b&gt; &amp; c</div>');
  });
  it('checks addresses loosely', () => {
    expect([' a@b.org ', 'a b@c.org', 'nope'].map(isEmailAddress)).toEqual([true, false, false]);
  });
  it('keeps email settings through export/import and fills them in for older files', () => {
    const s = defaultSettings();
    s.email.to = ['records@example.org'];
    s.email.name = 'Sam';
    expect(importSettings(exportSettings(s)).email).toEqual(s.email);
    expect(importSettings('{"flagHospitals":["X"]}').email).toEqual(defaultEmailSettings());
  });
});

describe('an email for a list from your prepped sheet, with no run sheet', () => {
  const sheetPatient = (mrn: string, name: string) => ({ mrn, name }) as Patient;
  const session = (over: Partial<Session>) => ({ date: '', or: '', ampm: '', consultant: '', ...over }) as Session;

  it('uses the sheet’s date and theatre, or says “this list” when it has neither', () => {
    const withHeading = buildEmail(defaultEmailSettings(), [{ session: session({ date: '14/10/2026', or: 'OR 4', ampm: 'AM' }), patients: [sheetPatient('3000001', 'ALPHA, Ann')] }]);
    expect(withHeading.subject).toBe('Files for surgery on Wednesday 14/10');
    expect(withHeading.body).toContain('- 3000001 ALPHA, Ann');
    const noDate = buildEmail(defaultEmailSettings(), [{ session: session({ or: 'OR 4', ampm: 'AM' }), patients: [sheetPatient('3000001', 'ALPHA, Ann')] }]);
    expect(noDate.subject).toBe('Files for surgery on OR 4 AM');
    const nothing = buildEmail(defaultEmailSettings(), [{ session: session({}), patients: [sheetPatient('3000001', 'ALPHA, Ann')] }]);
    expect(nothing.subject).toBe('Files for surgery on this list');
  });
});

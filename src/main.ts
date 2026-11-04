import './style.css';
import { copyHtml, copyText, htmlTableRows, toTsv } from './clipboard';
import { autoGrow, h } from './dom';
import { buildEmail, emailBodyHtml, emailBodyText, mailtoUrl, MAX_URL_LENGTH, outlookComposeUrl, weekdayLabel } from './email';
import { extractSessions, type Patient, type Session } from './extract';
import { copyBlock, freshRows, isRectangular, mergePrepped, MergeError, parsePastedRows, withoutTrailingBlankRows, type GridRow, type MergeSummary } from './merge';
import { ParseError, parseTable, type Table } from './parse';
import { classifyFile } from './files';
import { defaultSettings, exportSettings, importSettings, loadSettings, saveSettings, STORAGE_KEY, SYNC_APP, type KeyValueStore } from './settings';
import { mountSyncControl, pullBeforeStart, pushSoon } from './sync.js';
import { renderSettings } from './settings-view';

interface Grid {
  session: Session;
  rows: GridRow[];
  /** After an update from your sheet: how many rows the copy must cover to replace the old ones. */
  padTo: number;
  summary: MergeSummary | null;
  /** The list as it was before the last update from your sheet. */
  undo: { rows: GridRow[]; padTo: number; summary: MergeSummary | null } | null;
}

const storage: KeyValueStore | null = (() => {
  try {
    return window.localStorage;
  } catch {
    return null;
  }
})();

const state = {
  settings: loadSettings(storage),
  view: 'main' as 'main' | 'settings',
  /** Settings as they were when the Settings page opened, to tell if the list needs rebuilding. */
  settingsSnapshot: '',
  inputText: '',
  table: null as Table | null,
  sessions: [] as Session[],
  error: '',
  sessionId: '',
  grid: null as Grid | null,
  /** Each list's grid, kept so edits and unticks survive switching between lists. */
  grids: new Map<string, Grid>(),
  /** Other lists, besides the one in view, to include in the Medical Records email. */
  emailLists: new Set<string>(),
  status: '',
  /** The "Update from my sheet" paste box is open. */
  updateOpen: false,
  updateError: '',
  /** A message under the header, e.g. after importing a settings file. */
  notice: null as { text: string; error: boolean } | null,
};

const app = document.getElementById('app')!;

/** Saves settings on this device and, if signed in to sync, to your other devices. */
function persist(): void {
  saveSettings(storage, state.settings);
  pushSoon(SYNC_APP, STORAGE_KEY);
}

/** Reloads settings after sync brought newer ones from another device. */
function applySyncedSettings(): void {
  state.settings = loadSettings(storage);
  state.notice = { text: 'Settings updated from your other devices.', error: false };
  if (state.view === 'settings') state.settingsSnapshot = exportSettings(state.settings);
  else if (state.table) refreshSessions();
  render();
}

let unmountSync: (() => void) | null = null;

// ---- State changes ------------------------------------------------------------

function listedSessions(): Session[] {
  return state.sessions.filter((s) => s.patients.length > 0);
}

function loadRunSheet(text: string): void {
  state.status = '';
  state.inputText = text;
  if (!text.trim()) {
    state.table = null;
    state.sessions = [];
    state.error = '';
  } else {
    try {
      state.table = parseTable(text);
      state.error = '';
    } catch (e) {
      state.table = null;
      state.error = e instanceof ParseError ? e.message : 'Couldn’t read that run sheet.';
    }
  }
  state.sessionId = '';
  refreshSessions();
  render();
}

/** Re-extracts after the run sheet or settings change, opening the list straight away if there is only one. */
function refreshSessions(): void {
  const { flagHospitals, hospitalCodes } = state.settings;
  state.grids.clear();
  state.emailLists.clear();
  state.sessions = state.table ? extractSessions(state.table, { flagHospitals, hospitalCodes }) : [];
  const lists = listedSessions();
  if (!lists.some((s) => s.id === state.sessionId)) {
    state.sessionId = lists.length === 1 ? lists[0].id : '';
  }
  openSession();
}

function gridFor(session: Session): Grid {
  let grid = state.grids.get(session.id);
  if (!grid) {
    grid = { session, rows: freshRows(session.patients, state.settings.layout), padTo: 0, summary: null, undo: null };
    state.grids.set(session.id, grid);
  }
  return grid;
}

function openSession(): void {
  const session = state.sessions.find((s) => s.id === state.sessionId);
  state.grid = session ? gridFor(session) : null;
}

function chooseSession(id: string): void {
  state.sessionId = id;
  state.status = '';
  state.updateOpen = false;
  state.updateError = '';
  openSession();
  render();
}

/** The lists in the email: the one in view plus any others ticked, in run-sheet order. */
function emailGrids(): Grid[] {
  return listedSessions()
    .filter((s) => s.id === state.sessionId || state.emailLists.has(s.id))
    .map(gridFor);
}

/** Ticked patients still on the run sheet (not cancelled). */
function tickedPatients(grid: Grid): Patient[] {
  return grid.rows.filter((r) => r.included && r.status !== 'cancelled' && r.patient).map((r) => r.patient!);
}

function emailCount(): number {
  return emailGrids().reduce((n, g) => n + tickedPatients(g).length, 0);
}

function medicalRecordsEmail() {
  const groups = emailGrids().map((g) => ({ session: g.session, patients: tickedPatients(g) }));
  return buildEmail(state.settings.email, groups);
}

function emailSummary(): string {
  const count = emailCount();
  const lists = emailGrids().filter((g) => tickedPatients(g).length > 0).length;
  return `${count} patient${count === 1 ? '' : 's'}${lists > 1 ? ` from ${lists} lists` : ''}`;
}

/** Opens a draft in Outlook on the web; copies the text instead if the link is too long or blocked. */
async function emailMedicalRecords(): Promise<void> {
  const email = medicalRecordsEmail();
  const url = outlookComposeUrl(state.settings.email.outlookUrl, email);
  const tooLong = url.length > MAX_URL_LENGTH;
  // 'noopener' would make window.open return null even on success, so detach the new tab by hand.
  const tab = tooLong ? null : window.open(url, '_blank');
  if (!tab) {
    await copyEmailText();
    return;
  }
  tab.opener = null;
  state.status = `Opened a draft email with ${emailSummary()}. Check it and press Send in Outlook.`;
  render();
}

/** Hands the draft to the computer's default email app via a mailto link, without leaving this page. */
async function emailInApp(): Promise<void> {
  const url = mailtoUrl(medicalRecordsEmail());
  if (url.length > MAX_URL_LENGTH) {
    await copyEmailText();
    return;
  }
  const summary = emailSummary();
  // A hidden frame keeps this page in place even if no email app handles the link.
  const frame = h('iframe', { src: url, hidden: true, 'aria-hidden': 'true' });
  document.body.append(frame);
  setTimeout(() => frame.remove(), 2000);

  // When an email app opens, this window loses focus. If it doesn't, say so.
  let handedOver = false;
  const noteHandover = () => {
    handedOver = true;
  };
  window.addEventListener('blur', noteHandover, { once: true });
  state.status = `Asked your email app to open a draft with ${summary}…`;
  render();
  setTimeout(() => {
    window.removeEventListener('blur', noteHandover);
    state.status = handedOver
      ? `Opened a draft with ${summary}. Check it and press Send.`
      : 'No email app seems to have opened. This computer may not have a default email app set up: use Outlook on the web or Copy email text instead.';
    render();
  }, 1500);
}

const canShare = typeof navigator !== 'undefined' && typeof navigator.share === 'function';

/** Phone/tablet share sheet. Recipients can't be passed this way, so they're added in the email app. */
async function shareEmail(): Promise<void> {
  const email = medicalRecordsEmail();
  try {
    await navigator.share({ title: email.subject, text: email.body });
    state.status = 'Shared. Add the Medical Records address in your email app before sending.';
  } catch (e) {
    if (e instanceof DOMException && e.name === 'AbortError') return;
    await copyEmailText();
    return;
  }
  render();
}

async function copyEmailText(): Promise<void> {
  try {
    const email = medicalRecordsEmail();
    await copyHtml(emailBodyHtml(email), emailBodyText(email));
    state.status = `Copied the message for ${emailSummary()} (without the recipients or subject). Paste it into the body of a new email.`;
  } catch {
    state.status = 'Your browser blocked copying the email text.';
  }
  render();
}

/** Handles a file dropped anywhere on the page: settings (.json) or a run sheet. */
async function handleDroppedFile(file: File): Promise<void> {
  if (classifyFile(file.name, file.type) === 'runsheet') {
    state.notice = null;
    if (state.view === 'settings') closeSettings();
    loadRunSheet(await file.text());
    return;
  }
  try {
    const imported = importSettings(await file.text());
    const hasCustom = exportSettings(state.settings) !== exportSettings(defaultSettings());
    if (hasCustom && !confirm('Replace your current settings with the ones in this file?')) return;
    state.settings = imported;
    persist();
    state.notice = { text: `Settings imported from ${file.name}.`, error: false };
    if (state.view === 'main') refreshSessions();
  } catch {
    state.notice = { text: `${file.name} isn’t a Theatre List Prepper settings file, so nothing was changed.`, error: true };
  }
  render();
}

function openSettings(): void {
  state.settingsSnapshot = exportSettings(state.settings);
  state.view = 'settings';
  render();
  window.scrollTo(0, 0);
}

/** Back to the list; rebuilds it (losing edits) only if the settings actually changed. */
function closeSettings(): void {
  state.view = 'main';
  if (exportSettings(state.settings) !== state.settingsSnapshot) {
    state.status = '';
    refreshSessions();
  }
  render();
}

/** Moves a row within the list; cancelled rows stay together at the bottom. */
function moveRow(grid: Grid, from: number, to: number): void {
  if (to < 0 || to >= activeRows(grid).length) return;
  const [row] = grid.rows.splice(from, 1);
  grid.rows.splice(to, 0, row);
  render();
}

function activeRows(grid: Grid): GridRow[] {
  return grid.rows.filter((r) => r.status !== 'cancelled');
}

function cancelledRows(grid: Grid): GridRow[] {
  return grid.rows.filter((r) => r.status === 'cancelled');
}

/** "OR 3 · PM · Smith", with the day when the run sheet covers several. */
function listName(s: Session): string {
  const day = sheetDates().length > 1 && s.date ? `${shortDay(s.date)} · ` : '';
  return `${day}${[s.or, s.ampm, s.consultant || s.slotName].filter(Boolean).join(' · ')}`;
}

/** Merges rows pasted from your cheat sheet into this list, keeping your notes. */
function updateFromSheet(grid: Grid, text: string, html: string): void {
  let pasted = parsePastedRows(text);
  if (html && (pasted.length === 0 || !isRectangular(pasted))) {
    const fromHtml = htmlTableRows(html);
    if (fromHtml?.length) pasted = withoutTrailingBlankRows(fromHtml);
  }
  if (pasted.length === 0) return;
  try {
    const result = mergePrepped(pasted, grid.session, listedSessions(), state.settings.layout, listName);
    grid.undo = { rows: grid.rows, padTo: grid.padTo, summary: grid.summary };
    grid.rows = result.rows;
    grid.padTo = result.padTo;
    grid.summary = result.summary;
    state.updateOpen = false;
    state.updateError = '';
    state.status = '';
  } catch (e) {
    state.updateError = e instanceof MergeError ? e.message : 'Couldn’t read those rows. Copy them from your sheet and paste again.';
  }
  render();
}

function undoUpdate(grid: Grid): void {
  if (!grid.undo) return;
  Object.assign(grid, grid.undo, { undo: null });
  state.status = 'Update undone.';
  render();
}

function summaryText(s: MergeSummary): string {
  const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? '' : 's'}`;
  return [
    `kept your notes for ${plural(s.kept, 'patient')}`,
    s.added ? `${s.added} new (need prep)` : '',
    s.cancelled ? `${s.cancelled} cancelled` : '',
    s.moved ? `${s.moved} moved` : '',
    s.changed ? `${s.changed} with run-sheet changes to check` : '',
  ]
    .filter(Boolean)
    .join(' · ');
}

async function copyRows(grid: Grid): Promise<void> {
  const { layout } = state.settings;
  const rows = copyBlock(grid.rows, layout, grid.padTo);
  if (layout.includeHeadings) rows.unshift(layout.columns.map((c) => c.heading));
  try {
    await copyText(toTsv(rows));
    const count = activeRows(grid).filter((r) => r.included).length;
    const cancelled = cancelledRows(grid).filter((r) => r.included).length;
    const where = layout.pasteStart || 'the first cell';
    state.status = grid.padTo
      ? `Copied ${count} patient${count === 1 ? '' : 's'}${cancelled ? ` and ${cancelled} cancelled` : ''}. Click ${where} (where you copied your rows from) and press Ctrl+V to replace the old rows.`
      : `Copied ${count} patient${count === 1 ? '' : 's'}. Click ${where} in your sheet and press Ctrl+V.`;
  } catch {
    state.status = 'Your browser blocked copying. Try again, or use Ctrl+C after selecting the table.';
  }
  render();
}

// ---- Rendering ----------------------------------------------------------------

function render(): void {
  const scroll = window.scrollY;
  app.replaceChildren(
    h(
      'header',
      { class: 'topbar' },
      h('h1', {}, 'Theatre List Prepper ', h('span', { class: 'version' }, `v${__APP_VERSION__}`)),
      h(
        'button',
        { class: 'ghost', onclick: () => (state.view === 'settings' ? closeSettings() : openSettings()) },
        state.view === 'settings' ? 'Done' : 'Settings',
      ),
    ),
    ...(state.notice
      ? [h(
          'div',
          { class: state.notice.error ? 'notice error-notice' : 'notice', role: state.notice.error ? 'alert' : 'status' },
          state.notice.text,
          h('button', { class: 'link', onclick: () => ((state.notice = null), render()) }, 'Dismiss'),
        )]
      : []),
    state.view === 'settings'
      ? withSyncPanel(renderSettings({
          settings: state.settings,
          table: state.table,
          save: () => persist(),
          rerender: render,
          replace: (settings) => {
            state.settings = settings;
            persist();
            render();
          },
          done: closeSettings,
        }))
      : h('main', {}, renderInputStep(), renderListStep(), renderGrid()),
  );
  for (const area of app.querySelectorAll<HTMLTextAreaElement>('textarea.cell')) autoGrow(area);
  window.scrollTo(0, scroll);
}

/** Puts the sync sign-in panel at the top of the Settings page. */
function withSyncPanel(view: HTMLElement): HTMLElement {
  unmountSync?.();
  const slot = h('div', { class: 'sync-slot' });
  view.prepend(slot);
  unmountSync = mountSyncControl(slot, SYNC_APP, STORAGE_KEY, applySyncedSettings);
  return view;
}

function step(n: number, title: string, ...body: (Node | string | null | false)[]): HTMLElement {
  return h('section', { class: 'step' }, h('h2', {}, h('span', { class: 'num' }, String(n)), title), ...body);
}

function renderInputStep(): HTMLElement {
  const fileInput = h('input', {
    type: 'file',
    accept: '.csv,.tsv,.txt,text/csv',
    hidden: true,
    onchange: async (e: Event) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (file) loadRunSheet(await file.text());
    },
  });

  let body: (Node | string)[];
  if (state.table) {
    const patients = state.sessions.reduce((n, s) => n + s.patients.length, 0);
    const lists = state.sessions.filter((s) => s.patients.length > 0).length;
    const dates = sheetDates();
    const when = dates.length === 1 ? ` for ${dates[0]}` : dates.length > 1 ? ` across ${dates.length} days (${dates[0]} to ${dates[dates.length - 1]})` : '';
    body = [
      h(
        'p',
        { class: 'loaded' },
        `Run sheet loaded${when}: ${patients} patients on ${lists} lists. `,
        h('button', { class: 'link', onclick: () => loadRunSheet('') }, 'Clear'),
      ),
    ];
  } else {
    body = [
      h('textarea', {
        class: 'paste',
        rows: 5,
        value: state.inputText,
        placeholder: 'Copy the whole run sheet (including its heading row) from Excel and paste it here with Ctrl+V…',
        oninput: (e: Event) => loadRunSheet((e.target as HTMLTextAreaElement).value),
      }),
      h(
        'p',
        { class: 'hint' },
        'Or ',
        h('button', { class: 'link', onclick: () => fileInput.click() }, 'choose the CSV file'),
        ' or drop the CSV, or a settings file, anywhere on this page.',
      ),
      fileInput,
    ];
    if (state.error) {
      body.push(
        h('p', { class: 'error', role: 'alert' }, state.error),
        h('p', { class: 'hint' }, 'Pasting rows from your cheat sheet to update a list? Load the latest run sheet here first and choose the list, then use “Update from my sheet”.'),
      );
    }
  }

  return step(1, 'Run sheet', ...body);
}

/** Distinct dates on the run sheet, in order. */
function sheetDates(): string[] {
  return [...new Set(listedSessions().map((s) => s.date).filter(Boolean))];
}

/** "Tue 13/10" */
function shortDay(date: string): string {
  return weekdayLabel(date).replace(/^(\w{3})\w*/, '$1');
}

function sessionLabel(s: Session): string {
  const count = `${s.patients.length} patient${s.patients.length === 1 ? '' : 's'}`;
  return [s.or, s.ampm, s.consultant || s.slotName, count].filter(Boolean).join(' · ');
}

function renderListStep(): HTMLElement | null {
  if (!state.table) return null;
  const lists = listedSessions();
  if (lists.length === 0) {
    return step(2, 'Which list?', h('p', { class: 'hint' }, 'No lists with patients were found on this run sheet.'));
  }
  const button = (s: Session) =>
    h('button', { class: s.id === state.sessionId ? 'choice selected' : 'choice', onclick: () => chooseSession(s.id) }, sessionLabel(s));
  const dates = sheetDates();
  return step(
    2,
    lists.length === 1 ? 'List' : 'Which list are you prepping?',
    ...(dates.length > 1
      ? dates.map((date) =>
          h(
            'div',
            { class: 'day-group' },
            h('div', { class: 'day-heading' }, weekdayLabel(date)),
            h('div', { class: 'choices' }, lists.filter((s) => s.date === date).map(button)),
          ),
        )
      : [h('div', { class: 'choices' }, lists.map(button))]),
  );
}

/** Tick other lists to put in the same Medical Records email as the one in view. */
function renderEmailLists(): HTMLElement | null {
  const lists = listedSessions();
  if (lists.length < 2) return null;
  const multiDay = sheetDates().length > 1;
  return h(
    'fieldset',
    { class: 'email-lists' },
    h('legend', {}, 'Include in the Medical Records email'),
    lists.map((s) => {
      const current = s.id === state.sessionId;
      const ticked = tickedPatients(gridFor(s)).length;
      const label = `${multiDay ? `${shortDay(s.date)} · ` : ''}${s.or} · ${s.ampm} · ${s.consultant || s.slotName} · ${ticked} of ${s.patients.length} ticked${current ? ' (this list)' : ''}`;
      return h(
        'label',
        {},
        h('input', {
          type: 'checkbox',
          checked: current || state.emailLists.has(s.id),
          disabled: current,
          onchange: (e: Event) => {
            if ((e.target as HTMLInputElement).checked) state.emailLists.add(s.id);
            else state.emailLists.delete(s.id);
            render();
          },
        }),
        ` ${label}`,
      );
    }),
  );
}

function renderGrid(): HTMLElement | null {
  const grid = state.grid;
  if (!grid) return null;
  const { session } = grid;
  const { layout } = state.settings;

  const headerFacts: [string, string][] = [
    ['Consultant', session.consultant],
    ['Date', session.date],
    ['OR', session.or],
    ['Session', session.ampm],
  ];

  const facts = h(
    'div',
    { class: 'facts' },
    headerFacts.map(([label, value]) =>
      h(
        'div',
        { class: 'fact' },
        h('span', { class: 'fact-label' }, label),
        h('span', { class: 'fact-value' }, value || '—'),
        h(
          'button',
          {
            class: 'mini',
            title: `Copy ${label.toLowerCase()}`,
            onclick: async () => {
              await copyText(value).catch(() => undefined);
              state.status = `Copied ${label.toLowerCase()} “${value}”.`;
              render();
            },
          },
          'Copy',
        ),
      ),
    ),
  );

  const thead = h(
    'thead',
    {},
    h(
      'tr',
      {},
      h('th', { class: 'tick', title: 'Include in copy' }, '✓'),
      h('th', { class: 'move' }, ''),
      h('th', { class: 'time' }, 'Time'),
      layout.columns.map((c) => h('th', {}, c.heading)),
    ),
  );

  const active = activeRows(grid);
  const cancelled = cancelledRows(grid);
  const keyIndex = layout.columns.findIndex((c) => c.source === 'field:patient' || c.source === 'field:name' || c.source === 'field:mrn');
  const rowName = (r: GridRow) => r.patient?.name || (keyIndex >= 0 ? r.cells[keyIndex] : '') || 'this patient';

  const renderRow = (r: GridRow, index: number) => {
    const p = r.patient;
    const isCancelled = r.status === 'cancelled';
    const badges = [
      r.status === 'new' ? h('span', { class: 'badge badge-new', title: 'Added to the list since your sheet: needs prep' }, 'New') : null,
      r.reinstated ? h('span', { class: 'badge badge-new', title: 'Was cancelled on your sheet and is back on the run sheet' }, 'Back on list') : null,
      r.movedFrom !== undefined ? h('span', { class: 'badge badge-moved', title: 'The order changed: this patient has moved' }, `Moved, was #${r.movedFrom}`) : null,
      isCancelled ? h('span', { class: 'badge badge-cancelled' }, 'Cancelled') : null,
      r.note ? h('div', { class: 'who' }, r.note) : null,
    ];
    const classes = [r.included ? '' : 'excluded', r.status === 'new' || r.reinstated ? 'row-new' : '', isCancelled ? 'row-cancelled' : ''];
    return h(
      'tr',
      { class: classes.filter(Boolean).join(' ') },
      h(
        'td',
        { class: 'tick' },
        h('input', {
          type: 'checkbox',
          checked: r.included,
          'aria-label': `Include ${rowName(r)}`,
          onchange: (e: Event) => {
            r.included = (e.target as HTMLInputElement).checked;
            render();
          },
        }),
      ),
      h(
        'td',
        { class: 'move' },
        isCancelled
          ? null
          : [
              h('button', { class: 'mini', title: 'Move up', disabled: index === 0, onclick: () => moveRow(grid, index, index - 1) }, '↑'),
              h(
                'button',
                { class: 'mini', title: 'Move down', disabled: index === active.length - 1, onclick: () => moveRow(grid, index, index + 1) },
                '↓',
              ),
            ],
      ),
      h('td', { class: 'time', title: p?.consultant ?? '' }, p?.time ?? '', p ? h('div', { class: 'who' }, p.consultantSurname) : null, ...badges),
      layout.columns.map((c, col) => {
        const area = h('textarea', {
          class: 'cell',
          rows: 1,
          value: r.cells[col],
          'aria-label': `${c.heading} for ${rowName(r)}`,
          oninput: (e: Event) => {
            const target = e.target as HTMLTextAreaElement;
            r.cells[col] = target.value;
            target.parentElement!.className = cellClass(c.source, target.value);
            autoGrow(target);
          },
        });
        const latest = r.changes[col];
        const change =
          latest === undefined
            ? null
            : h(
                'div',
                { class: 'change', title: 'The run sheet now says something different from your sheet' },
                'Now: ',
                h('strong', {}, latest),
                h(
                  'div',
                  {},
                  h('button', { class: 'mini', title: 'Use the run sheet’s value', onclick: () => ((r.cells[col] = latest), delete r.changes[col], render()) }, 'Use'),
                  h('button', { class: 'mini', title: 'Keep what’s on your sheet', onclick: () => (delete r.changes[col], render()) }, 'Keep'),
                ),
              );
        return h('td', { class: cellClass(c.source, r.cells[col]), style: `min-width:${columnWidth(c.source)}px` }, area, change);
      }),
    );
  };

  const tbody = h(
    'tbody',
    {},
    active.map(renderRow),
    cancelled.length
      ? h(
          'tr',
          { class: 'cancelled-heading' },
          h('td', { colspan: layout.columns.length + 3 }, `Cancelled: not on the latest run sheet (${cancelled.length}). Copied below the list, under a CANCELLED row. Untick any you don’t want to keep.`),
        )
      : null,
    cancelled.map((r) => renderRow(r, -1)),
  );

  const count = active.filter((r) => r.included).length;
  const cancelledCount = cancelled.filter((r) => r.included).length;
  const emailPatients = emailCount();
  const hasRecipients = medicalRecordsEmail().to.length > 0;
  const where = layout.pasteStart || 'the first cell';
  return step(
    3,
    'Check, edit, copy',
    facts,
    renderUpdate(grid),
    h('div', { class: 'table-wrap' }, h('table', { class: 'grid' }, thead, tbody)),
    h(
      'div',
      { class: 'copy-bar' },
      h(
        'button',
        { class: 'primary', disabled: count + cancelledCount === 0, onclick: () => copyRows(grid) },
        `Copy ${count} patient${count === 1 ? '' : 's'}${cancelledCount ? ` + ${cancelledCount} cancelled` : ''} for Excel`,
      ),
      h(
        'span',
        { class: 'hint' },
        grid.padTo
          ? `Then click ${where} (where you copied your rows from) and press Ctrl+V. It replaces the old rows, blanking any left over.`
          : `Then click ${where} in your sheet and press Ctrl+V.`,
      ),
    ),
    renderEmailLists(),
    h(
      'div',
      { class: 'copy-bar' },
      h(
        'button',
        {
          disabled: emailPatients === 0 || !hasRecipients,
          title: hasRecipients ? '' : 'Add recipients in Settings first',
          onclick: () => emailMedicalRecords(),
        },
        'Email Medical Records (Outlook on the web)',
      ),
      h(
        'button',
        {
          disabled: emailPatients === 0 || !hasRecipients,
          title: 'Opens a draft in the email app set up on this computer',
          onclick: () => emailInApp(),
        },
        'Open in email app',
      ),
      canShare
        ? h('button', { disabled: emailPatients === 0, title: 'Share to Outlook or another app on your phone', onclick: () => shareEmail() }, 'Share…')
        : null,
      h('button', { class: 'ghost', disabled: emailPatients === 0, onclick: () => copyEmailText() }, 'Copy email text'),
      hasRecipients
        ? h('span', { class: 'hint' }, 'Opens a draft for you to check. Nothing is sent until you press Send. “Share…” sends the text only, without recipients.')
        : h(
            'span',
            { class: 'hint' },
            'Add the Medical Records address in ',
            h('button', { class: 'link', onclick: openSettings }, 'Settings'),
            ' to enable this.',
          ),
    ),
    state.status ? h('p', { class: 'status', role: 'status' }, state.status) : null,
    h('p', { class: 'hint small' }, 'Time and consultant are shown for reference and aren’t copied unless you add them as columns in Settings.'),
  );
}

function openUpdate(): void {
  state.updateOpen = true;
  state.updateError = '';
  render();
  app.querySelector<HTMLTextAreaElement>('.update-open textarea')?.focus();
}

/** "List changed?": paste your rows back from the cheat sheet to merge them with the latest run sheet. */
function renderUpdate(grid: Grid): HTMLElement {
  const where = state.settings.layout.pasteStart || 'the first patient row';
  const summary = grid.summary
    ? h(
        'p',
        { class: 'update-summary', role: 'status' },
        `Updated from your sheet: ${summaryText(grid.summary)}. `,
        grid.undo ? h('button', { class: 'link', onclick: () => undoUpdate(grid) }, 'Undo') : null,
      )
    : null;
  if (!state.updateOpen) {
    return h(
      'div',
      { class: 'update' },
      summary,
      h(
        'div',
        { class: 'copy-bar' },
        h('button', { onclick: openUpdate }, grid.summary ? 'Update again from my sheet…' : 'Update from my sheet…'),
        h('span', { class: 'hint' }, 'List changed? With the latest run sheet loaded above, paste your prepped rows back in to keep your notes.'),
      ),
    );
  }
  const box = h('textarea', {
    class: 'paste',
    rows: 4,
    placeholder: 'Paste your rows here with Ctrl+V…',
    'aria-label': 'Rows copied from your cheat sheet',
    onpaste: (e: Event) => {
      e.preventDefault();
      const data = (e as ClipboardEvent).clipboardData;
      updateFromSheet(grid, data?.getData('text/plain') ?? '', data?.getData('text/html') ?? '');
    },
  });
  return h(
    'div',
    { class: 'update update-open' },
    summary,
    h('p', {}, h('strong', {}, 'Update from my sheet. '), `In your cheat sheet, select all your patient rows, from ${where} down to the last one (include any cancelled rows), and copy them with Ctrl+C. Then click the box below and press Ctrl+V.`),
    h('p', { class: 'hint' }, 'Patients are matched by MRN. Your notes stay with each patient in the new order; new patients are added and marked New; patients no longer on this list are kept at the bottom as cancelled.'),
    box,
    state.updateError ? h('p', { class: 'error', role: 'alert' }, state.updateError) : null,
    h('button', { class: 'ghost', onclick: () => ((state.updateOpen = false), (state.updateError = ''), render()) }, 'Cancel'),
  );
}

const COLUMN_WIDTHS: Record<string, number> = {
  'field:eye': 72,
  'field:patient': 210,
  'field:procedureRaw': 200,
  'field:comment': 200,
  'field:history': 160,
  'field:planNotes': 140,
  'field:alerts': 140,
  blank: 90,
};

function columnWidth(source: string): number {
  return COLUMN_WIDTHS[source] ?? 110;
}

function cellClass(source: string, value: string): string {
  if (source === 'field:eye') {
    const eye = value.trim().toUpperCase();
    if (eye === 'RIGHT') return 'eye-right';
    if (eye === 'LEFT') return 'eye-left';
    if (eye === 'BOTH') return 'eye-both';
    return 'eye-missing';
  }
  if (source === 'field:alerts' && value.trim()) return 'alert';
  return '';
}

// Drop a run sheet or settings file anywhere on the page.
const overlay = h('div', { class: 'drop-overlay', hidden: true }, 'Drop a run sheet (.csv) or a settings file (.json)');
document.body.append(overlay);
let dragDepth = 0;
const hasFiles = (e: DragEvent) => Boolean(e.dataTransfer?.types.includes('Files'));
document.addEventListener('dragenter', (e) => {
  if (!hasFiles(e)) return;
  dragDepth += 1;
  overlay.hidden = false;
});
document.addEventListener('dragleave', (e) => {
  if (!hasFiles(e)) return;
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) overlay.hidden = true;
});
document.addEventListener('dragover', (e) => {
  if (hasFiles(e)) e.preventDefault();
});
document.addEventListener('drop', (e) => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  overlay.hidden = true;
  const file = e.dataTransfer?.files[0];
  if (file) void handleDroppedFile(file);
});

render();

// Newer settings from another device (if signed in to sync) replace these once they arrive.
pullBeforeStart(SYNC_APP, STORAGE_KEY).then((result) => {
  if (result === 'pulled') applySyncedSettings();
});

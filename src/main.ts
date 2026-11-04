import './style.css';
import { copyHtml, copyText, htmlTableRows, toTsv } from './clipboard';
import { autoGrow, h } from './dom';
import { startDrag } from './drag';
import { buildEmail, emailBodyHtml, emailBodyText, mailtoUrl, MAX_URL_LENGTH, outlookComposeUrl, weekdayLabel } from './email';
import { extractSessions, type Patient, type Session } from './extract';
import { combineSessions, listOfPatient, orderedSelection } from './lists';
import { copyBlock, findInSheet, findPreppedRows, freshRows, isRectangular, mergePrepped, MergeError, parsePastedRows, withoutTrailingBlankRows, type GridRow, type MergeSummary } from './merge';
import { ParseError, parseTable, type Table } from './parse';
import { findingGroups, hasToken, isLensSurgery, promptColumns, readLens, readTarget, toggleFinding, writeLens, writeTarget } from './prompts';
import Papa from 'papaparse';
import { classifyFile } from './files';
import { readWorkbook, WorkbookError, type SheetData } from './xlsx';
import { defaultSettings, exportSettings, importSettings, loadSettings, saveSettings, STORAGE_KEY, SYNC_APP, type KeyValueStore } from './settings';
import { mountSyncControl, pullBeforeStart, pushSoon } from './sync.js';
import { renderSettings } from './settings-view';

/** The list being prepped: one list, or several ticked together that run on as one. */
interface Grid {
  sessions: Session[];
  /** The ticked lists as one, e.g. "OR 4 · AM + PM". */
  combined: Session;
  /** Which list each patient is booked on. */
  listOf: Map<Patient, Session>;
  rows: GridRow[];
  /** After an update from your sheet: how many rows the copy must cover to replace the old ones. */
  padTo: number;
  /** Where the rows were found in your Excel file (e.g. "B7"), so the copy goes back there. */
  pasteAt: string | null;
  summary: MergeSummary | null;
  /** The list as it was before the last update from your sheet. */
  undo: { rows: GridRow[]; padTo: number; pasteAt: string | null; summary: MergeSummary | null } | null;
  /** The prepped Excel file this list was last updated from, so it's applied once. */
  fromWorkbook?: SheetFile;
  /** The sheet in that file it came from; null when you chose not to use the file. */
  fromSheet?: string | null;
  /** The list as it was before the file, so choosing another sheet starts afresh. */
  beforeWorkbook?: { rows: GridRow[]; padTo: number; pasteAt: string | null; summary: MergeSummary | null; undo: Grid['undo'] };
}

/** Your previously prepped sheet, chosen in step 1 or the update panel. Read on this computer only. */
interface SheetFile {
  name: string;
  sheets: SheetData[];
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
  /** Ticked lists, in run-sheet order. */
  selected: [] as string[],
  grid: null as Grid | null,
  /** Grids by ticked lists, most recently used last, so edits survive switching lists. */
  grids: new Map<string, Grid>(),
  /** A short message that pops up at the bottom and fades. */
  status: '',
  /** The "Update from my sheet" paste box is open. */
  updateOpen: false,
  updateError: '',
  emailMenuOpen: false,
  /** Your previously prepped sheet: each list you open is updated from it. */
  workbook: null as SheetFile | null,
  /** Something went wrong reading a chosen file (shown in step 1). */
  fileError: '',
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
  if (state.view === 'settings') state.settingsSnapshot = listSettings();
  else if (state.table) refreshSessions();
  render();
}

let unmountSync: (() => void) | null = null;
let statusTimer = 0;

/** Shows a message at the bottom of the screen for a few seconds. */
function say(text: string, redraw = true): void {
  state.status = text;
  clearTimeout(statusTimer);
  statusTimer = window.setTimeout(() => {
    state.status = '';
    render();
  }, 9000);
  if (redraw) render();
}

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
  state.selected = [];
  refreshSessions();
  render();
}

/** Re-extracts after the run sheet or settings change, opening the list straight away if there is only one. */
function refreshSessions(): void {
  const { flagHospitals, hospitalCodes } = state.settings;
  state.grids.clear();
  state.sessions = state.table ? extractSessions(state.table, { flagHospitals, hospitalCodes }) : [];
  const lists = listedSessions();
  state.selected = orderedSelection(lists, state.selected);
  if (state.selected.length === 0 && lists.length === 1) state.selected = [lists[0].id];
  openSelection();
}

/**
 * The grid for the ticked lists. A new combination starts from the rows you've already edited on
 * each list (on its own or in another combination), so nothing is lost by ticking another list.
 */
function gridFor(ids: string[]): Grid {
  const key = ids.join('+');
  const cached = state.grids.get(key);
  if (cached) {
    state.grids.delete(key);
    state.grids.set(key, cached);
    return cached;
  }
  const sessions = ids.map((id) => state.sessions.find((s) => s.id === id)).filter((s): s is Session => Boolean(s));
  const recent = [...state.grids.values()].reverse();
  const rows = sessions.flatMap((s) => {
    const source = recent.find((g) => g.sessions.includes(s));
    return source
      ? source.rows.filter((r) => r.patient && r.status !== 'cancelled' && source.listOf.get(r.patient) === s)
      : freshRows(s.patients, state.settings.layout);
  });
  const grid: Grid = { sessions, combined: combineSessions(sessions), listOf: listOfPatient(sessions), rows, padTo: 0, pasteAt: null, summary: null, undo: null };
  state.grids.set(key, grid);
  return grid;
}

function openSelection(): void {
  state.grid = state.selected.length ? gridFor(state.selected) : null;
  state.updateOpen = false;
  state.updateError = '';
  state.emailMenuOpen = false;
  // With your prepped sheet chosen, each list is updated from it the first time it's opened.
  if (state.grid && state.workbook && state.grid.fromWorkbook !== state.workbook) applyWorkbook(state.grid, state.workbook);
}

/** Opens one list on its own. */
function chooseList(id: string): void {
  state.selected = [id];
  state.status = '';
  openSelection();
  render();
}

/** Ticks or unticks a list, to prep lists that run on together as one. */
function tickList(id: string, on: boolean): void {
  const ids = on ? [...state.selected, id] : state.selected.filter((x) => x !== id);
  state.selected = orderedSelection(listedSessions(), ids);
  state.status = '';
  openSelection();
  render();
}

/** Ticked patients still on the run sheet (not cancelled). */
function tickedPatients(grid: Grid): Patient[] {
  return grid.rows.filter((r) => r.included && r.status !== 'cancelled' && r.patient).map((r) => r.patient!);
}

/** One group per ticked list, so the email names each list when there are several. */
function medicalRecordsEmail(grid: Grid) {
  const ticked = tickedPatients(grid);
  const groups = grid.sessions.map((session) => ({ session, patients: ticked.filter((p) => grid.listOf.get(p) === session) }));
  return buildEmail(state.settings.email, groups);
}

function emailSummary(grid: Grid): string {
  const ticked = tickedPatients(grid);
  const lists = new Set(ticked.map((p) => grid.listOf.get(p))).size;
  return `${ticked.length} patient${ticked.length === 1 ? '' : 's'}${lists > 1 ? ` from ${lists} lists` : ''}`;
}

/** Opens a draft in Outlook on the web; copies the text instead if the link is too long or blocked. */
async function emailMedicalRecords(grid: Grid): Promise<void> {
  const email = medicalRecordsEmail(grid);
  const url = outlookComposeUrl(state.settings.email.outlookUrl, email);
  const tooLong = url.length > MAX_URL_LENGTH;
  // 'noopener' would make window.open return null even on success, so detach the new tab by hand.
  const tab = tooLong ? null : window.open(url, '_blank');
  if (!tab) {
    await copyEmailText(grid);
    return;
  }
  tab.opener = null;
  say(`Opened a draft email with ${emailSummary(grid)}. Check it and press Send in Outlook.`);
}

/** Hands the draft to the computer's default email app via a mailto link, without leaving this page. */
async function emailInApp(grid: Grid): Promise<void> {
  const url = mailtoUrl(medicalRecordsEmail(grid));
  if (url.length > MAX_URL_LENGTH) {
    await copyEmailText(grid);
    return;
  }
  const summary = emailSummary(grid);
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
  say(`Asked your email app to open a draft with ${summary}…`);
  setTimeout(() => {
    window.removeEventListener('blur', noteHandover);
    say(
      handedOver
        ? `Opened a draft with ${summary}. Check it and press Send.`
        : 'No email app seems to have opened. This computer may not have a default email app set up: use Outlook on the web or Copy email text instead.',
    );
  }, 1500);
}

const canShare = typeof navigator !== 'undefined' && typeof navigator.share === 'function';

/** Phone/tablet share sheet. Recipients can't be passed this way, so they're added in the email app. */
async function shareEmail(grid: Grid): Promise<void> {
  const email = medicalRecordsEmail(grid);
  try {
    await navigator.share({ title: email.subject, text: email.body });
    say('Shared. Add the Medical Records address in your email app before sending.');
  } catch (e) {
    if (e instanceof DOMException && e.name === 'AbortError') return;
    await copyEmailText(grid);
  }
}

async function copyEmailText(grid: Grid): Promise<void> {
  try {
    const email = medicalRecordsEmail(grid);
    await copyHtml(emailBodyHtml(email), emailBodyText(email));
    say(`Copied the message for ${emailSummary(grid)} (without the recipients or subject). Paste it into the body of a new email.`);
  } catch {
    say('Your browser blocked copying the email text.');
  }
}

/** Reads a chosen or dropped file as sheets of cells: an Excel workbook, or a CSV as one sheet. */
async function readSheets(file: File): Promise<SheetData[]> {
  if (classifyFile(file.name, file.type) === 'workbook') return readWorkbook(await file.arrayBuffer());
  const rows = Papa.parse<string[]>((await file.text()).replace(/^\uFEFF/, ''), { skipEmptyLines: false }).data;
  return [{ name: file.name, rows }];
}

/** The run sheet in a workbook, as text for the parser, if one of its sheets has the run-sheet headings. */
function runSheetIn(sheets: SheetData[]): string | null {
  for (const sheet of sheets) {
    const text = toTsv(sheet.rows);
    try {
      parseTable(text);
      return text;
    } catch {
      // Not this sheet.
    }
  }
  return null;
}

/**
 * Opens a file chosen or dropped: the run sheet (CSV or Excel), your prepped Excel sheet, or a
 * settings file. A workbook with run-sheet headings is the run sheet; any other is your prepped sheet.
 */
async function openFile(file: File, as: 'any' | 'runsheet' | 'prepped' = 'any'): Promise<void> {
  const kind = classifyFile(file.name, file.type);
  if (kind === 'settings') return importSettingsFile(file);
  state.notice = null;
  state.fileError = '';
  if (state.view === 'settings') closeSettings();
  if (kind === 'runsheet' && as !== 'prepped') {
    loadRunSheet(await file.text());
    return;
  }
  let sheets: SheetData[];
  try {
    sheets = await readSheets(file);
  } catch (e) {
    state.fileError = e instanceof WorkbookError ? `${file.name}: ${e.message}` : `Couldn’t read ${file.name}.`;
    render();
    return;
  }
  const runSheet = as === 'prepped' ? null : runSheetIn(sheets);
  if (runSheet) {
    loadRunSheet(runSheet);
    say(`Run sheet loaded from ${file.name}.`);
    return;
  }
  if (as === 'runsheet') {
    state.fileError = `${file.name} doesn’t have a run sheet in it (no THEATRE and PT_NAME headings). If it’s your prepped sheet, choose it under “Previously prepped”.`;
    render();
    return;
  }
  state.workbook = { name: file.name, sheets };
  if (state.grid) applyWorkbook(state.grid, state.workbook);
  else say(state.table ? `Prepped sheet ${file.name} ready: choose a list to update it.` : `Prepped sheet ${file.name} ready. Now load the latest run sheet.`, false);
  render();
}

/**
 * Updates a list from your prepped sheet. Without a sheet named, the workbook's first sheet is used
 * (choose another above the list); null means don't use the file.
 */
function applyWorkbook(grid: Grid, workbook: SheetFile, sheetName?: string | null): void {
  // Start from the list as it was before any prepped file, so changing sheet doesn't stack.
  if (grid.beforeWorkbook) Object.assign(grid, grid.beforeWorkbook);
  else grid.beforeWorkbook = { rows: grid.rows, padTo: grid.padTo, pasteAt: grid.pasteAt, summary: grid.summary, undo: grid.undo };
  grid.fromWorkbook = workbook;
  grid.fromSheet = null;
  if (sheetName === null) {
    say(`Not using ${workbook.name} for this list: it’s as on the run sheet.`, false);
    return;
  }
  const chosen = sheetName ?? workbook.sheets[0]?.name;
  const isDefault = sheetName === undefined && workbook.sheets.length > 1;
  let found;
  try {
    found = findPreppedRows(workbook.sheets.filter((s) => s.name === chosen), grid.combined, state.settings.layout);
  } catch (e) {
    say(e instanceof MergeError ? e.message : `Couldn’t read ${workbook.name}.`, false);
    return;
  }
  if (!found) {
    say(
      workbook.sheets.length > 1
        ? `None of the patients on ${listName(grid.combined)} are on ${isDefault ? 'the first sheet, ' : 'sheet '}“${chosen}”, so the list is as on the run sheet. Choose its sheet above the list.`
        : `None of the patients on ${listName(grid.combined)} are in ${workbook.name}, so the list is as on the run sheet.`,
      false,
    );
    return;
  }
  const error = mergeInto(grid, found.rows, found.cell);
  if (error) {
    say(error, false);
    return;
  }
  grid.fromSheet = found.sheet;
  const where = workbook.sheets.length > 1 ? `${isDefault ? 'first sheet' : 'sheet'} “${found.sheet}”, from ${found.cell}` : `from ${found.cell}`;
  say(`Updated from ${workbook.name} (${where}).${isDefault ? ' Wrong sheet? Choose another above the list.' : ''}`, false);
}

/** Above the list when a prepped file is chosen: which sheet this list came from, to change it. */
function renderSource(grid: Grid): HTMLElement | null {
  const wb = state.workbook;
  if (!wb || grid.fromWorkbook !== wb) return null;
  const options = wb.sheets.map((sheet) => {
    const found = (() => {
      try {
        return findInSheet(sheet, grid.combined, state.settings.layout);
      } catch {
        return null;
      }
    })();
    const count = found?.matches ?? 0;
    return h(
      'option',
      { value: sheet.name },
      `${sheet.name} (${count ? `${count} of ${grid.combined.patients.length} patients` : 'none of this list'})`,
    );
  });
  const select = h(
    'select',
    {
      'aria-label': 'Sheet to update this list from',
      onchange: (e: Event) => {
        const value = (e.target as HTMLSelectElement).value;
        applyWorkbook(grid, wb, value === '' ? null : value);
        render();
      },
    },
    h('option', { value: '' }, 'Don’t use (as on the run sheet)'),
    options,
  );
  select.value = grid.fromSheet ?? '';
  return h('div', { class: 'source-bar' }, h('span', { class: 'muted' }, 'From your prepped sheet '), h('strong', {}, wb.name), wb.sheets.length > 1 || !grid.fromSheet ? [h('span', { class: 'muted' }, ' · Sheet '), select] : null);
}

function forgetWorkbook(): void {
  state.workbook = null;
  render();
}

async function importSettingsFile(file: File): Promise<void> {
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
  state.settingsSnapshot = listSettings();
  state.view = 'settings';
  render();
  window.scrollTo(0, 0);
}

/** The settings that shape the list's rows: changing these rebuilds it. Email and prompts don't. */
function listSettings(): string {
  const { layout, flagHospitals, hospitalCodes } = state.settings;
  return JSON.stringify({ layout, flagHospitals, hospitalCodes });
}

/** Back to the list; rebuilds it (losing edits) only if the settings that shape it changed. */
function closeSettings(): void {
  state.view = 'main';
  if (listSettings() !== state.settingsSnapshot) {
    state.status = '';
    refreshSessions();
  }
  render();
}


/** Moves a row within the list; cancelled rows stay together at the bottom. */
function moveRow(grid: Grid, from: number, to: number): void {
  if (to < 0 || to >= activeRows(grid).length || from === to) return;
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
  const error = mergeInto(grid, pasted, null);
  state.updateError = error ?? '';
  if (!error) {
    state.updateOpen = false;
    state.status = '';
  }
  render();
}

/** Merges your rows into the list (Undo keeps what was there). Returns an error message, if any. */
function mergeInto(grid: Grid, rows: string[][], pasteAt: string | null): string | null {
  try {
    const others = listedSessions().filter((s) => !grid.sessions.includes(s));
    const result = mergePrepped(rows, grid.combined, others, state.settings.layout, listName);
    grid.undo = { rows: grid.rows, padTo: grid.padTo, pasteAt: grid.pasteAt, summary: grid.summary };
    grid.rows = result.rows;
    grid.padTo = result.padTo;
    grid.pasteAt = pasteAt;
    grid.summary = result.summary;
    return null;
  } catch (e) {
    return e instanceof MergeError ? e.message : 'Couldn’t read those rows. Copy them from your sheet and paste again.';
  }
}

function undoUpdate(grid: Grid): void {
  if (!grid.undo) return;
  // Undoing the update from your prepped file leaves the list as on the run sheet.
  if (grid.beforeWorkbook && grid.undo.rows === grid.beforeWorkbook.rows) grid.fromSheet = null;
  Object.assign(grid, grid.undo, { undo: null });
  say('Update undone.');
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
    const where = grid.pasteAt || layout.pasteStart || 'the first cell';
    say(
      grid.padTo
        ? `Copied ${count} patient${count === 1 ? '' : 's'}${cancelled ? ` and ${cancelled} cancelled` : ''}. Click ${where} (where you copied your rows from) and press Ctrl+V to replace the old rows.`
        : `Copied ${count} patient${count === 1 ? '' : 's'}. Click ${where} in your sheet and press Ctrl+V.`,
    );
  } catch {
    say('Your browser blocked copying. Try again, or use Ctrl+C after selecting the table.');
  }
}

// ---- Rendering ----------------------------------------------------------------

function render(): void {
  const scroll = window.scrollY;
  const focusKey = (document.activeElement as HTMLElement | null)?.dataset?.key;
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
      : h('main', {}, renderSetup(), renderPrep()),
    h(
      'div',
      { class: 'toast-area', role: 'status', 'aria-live': 'polite' },
      state.status
        ? h('div', { class: 'toast' }, h('span', {}, state.status), h('button', { class: 'toast-close', title: 'Dismiss', onclick: () => ((state.status = ''), render()) }, '×'))
        : null,
    ),
  );
  for (const area of app.querySelectorAll<HTMLTextAreaElement>('textarea.cell')) autoGrow(area);
  if (focusKey) app.querySelector<HTMLElement>(`[data-key="${CSS.escape(focusKey)}"]`)?.focus();
  window.scrollTo(0, scroll);
}

function withSyncPanel(view: HTMLElement): HTMLElement {
  unmountSync?.();
  const slot = h('div', { class: 'sync-slot' });
  view.prepend(slot);
  unmountSync = mountSyncControl(slot, SYNC_APP, STORAGE_KEY, applySyncedSettings);
  return view;
}


function step(n: number, title: string, ...body: (Node | string | null | false | (Node | string | null)[])[]): HTMLElement {
  return h('section', { class: 'step' }, h('h2', {}, h('span', { class: 'num' }, String(n)), title), ...body);
}

/** Steps 1 and 2: paste the run sheet, then pick the list (or lists). Compact once loaded. */
const RUN_SHEET_FILES = '.csv,.tsv,.txt,.xlsx,.xlsm,text/csv';
const PREPPED_FILES = '.xlsx,.xlsm,.csv';

/** Opens the file chooser; works without putting an input on the page. */
function pickFile(accept: string, onFile: (file: File) => void): void {
  const input = h('input', { type: 'file', accept });
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (file) onFile(file);
  });
  input.click();
}

/** "Previously prepped": your Excel sheet, so each list keeps your notes when the run sheet has changed. */
function renderPrepped(compact: boolean): HTMLElement {
  const choose = () => pickFile(PREPPED_FILES, (file) => void openFile(file, 'prepped'));
  const wb = state.workbook;
  if (compact) {
    return h(
      'span',
      { class: 'prepped-line' },
      wb
        ? [h('span', { class: 'muted' }, 'Prepped sheet: '), h('strong', {}, wb.name), ' ', h('button', { class: 'link', onclick: choose }, 'Change'), ' · ', h('button', { class: 'link', onclick: forgetWorkbook }, 'Remove')]
        : h('button', { class: 'link', title: 'Already prepped these lists? Your notes are kept when the list has changed.', onclick: choose }, 'Add your prepped Excel sheet…'),
    );
  }
  return h(
    'div',
    { class: 'prepped-box' },
    h('h3', {}, 'Previously prepped? ', h('span', { class: 'muted' }, '(optional)')),
    h('p', { class: 'hint small' }, 'Choose your cheat sheet from last time (Excel). Each list you open is updated from it: your notes stay with each patient, new patients are marked New and cancellations go to the bottom. It’s read on this computer only.'),
    wb
      ? h('p', {}, '✓ ', h('strong', {}, wb.name), ' ', h('button', { class: 'link', onclick: choose }, 'Change'), ' · ', h('button', { class: 'link', onclick: forgetWorkbook }, 'Remove'))
      : h('button', { onclick: choose }, 'Choose prepped Excel sheet…'),
  );
}

/** Steps 1 and 2: the run sheet (and your prepped sheet), then the list or lists. Compact once loaded. */
function renderSetup(): HTMLElement {
  const fileError = state.fileError ? h('p', { class: 'error', role: 'alert' }, state.fileError) : null;
  if (!state.table) {
    return step(
      1,
      'Run sheet',
      h(
        'div',
        { class: 'setup-inputs' },
        h(
          'div',
          {},
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
            h('button', { class: 'link', onclick: () => pickFile(RUN_SHEET_FILES, (file) => void openFile(file, 'runsheet')) }, 'choose the run sheet file'),
            ' (CSV or Excel), or drop it anywhere on this page.',
          ),
          state.error ? h('p', { class: 'error', role: 'alert' }, state.error) : null,
          state.error
            ? h('p', { class: 'hint' }, 'Got rows from your cheat sheet? Load the latest run sheet here first, then choose your prepped sheet on the right or paste the rows under “Update from my sheet”.')
            : null,
        ),
        renderPrepped(false),
      ),
      fileError,
    );
  }

  const lists = listedSessions();
  const patients = lists.reduce((n, s) => n + s.patients.length, 0);
  const dates = sheetDates();
  const when = dates.length === 1 ? weekdayLabel(dates[0]) : dates.length > 1 ? `${shortDay(dates[0])} to ${shortDay(dates[dates.length - 1])}` : '';
  const card = (s: Session) => {
    const on = state.selected.includes(s.id);
    const only = on && state.selected.length === 1;
    return h(
      'div',
      { class: on ? 'list-card on' : 'list-card' },
      h('input', {
        type: 'checkbox',
        checked: on,
        title: on ? 'Untick to leave this list out' : 'Tick to add this list to the one you’re prepping',
        'aria-label': `Include ${listName(s)}`,
        onchange: (e: Event) => tickList(s.id, (e.target as HTMLInputElement).checked),
      }),
      h(
        'button',
        { class: 'list-open', 'aria-pressed': only ? 'true' : 'false', title: 'Open this list on its own', onclick: () => chooseList(s.id) },
        h('span', { class: 'list-title' }, `${s.or} · ${s.ampm}`),
        h('span', { class: 'list-sub' }, `${s.consultant || s.slotName} · ${s.patients.length} patient${s.patients.length === 1 ? '' : 's'}`),
      ),
    );
  };
  const groups =
    dates.length > 1
      ? dates.map((date) => h('div', { class: 'day-group' }, h('div', { class: 'day-heading' }, weekdayLabel(date)), h('div', { class: 'list-cards' }, lists.filter((s) => s.date === date).map(card))))
      : [h('div', { class: 'list-cards' }, lists.map(card))];

  return h(
    'section',
    { class: 'step setup' },
    h(
      'div',
      { class: 'setup-head' },
      h('h2', {}, h('span', { class: 'num' }, '1'), 'Run sheet'),
      h('span', { class: 'muted' }, [when, `${patients} patient${patients === 1 ? '' : 's'} on ${lists.length} list${lists.length === 1 ? '' : 's'}`].filter(Boolean).join(' · ')),
      h('button', { class: 'link', onclick: () => loadRunSheet('') }, 'Load another'),
      renderPrepped(true),
    ),
    fileError,
    lists.length === 0
      ? h('p', { class: 'hint' }, 'No lists with patients were found on this run sheet.')
      : h(
          'div',
          { class: 'setup-lists' },
          h('h2', {}, h('span', { class: 'num' }, '2'), lists.length === 1 ? 'List' : 'Lists'),
          h('div', { class: 'list-groups' }, groups),
          lists.length > 1 ? h('p', { class: 'hint small' }, 'Click a list to open it. Tick more than one to prep lists that run on without a break as one list.') : null,
        ),
  );
}

/** Distinct dates on the run sheet, in order. */
function sheetDates(): string[] {
  return [...new Set(listedSessions().map((s) => s.date).filter(Boolean))];
}

/** "Tue 13/10" */
function shortDay(date: string): string {
  return weekdayLabel(date).replace(/^(\w{3})\w*/, '$1');
}


/** Step 3: the list itself, with one toolbar for copying, updating and emailing. */
function renderPrep(): HTMLElement | null {
  const grid = state.grid;
  if (!grid) return null;
  const { combined } = grid;
  const { layout } = state.settings;

  const facts: [string, string][] = [
    ['Consultant', combined.consultant],
    ['Date', combined.date],
    ['OR', combined.or],
    ['Session', combined.ampm],
  ];
  const factChips = h(
    'div',
    { class: 'facts' },
    facts.map(([label, value]) =>
      h(
        'button',
        {
          class: 'fact',
          title: `Copy ${label.toLowerCase()}`,
          disabled: !value,
          onclick: async () => {
            await copyText(value).catch(() => undefined);
            say(`Copied ${label.toLowerCase()} “${value}”.`);
          },
        },
        h('span', { class: 'fact-label' }, label),
        h('span', { class: 'fact-value' }, value || '—'),
        h('span', { class: 'copy-icon', 'aria-hidden': 'true' }, '⧉'),
      ),
    ),
  );

  const active = activeRows(grid);
  const cancelled = cancelledRows(grid);
  const count = active.filter((r) => r.included).length;
  const cancelledCount = cancelled.filter((r) => r.included).length;
  const where = grid.pasteAt || layout.pasteStart || 'the first cell';

  const toolbar = h(
    'div',
    { class: 'toolbar' },
    h(
      'button',
      {
        class: 'primary',
        disabled: count + cancelledCount === 0,
        title: grid.padTo ? `Paste at ${where}, where you copied your rows from. It replaces the old rows.` : `Then click ${where} in your sheet and press Ctrl+V.`,
        onclick: () => copyRows(grid),
      },
      `Copy ${count} patient${count === 1 ? '' : 's'}${cancelledCount ? ` + ${cancelledCount} cancelled` : ''} for Excel`,
    ),
    h(
      'button',
      {
        class: state.updateOpen ? 'on' : '',
        title: 'List changed? With the latest run sheet loaded, paste your prepped rows back in to keep your notes.',
        onclick: () => (state.updateOpen ? closeUpdate() : openUpdate()),
      },
      grid.summary ? 'Update again from my sheet…' : 'Update from my sheet…',
    ),
    renderEmailMenu(grid),
    h('span', { class: 'toolbar-count muted' }, `${count} of ${active.length} ticked`),
  );

  const columnCount = layout.columns.length + 2;
  const allTicked = active.length > 0 && active.every((r) => r.included);
  const thead = h(
    'thead',
    {},
    h(
      'tr',
      {},
      h(
        'th',
        { class: 'handle' },
        h('input', {
          type: 'checkbox',
          checked: allTicked,
          title: allTicked ? 'Untick all' : 'Tick all',
          'aria-label': 'Tick all patients',
          onchange: (e: Event) => {
            const on = (e.target as HTMLInputElement).checked;
            for (const r of active) r.included = on;
            render();
          },
        }),
      ),
      h('th', { class: 'time', title: 'Shown for reference. Add Time as a column in Settings to copy it.' }, 'Time'),
      layout.columns.map((c) => h('th', {}, c.heading)),
    ),
  );

  const keyIndex = layout.columns.findIndex((c) => c.source === 'field:patient' || c.source === 'field:name' || c.source === 'field:mrn');
  const rowName = (r: GridRow) => r.patient?.name || (keyIndex >= 0 ? r.cells[keyIndex] : '') || 'this patient';
  const combinedLists = grid.sessions.length > 1;

  const renderRow = (r: GridRow, index: number) => {
    const p = r.patient;
    const isCancelled = r.status === 'cancelled';
    const list = p ? grid.listOf.get(p) : undefined;
    const who = [
      combinedLists && list ? `${list.or} ${list.ampm}` : '',
      p && list && p.consultantSurname && p.consultantSurname !== list.consultant ? p.consultantSurname : '',
    ].filter(Boolean).join(' · ');
    const badges = [
      r.status === 'new' ? h('span', { class: 'badge badge-new', title: 'Added to the list since your sheet: needs prep' }, 'New') : null,
      r.reinstated ? h('span', { class: 'badge badge-new', title: 'Was cancelled on your sheet and is back on the run sheet' }, 'Back on list') : null,
      r.movedFrom !== undefined ? h('span', { class: 'badge badge-moved', title: 'The order changed: this patient has moved' }, `Moved, was #${r.movedFrom}`) : null,
      isCancelled ? h('span', { class: 'badge badge-cancelled' }, 'Cancelled') : null,
      r.note ? h('div', { class: 'who' }, r.note) : null,
    ];
    const prompts = isCancelled ? null : renderPrompts(r);
    const gripKey = p ? `${p.id}:grip` : undefined;
    const grip = isCancelled
      ? null
      : h(
          'button',
          {
            class: 'grip',
            title: 'Drag to move (or use the ↑ ↓ keys)',
            'aria-label': `Move ${rowName(r)}: use the up and down arrow keys`,
            'data-key': gripKey,
            onpointerdown: (e: Event) =>
              startDrag(e as PointerEvent, {
                items: [...app.querySelectorAll<HTMLElement>('tbody.patient.active')],
                from: index,
                label: rowName(r),
                onDrop: (from, to) => moveRow(grid, from, to),
              }),
            onkeydown: (e: Event) => {
              const key = (e as KeyboardEvent).key;
              if (key !== 'ArrowUp' && key !== 'ArrowDown') return;
              e.preventDefault();
              moveRow(grid, index, key === 'ArrowUp' ? index - 1 : index + 1);
            },
          },
          '⠿',
        );
    const classes = [
      'patient',
      isCancelled ? 'row-cancelled' : 'active',
      r.included ? '' : 'excluded',
      r.status === 'new' || r.reinstated ? 'row-new' : '',
    ];
    const main = h(
      'tr',
      {},
      h(
        'td',
        { class: 'handle' },
        grip,
        h('input', {
          type: 'checkbox',
          checked: r.included,
          title: isCancelled ? 'Untick to leave this cancelled patient off the copy' : 'Untick to leave this patient off the copy and the email',
          'aria-label': `Include ${rowName(r)}`,
          onchange: (e: Event) => {
            r.included = (e.target as HTMLInputElement).checked;
            render();
          },
        }),
      ),
      h('td', { class: 'time', title: p?.consultant ?? '' }, h('div', { class: 'time-value' }, p?.time ?? ''), who ? h('div', { class: 'who' }, who) : null, ...badges),
      layout.columns.map((c, col) => {
        const area = h('textarea', {
          class: 'cell',
          rows: 1,
          value: r.cells[col],
          'aria-label': `${c.heading} for ${rowName(r)}`,
          'data-key': p ? `${p.id}:cell${col}` : undefined,
          // Leaving an edited cell redraws the prompts, so typing "PXF" by hand lights up its button too.
          onchange: prompts ? () => setTimeout(render) : undefined,
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
    return h(
      'tbody',
      { class: classes.filter(Boolean).join(' ') },
      main,
      prompts ? h('tr', { class: 'prompt-row' }, h('td', { colspan: 2 }), h('td', { colspan: layout.columns.length }, prompts)) : null,
    );
  };

  const table = h(
    'table',
    { class: 'grid' },
    thead,
    active.map(renderRow),
    cancelled.length
      ? h(
          'tbody',
          { class: 'cancelled-heading' },
          h('tr', {}, h('td', { colspan: columnCount }, `Cancelled: not on the latest run sheet (${cancelled.length}). Copied below the list under a CANCELLED row; untick any you don’t want to keep.`)),
        )
      : null,
    cancelled.map((r) => renderRow(r, -1)),
  );

  return h(
    'section',
    { class: 'step prep' },
    h(
      'div',
      { class: 'prep-head' },
      h('h2', {}, h('span', { class: 'num' }, '3'), listName(combined)),
      factChips,
    ),
    toolbar,
    renderSource(grid),
    renderUpdate(grid),
    h('div', { class: 'table-wrap' }, table),
  );
}

/** Email Medical Records, with the ways to send it in a small menu. */
function renderEmailMenu(grid: Grid): HTMLElement {
  const patients = tickedPatients(grid).length;
  const hasRecipients = medicalRecordsEmail(grid).to.length > 0;
  const pick = (action: (g: Grid) => Promise<void>) => () => {
    state.emailMenuOpen = false;
    render();
    void action(grid);
  };
  const item = (label: string, detail: string, onclick: () => void, disabled = false) =>
    h('button', { class: 'menu-item', role: 'menuitem', disabled, onclick }, h('span', {}, label), h('span', { class: 'menu-detail' }, detail));
  return h(
    'div',
    { class: 'menu-wrap' },
    h(
      'button',
      {
        class: state.emailMenuOpen ? 'on' : '',
        disabled: patients === 0,
        'aria-haspopup': 'menu',
        'aria-expanded': state.emailMenuOpen ? 'true' : 'false',
        onclick: () => ((state.emailMenuOpen = !state.emailMenuOpen), render()),
      },
      'Email Medical Records ▾',
    ),
    state.emailMenuOpen
      ? h(
          'div',
          { class: 'menu', role: 'menu' },
          h('p', { class: 'menu-note' }, `Asks for the paper files of ${emailSummary(grid)}. Opens a draft for you to check: nothing is sent until you press Send.`),
          item('Outlook on the web', 'Opens a draft in a new tab', pick(emailMedicalRecords), !hasRecipients),
          item('Email app', 'The email app set up on this computer', pick(emailInApp), !hasRecipients),
          canShare ? item('Share…', 'Text only: add the address yourself', pick(shareEmail)) : null,
          item('Copy email text', 'To paste into any email', pick(copyEmailText)),
          hasRecipients
            ? null
            : h('p', { class: 'menu-note' }, 'Add the Medical Records address in ', h('button', { class: 'link', onclick: () => ((state.emailMenuOpen = false), openSettings()) }, 'Settings'), ' to open a draft.'),
        )
      : null,
  );
}

function openUpdate(): void {
  state.updateOpen = true;
  state.updateError = '';
  render();
  app.querySelector<HTMLTextAreaElement>('.update-open textarea')?.focus();
}

function closeUpdate(): void {
  state.updateOpen = false;
  state.updateError = '';
  render();
}

/** "List changed?": paste your rows back from the cheat sheet to merge them with the latest run sheet. */
function renderUpdate(grid: Grid): HTMLElement | null {
  const where = grid.pasteAt || state.settings.layout.pasteStart || 'the first patient row';
  const summary = grid.summary
    ? h(
        'p',
        { class: 'update-summary', role: 'status' },
        `Updated from your sheet: ${summaryText(grid.summary)}. `,
        grid.undo ? h('button', { class: 'link', onclick: () => undoUpdate(grid) }, 'Undo') : null,
      )
    : null;
  if (!state.updateOpen) return summary;
  const box = h('textarea', {
    class: 'paste',
    rows: 3,
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
    { class: 'update-open' },
    summary,
    h(
      'p',
      {},
      h('strong', {}, 'Update from my sheet. '),
      `In your cheat sheet, select all your patient rows from ${where} down (include any cancelled rows), copy them with Ctrl+C, then click the box below and press Ctrl+V.`,
    ),
    h('p', { class: 'hint small' }, 'Patients are matched by MRN. Your notes stay with each patient in the new order; new patients are marked New, and patients no longer on the list are kept at the bottom as cancelled.'),
    box,
    h(
      'p',
      { class: 'hint small' },
      'Or ',
      h('button', { class: 'link', onclick: () => pickFile(PREPPED_FILES, (file) => void openFile(file, 'prepped')) }, 'choose your Excel file'),
      ' instead: the rows for this list are found in it for you.',
    ),
    state.updateError ? h('p', { class: 'error', role: 'alert' }, state.updateError) : null,
    h('button', { class: 'ghost', onclick: closeUpdate }, 'Cancel'),
  );
}

/** One-click findings, lens and target under a lens patient. Each click edits the cell itself. */
function renderPrompts(r: GridRow): HTMLElement | null {
  const { prompts, layout } = state.settings;
  if (!prompts.enabled || !r.patient || !isLensSurgery(r.patient)) return null;
  const cols = promptColumns(layout, prompts);
  if (cols.findings === -1 && cols.lens === -1) return null;
  const key = r.patient.id;
  const set = (col: number, value: string) => {
    r.cells[col] = value;
    // After the event, so a Tab to the next box lands before the redraw (which keeps that focus).
    setTimeout(render);
  };
  const chip = (label: string, on: boolean, onclick: () => void, title = '') =>
    h('button', { class: on ? 'chip on' : 'chip', 'aria-pressed': on ? 'true' : 'false', title, 'data-key': `${key}:chip:${label}`, onclick }, label);
  const field = (name: string, value: string, placeholder: string, label: string, onChange: (v: string) => void, size = 6.5) =>
    h('input', {
      type: 'text',
      class: 'prompt-input',
      style: `width:${size}em`,
      value,
      placeholder,
      'aria-label': label,
      'data-key': `${key}:${name}`,
      onchange: (e: Event) => onChange((e.target as HTMLInputElement).value),
    });

  const parts: HTMLElement[] = [];
  if (cols.findings !== -1) {
    const col = cols.findings;
    const cell = r.cells[col];
    parts.push(
      h(
        'span',
        { class: 'prompt-set' },
        findingGroups(prompts.findings).map((g) => {
          const chips = g.options.map((o) =>
            chip(o.short, hasToken(cell, o.token), () => set(col, toggleFinding(cell, g, o.token, prompts.findings)), o.token),
          );
          return g.options.length > 1 ? h('span', { class: 'segmented' }, g.label ? h('span', { class: 'prompt-label' }, g.label) : null, chips) : chips[0];
        }),
      ),
    );
  }
  if (cols.lens !== -1) {
    const col = cols.lens;
    const cell = r.cells[col];
    const pick = readLens(cell, prompts.lenses);
    const target = readTarget(cell);
    const lens = h(
      'span',
      { class: 'prompt-set' },
      h('span', { class: 'prompt-label' }, 'Lens'),
      prompts.lenses.map((l) =>
        chip(
          l.replace('x', '…'),
          pick?.lens === l,
          () => set(col, writeLens(cell, prompts.lenses, pick?.lens === l ? null : { lens: l, extra: '', power: pick?.power ?? '' })),
          l.includes('x') ? `${l}: type the number in the box that appears` : l,
        ),
      ),
      pick?.lens.includes('x')
        ? field('extra', pick.extra, 'no.', `Number for ${pick.lens}`, (v) => set(col, writeLens(r.cells[col], prompts.lenses, { ...pick, extra: v.trim() })), 3.5)
        : null,
      pick ? field('power', pick.power, 'power', 'Lens power', (v) => set(col, writeLens(r.cells[col], prompts.lenses, { ...pick, power: v })), 6) : null,
    );
    const targetSet = h(
      'span',
      { class: 'prompt-set' },
      h('span', { class: 'prompt-label' }, 'Target'),
      chip('Plano', target === 'plano', () => set(col, writeTarget(cell, prompts.lenses, target === 'plano' ? null : 'plano'))),
      field('target', target && target !== 'plano' ? target : '', 'e.g. -0.50', 'Target refraction', (v) =>
        set(col, writeTarget(r.cells[col], prompts.lenses, v.trim() ? v : null)),
      ),
    );
    parts.push(lens, targetSet);
  }
  return h('div', { class: 'prompts' }, parts);
}

const COLUMN_WIDTHS: Record<string, number> = {
  'field:eye': 68,
  'field:patient': 180,
  'field:procedureRaw': 180,
  'field:comment': 180,
  'field:history': 150,
  'field:planNotes': 120,
  'field:alerts': 120,
  blank: 72,
};

function columnWidth(source: string): number {
  return COLUMN_WIDTHS[source] ?? 96;
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
  if (file) void openFile(file);
});


// Close the email menu on a click elsewhere or Escape.
document.addEventListener('pointerdown', (e) => {
  if (state.emailMenuOpen && !(e.target as HTMLElement).closest('.menu-wrap')) {
    state.emailMenuOpen = false;
    render();
  }
});
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && state.emailMenuOpen) {
    state.emailMenuOpen = false;
    render();
  }
});
render();

// Newer settings from another device (if signed in to sync) replace these once they arrive.
pullBeforeStart(SYNC_APP, STORAGE_KEY).then((result) => {
  if (result === 'pulled') applySyncedSettings();
});

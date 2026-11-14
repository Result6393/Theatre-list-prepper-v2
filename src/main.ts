import './style.css';
import { copyHtml, copyText, htmlTableRows, toTsv } from './clipboard';
import { autoGrow, downloadBlob, h } from './dom';
import { startDrag } from './drag';
import { allergyColumn, allergyPending, lacksAllergyDetails, pendingAllergies } from './allergy';
import { buildEmail, emailBodyHtml, emailBodyText, mailtoUrl, MAX_URL_LENGTH, outlookComposeUrl, weekdayLabel } from './email';
import { extractSessions, type Patient, type Session } from './extract';
import { combineSessions, listOfPatient, orderedSelection } from './lists';
import { copyBlock, copyPictures, findInSheet, findListBlock, findPreppedRows, freshRows, keepReports, sheetRows, isRectangular, mergePrepped, MergeError, parsePastedRows, withoutTrailingBlankRows, type GridRow, type MergeSummary } from './merge';
import { ParseError, parseTable, type Table } from './parse';
import { lensEmailHtml, lensEmailText, lensLines, type LensItem } from './lenslist';
import { findingGroups, hasToken, isLensSurgery, lensCode, promptColumns, readEyeOrder, readLens, readTarget, toggleFinding, writeEyeOrder, writeLens, writeTarget } from './prompts';
import Papa from 'papaparse';
import { classifyFile } from './files';
import { headingCells, type HeadingCell } from './heading';
import { openWorkbook, overwrittenBelow, saveWorkbook, sheetNameProblem, withEdit, WorkbookError, type SheetData, type SheetEdit, type Workbook } from './xlsx';
import { defaultSettings, exportSettings, importSettings, loadSettings, saveSettings, STORAGE_KEY, SYNC_APP, type KeyValueStore } from './settings';
import { mountSyncControl, pullBeforeStart, pushSoon } from './sync.js';
import { renderSettings } from './settings-view';
import { reportImage, type ReportImage } from './report';
import { checkLens, combineReadings, readReport, type LensCheck, type PickedLens, type ReportReading } from './iolreport';
import { readWords } from './ocr';

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
  /** A list straight from your prepped sheet, before any run sheet is loaded. */
  sheetOnly?: boolean;
  /** The list as it was before the file, so choosing another sheet starts afresh. */
  beforeWorkbook?: { rows: GridRow[]; padTo: number; pasteAt: string | null; summary: MergeSummary | null; undo: Grid['undo'] };
}

/** Your previously prepped sheet, chosen in step 1 or the update panel. Read on this computer only. */
interface SheetFile {
  name: string;
  /** The sheets' values, with edits made before the run sheet was loaded. */
  sheets: SheetData[];
  /** The workbook as in the file, for saving changes into it (null for a CSV). */
  book: Workbook | null;
  /** Lets the browser save over the file itself (Edge/Chrome), with your permission. */
  handle: FileSystemFileHandle | null;
  /** When the file was last changed, as we read or saved it: to notice saves made in Excel since. */
  lastModified: number | null;
}

/** How the last save went, shown under the toolbar until dismissed (a toast is too easy to miss). */
interface SaveReport {
  ok: boolean;
  text: string;
  /** Still saving (e.g. waiting for the browser's permission). */
  busy?: boolean;
  /** Buttons in the message, e.g. "Save anyway" or "Download a copy instead". */
  actions?: { label: string; primary?: boolean; run: () => void }[];
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
  saveMenuOpen: false,
  /** Phone: the ⋯ menu in the bottom bar (Update, Email). */
  moreMenuOpen: false,
  /** Phone: patient cards opened for editing, by row. */
  openCards: new Set<string>(),
  saveReport: null as SaveReport | null,
  /** The "Save as a new sheet" form, while it's open. */
  newSheet: null as { name: string; copyOf: string; error: string } | null,
  /** The "Save to a new Excel file" form, while it's open. `source` is "file:<sheet>" or "template:<sheet>". */
  newFile: null as { name: string; source: string; error: string } | null,
  /** A template workbook for new files (kept on this computer if it has no patients in it). */
  template: null as { name: string; book: Workbook } | null,
  /** Your previously prepped sheet: each list you open is updated from it. */
  workbook: null as SheetFile | null,
  /** Without a run sheet: the sheet of your prepped file being worked on. */
  sheetSelected: '' as string,
  /** Something went wrong reading a chosen file (shown in step 1). */
  fileError: '',
  /** A message under the header, e.g. after importing a settings file. */
  notice: null as { text: string; error: boolean } | null,
  /** A patient's IOL report open full screen (for timeout): the row, and which screenshot. */
  viewer: null as { key: string; index: number; actualSize: boolean } | null,
};

const app = document.getElementById('app')!;

/** Phones get patient cards and a bottom bar instead of the wide table. */
const phoneQuery = window.matchMedia('(max-width: 760px)');
const isPhone = () => phoneQuery.matches;
phoneQuery.addEventListener('change', () => render());

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
  if (text.trim()) keepSheetEdits();
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
  state.saveMenuOpen = false;
  state.moreMenuOpen = false;
  state.newSheet = null;
  state.newFile = null;
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
async function readSheets(file: File): Promise<{ sheets: SheetData[]; book: Workbook | null }> {
  if (classifyFile(file.name, file.type) === 'workbook') {
    const book = await openWorkbook(await file.arrayBuffer());
    return { sheets: [...book.sheets], book };
  }
  const rows = Papa.parse<string[]>((await file.text()).replace(/^\uFEFF/, ''), { skipEmptyLines: false }).data;
  return { sheets: [{ name: file.name, rows }], book: null };
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
async function openFile(file: File, as: 'any' | 'runsheet' | 'prepped' = 'any', handle: FileSystemFileHandle | null = null): Promise<void> {
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
  let book: Workbook | null;
  try {
    ({ sheets, book } = await readSheets(file));
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
  state.workbook = { name: file.name, sheets, book, handle, lastModified: handle ? file.lastModified : null };
  state.saveReport = null;
  if (!state.table) {
    // No run sheet yet: open the prepped sheet's lists straight away.
    for (const [key, grid] of state.grids) if (grid.sheetOnly) state.grids.delete(key);
    const lists = sheetLists();
    state.sheetSelected = lists.length === 1 ? lists[0].sheet : '';
    openSheet();
    say(
      lists.length
        ? `Opened ${file.name}. Edit it here; load the latest run sheet when you have it to update the list${lists.length > 1 ? 's' : ''}.`
        : `${file.name} has no patient rows (MRNs) that the app can find. Check the Patient column in Settings matches your sheet.`,
      false,
    );
  } else if (state.grid) applyWorkbook(state.grid, state.workbook);
  else say(`Prepped sheet ${file.name} ready: choose a list to update it.`, false);
  render();
}

/** The lists on your prepped sheet (one block of patient rows per sheet), for working without a run sheet. */
function sheetLists() {
  const wb = state.workbook;
  if (!wb) return [];
  return wb.sheets.map((sheet) => findListBlock(sheet, state.settings.layout)).filter((b): b is NonNullable<typeof b> => b !== null);
}

/** Opens a sheet of your prepped file as a list (no run sheet loaded). */
function openSheet(): void {
  const wb = state.workbook;
  const block = sheetLists().find((b) => b.sheet === state.sheetSelected);
  state.updateOpen = false;
  state.emailMenuOpen = false;
  state.saveMenuOpen = false;
  state.moreMenuOpen = false;
  state.newSheet = null;
  state.newFile = null;
  if (!wb || !block) {
    state.grid = null;
    return;
  }
  const key = `sheet:${block.sheet}`;
  let grid = state.grids.get(key);
  if (!grid || grid.fromWorkbook !== wb) {
    const combined: Session = {
      id: key, theatre: '', or: '', date: '', start: '', end: '', ampm: '', slotName: block.sheet, slotOwner: '', consultant: '', patients: [],
    };
    grid = {
      sessions: [], combined, listOf: new Map(), rows: sheetRows(block.rows, key, block.pictures), padTo: block.rows.length, pasteAt: block.cell,
      summary: null, undo: null, fromWorkbook: wb, fromSheet: block.sheet, sheetOnly: true,
    };
    state.grids.set(key, grid);
  }
  state.grid = grid;
}

function chooseSheet(name: string): void {
  state.sheetSelected = name;
  state.status = '';
  openSheet();
  render();
}

/** Before the run sheet replaces the sheet-only lists, keep their edits as your sheet's values. */
function keepSheetEdits(): void {
  const wb = state.workbook;
  if (!wb) return;
  for (const grid of state.grids.values()) {
    if (!grid.sheetOnly || grid.fromWorkbook !== wb || !grid.fromSheet || !grid.pasteAt) continue;
    const at = wb.sheets.findIndex((s) => s.name === grid.fromSheet);
    if (at === -1) continue;
    wb.sheets[at] = withEdit(wb.sheets[at], blockEdit(grid, grid.fromSheet, grid.pasteAt, grid.padTo));
  }
  state.sheetSelected = '';
}

type Permission = 'granted' | 'denied' | 'prompt';
interface Permissions {
  queryPermission?: (d: { mode: 'readwrite' }) => Promise<Permission>;
  requestPermission?: (d: { mode: 'readwrite' }) => Promise<Permission>;
}
interface SaveFilePicker {
  (options: { suggestedName?: string; types?: { description: string; accept: Record<string, string[]> }[] }): Promise<FileSystemFileHandle>;
}
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * Asks to save over the file, straight away while the click still counts as yours: browsers only
 * show that question in answer to a click, and preparing the workbook takes a moment.
 */
function askToSave(wb: SheetFile): Promise<Permission | 'no-file'> {
  const handle = wb.handle as (FileSystemFileHandle & Permissions) | null;
  if (!handle) return Promise.resolve('no-file');
  if (!handle.requestPermission) return Promise.resolve('granted');
  return handle.requestPermission({ mode: 'readwrite' }).catch(() => 'prompt' as const);
}

/** Writes the bytes into a file; on failure, says why in words. */
async function writeTo(handle: FileSystemFileHandle, bytes: Uint8Array): Promise<string | null> {
  try {
    const writable = await handle.createWritable();
    await writable.write(bytes as BlobPart);
    await writable.close();
    return null;
  } catch (e) {
    const name = e instanceof DOMException ? e.name : '';
    if (name === 'NotAllowedError' || name === 'SecurityError') return 'the browser wasn’t allowed to change it';
    if (name === 'NoModificationAllowedError' || name === 'InvalidStateError' || name === 'InvalidModificationError') {
      return 'it looks to be open in Excel, or OneDrive was busy with it';
    }
    return `the browser couldn’t write to it${name ? ` (${name})` : ''}`;
  }
}

/**
 * Asks in the page rather than with a browser dialog (those can be hidden or blocked, and then the
 * save would stop with nothing on screen). Resolves with the chosen option's value.
 */
function askInPage<T>(text: string, options: { label: string; value: T; primary?: boolean }[]): Promise<T> {
  return new Promise((resolve) => {
    state.saveReport = {
      ok: false,
      text,
      actions: options.map((o) => ({ label: o.label, primary: o.primary, run: () => resolve(o.value) })),
    };
    render();
  });
}

function savingNotice(text: string): void {
  state.saveReport = { ok: true, busy: true, text };
  render();
}

type SaveOutcome = { how: 'over' | 'as'; name: string; reason?: string } | { how: 'downloaded'; why: string } | { how: 'cancelled' };

/**
 * Puts the saved workbook where you'll find it: over the file you opened when the browser allows
 * it; otherwise in a file you pick (choose the original to replace it; later saves go there too);
 * and only when the browser can do neither, as a download. The saved version is then what later
 * saves compare against.
 */
async function writeOut(wb: SheetFile, bytes: Uint8Array, permission: Promise<Permission | 'no-file'>): Promise<SaveOutcome> {
  let why = '';
  /** Why it couldn't save over the file, as a short phrase. */
  let reason = '';
  // The browser's "allow this page to edit the file?" question is easy to miss, so say so, and offer a way out.
  let giveUp!: (value: 'download') => void;
  const gaveUp = new Promise<'download'>((resolve) => (giveUp = resolve));
  const waiting = setTimeout(() => {
    state.saveReport = {
      ok: false,
      busy: true,
      text: `Waiting for your browser to allow saving over ${wb.name}. Look for its question (often a small box by the address bar, or behind this window) and choose Allow or Edit file.`,
      actions: [{ label: 'Download a copy instead', run: () => giveUp('download') }],
    };
    render();
  }, 1200);
  const allowed = await Promise.race([permission, gaveUp]);
  clearTimeout(waiting);
  if (allowed === 'download') {
    downloadBlob(wb.name, new Blob([bytes as BlobPart], { type: XLSX_TYPE }));
    return finishSave(wb, bytes, { how: 'downloaded', why: 'You chose to download a copy rather than wait for the browser.' });
  }
  savingNotice(`Saving into ${wb.name}…`);
  if (wb.handle && allowed !== 'denied') {
    // Saving over changes made in Excel since would lose them, so check first.
    const onDisk = await wb.handle.getFile().catch(() => null);
    if (onDisk && wb.lastModified && onDisk.lastModified !== wb.lastModified) {
      const go = await askInPage(`${wb.name} has been changed since you opened it here (perhaps saved in Excel). Saving replaces those changes with this version.`, [
        { label: 'Save anyway', value: true, primary: true },
        { label: 'Cancel', value: false },
      ]);
      if (!go) return { how: 'cancelled' };
      savingNotice(`Saving into ${wb.name}…`);
    }
    const problem = await writeTo(wb.handle, bytes);
    if (!problem) return finishSave(wb, bytes, { how: 'over', name: wb.name });
    why = `Couldn’t save over ${wb.name}: ${problem}.${/Excel/.test(problem) ? ' Close it in Excel and press Save again to save over it.' : ''}`;
    reason = `couldn’t save over ${wb.name}: ${problem}`;
  } else {
    why = wb.handle ? `You didn’t allow saving over ${wb.name}.` : `The browser can only save over a file you chose with “Choose prepped Excel sheet…” in Edge or Chrome.`;
    reason = wb.handle ? `you didn’t allow saving over ${wb.name}` : `${wb.name} wasn’t opened in a way that lets the browser save over it`;
  }
  const picker = (window as unknown as { showSaveFilePicker?: SaveFilePicker }).showSaveFilePicker;
  if (picker) {
    // The browser only opens its Save window straight after a click: if that moment has passed, ask for one.
    const activation = (navigator as Navigator & { userActivation?: { isActive: boolean } }).userActivation;
    if (activation && !activation.isActive) {
      const next = await askInPage(`${why} Choose where to save it: pick ${wb.name} itself to replace it.`, [
        { label: 'Choose where to save…', value: 'pick' as const, primary: true },
        { label: 'Download a copy', value: 'download' as const },
        { label: 'Cancel', value: 'cancel' as const },
      ]);
      if (next === 'cancel') return { how: 'cancelled' };
      if (next === 'download') {
        downloadBlob(wb.name, new Blob([bytes as BlobPart], { type: XLSX_TYPE }));
        return finishSave(wb, bytes, { how: 'downloaded', why });
      }
    }
    try {
      const handle = await picker({ suggestedName: wb.name, types: [{ description: 'Excel workbook', accept: { [XLSX_TYPE]: ['.xlsx'] } }] });
      const problem = await writeTo(handle, bytes);
      if (!problem) {
        wb.handle = handle;
        wb.name = handle.name;
        return finishSave(wb, bytes, { how: 'as', name: handle.name, reason });
      }
      why += ` Couldn’t save as ${handle.name} either: ${problem}.`;
    } catch (e) {
      if (e instanceof DOMException && e.name === 'AbortError') return { how: 'cancelled' };
      // The browser won't show a save dialog here: download instead.
    }
  }
  downloadBlob(wb.name, new Blob([bytes as BlobPart], { type: XLSX_TYPE }));
  return finishSave(wb, bytes, { how: 'downloaded', why });
}

async function finishSave(wb: SheetFile, bytes: Uint8Array, outcome: SaveOutcome): Promise<SaveOutcome> {
  wb.book = await openWorkbook(bytes.buffer as ArrayBuffer);
  wb.sheets = [...wb.book.sheets];
  wb.lastModified = wb.handle && outcome.how !== 'downloaded' ? ((await wb.handle.getFile().catch(() => null))?.lastModified ?? null) : wb.lastModified;
  return outcome;
}

/** Says plainly where the save went, and keeps saying it until dismissed. */
function reportSave(outcome: SaveOutcome, what: string, extra = ''): void {
  const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  const wb = state.workbook;
  if (outcome.how === 'cancelled') {
    state.saveReport = { ok: false, text: 'Not saved.' };
  } else if (outcome.how === 'downloaded') {
    state.saveReport = {
      ok: false,
      text: `${what} saved into a new copy of ${wb?.name ?? 'the file'} in your Downloads folder (it may be named like “${(wb?.name ?? 'file.xlsx').replace(/\.xlsx$/i, ' (1).xlsx')}”). The file you opened hasn’t changed: replace it with the downloaded copy (e.g. upload it to OneDrive). ${outcome.why}${extra}`,
    };
  } else {
    state.saveReport = {
      ok: true,
      text: `${outcome.how === 'as' && outcome.reason ? `${outcome.reason[0].toUpperCase()}${outcome.reason.slice(1)}, so you chose where to save. ` : ''}${what} saved into ${outcome.name} at ${time}.${outcome.how === 'as' ? ' Later saves go straight to this file.' : ''} If it’s open in Excel, close it there (without saving) and open it again to see the changes.${extra}`,
    };
  }
  state.status = '';
  render();
}

/**
 * The sheet's Consultant, Date, OR and AM/PM, filled in from the run sheet (lists from the run sheet
 * only: a list straight from your file has nothing newer to put there).
 */
function headingFor(grid: Grid, sheet: SheetData | undefined, sheetName: string, listCell: string): { cells: HeadingCell[]; edits: SheetEdit[] } {
  if (!sheet || grid.sheetOnly || !grid.sessions.length) return { cells: [], edits: [] };
  const { consultant, date, or, ampm, start, end } = grid.combined;
  // The hospital is the run sheet's theatre without its OR, e.g. "SSEH" from "SSEH OR 01".
  const hospital = [...new Set(grid.sessions.map((x) => x.theatre.replace(/\s*\bOR\s*0*\d+.*$/i, '').trim()).filter(Boolean))].join(' + ');
  const times = start && end ? `${start}–${end}` : '';
  const listRow = Number(listCell.match(/\d+/)?.[0] ?? 1) - 1;
  const cells = headingCells(sheet, listRow, { consultant, date, or, ampm, hospital, times }, state.settings.layout);
  return { cells, edits: cells.map((c) => ({ sheet: sheetName, cell: c.cell, rows: [[c.value]], numbers: c.field === 'or' })) };
}

/** " Heading filled in: Consultant Willow, Date 14/10/2026, OR 4, AM/PM AM." */
function headingNote(cells: HeadingCell[]): string {
  return cells.length ? ` Heading filled in: ${cells.map((c) => `${c.label} ${c.shown}`).join(', ')}.` : '';
}

/** A name for a new sheet for this list, e.g. "OR 4 AM 14-10", not already used. */
function defaultSheetName(grid: Grid, existing: string[]): string {
  const c = grid.combined;
  const day = (c.date.split(' + ')[0] ?? '').replace(/^(\d{1,2})\/(\d{1,2})\/\d{4}$/, '$1-$2');
  const base = (grid.sheetOnly ? `${c.slotName} (new)` : [c.or, c.ampm, day].filter(Boolean).join(' ')).replace(/[\\/?*[\]:]/g, '-').slice(0, 31).trim() || 'New list';
  let name = base;
  for (let n = 2; existing.some((e) => e.toLowerCase() === name.toLowerCase()); n++) name = `${base.slice(0, 26)} (${n})`;
  return name;
}

/** Where the list goes on a new sheet copied from `template`: in place of the template's own patients. */
function newSheetPlace(grid: Grid, template: SheetData, sameFile = true): { cell: string; oldLength: number } {
  if (sameFile && grid.fromSheet === template.name && grid.pasteAt) return { cell: grid.pasteAt, oldLength: grid.padTo };
  const block = findListBlock(template, state.settings.layout);
  if (block) return { cell: block.cell, oldLength: block.rows.length };
  return { cell: state.settings.layout.pasteStart || 'B5', oldLength: 0 };
}

// ---- Save to a new Excel file -------------------------------------------------------

const TEMPLATE_KEY = 'theatre-list-prepper.template';

function toBase64(bytes: Uint8Array): string {
  let text = '';
  for (let i = 0; i < bytes.length; i += 0x8000) text += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(text);
}

/** A template file chosen before, if this computer kept it. */
async function loadStoredTemplate(): Promise<void> {
  try {
    const raw = storage?.getItem(TEMPLATE_KEY);
    if (!raw) return;
    const { name, data } = JSON.parse(raw) as { name: string; data: string };
    const bytes = Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
    state.template = { name, book: await openWorkbook(bytes.buffer as ArrayBuffer) };
  } catch {
    // Nothing kept, or it can't be read any more: choose it again when needed.
  }
}

function hasPatients(book: Workbook): boolean {
  try {
    return book.sheets.some((s) => findListBlock(s, state.settings.layout));
  } catch {
    return false;
  }
}

/** The sheet to copy for a new list: one with the list's column headings and no patients, if there is one. */
function bestTemplateSheet(book: Workbook): string {
  const headings = new Set(state.settings.layout.columns.map((c) => c.heading.trim().toLowerCase()).filter(Boolean));
  const hasHeadings = (s: SheetData) => s.rows.some((r) => r.filter((c) => headings.has(c.trim().toLowerCase())).length >= 3);
  const empty = (s: SheetData) => {
    try {
      return !findListBlock(s, state.settings.layout);
    } catch {
      return true;
    }
  };
  return (book.sheets.find((s) => hasHeadings(s) && empty(s)) ?? book.sheets.find(hasHeadings) ?? book.sheets[0])?.name ?? '';
}

/** Where a new file's layout can come from: the loaded file's sheets, and the template's. */
function newFileSources(): { key: string; label: string; book: Workbook; sheet: string; sameFile: boolean }[] {
  const out: { key: string; label: string; book: Workbook; sheet: string; sameFile: boolean }[] = [];
  const wb = state.workbook;
  if (wb?.book) for (const s of wb.book.sheets) out.push({ key: `file:${s.name}`, label: `${s.name} (in ${wb.name})`, book: wb.book, sheet: s.name, sameFile: true });
  const tpl = state.template;
  if (tpl) for (const s of tpl.book.sheets) out.push({ key: `template:${s.name}`, label: `${s.name} (template ${tpl.name})`, book: tpl.book, sheet: s.name, sameFile: false });
  return out;
}

function defaultNewFileSource(grid: Grid): string {
  const wb = state.workbook;
  if (wb?.book && grid.fromWorkbook === wb && grid.fromSheet) return `file:${grid.fromSheet}`;
  if (state.template) return `template:${bestTemplateSheet(state.template.book)}`;
  if (wb?.book) return `file:${bestTemplateSheet(wb.book)}`;
  return '';
}

function openNewFile(grid: Grid): void {
  state.saveMenuOpen = false;
  state.newSheet = null;
  const base = grid.sheetOnly ? grid.combined.slotName : `Theatre list ${defaultSheetName(grid, [])}`;
  state.newFile = { name: `${base.replace(/[\\/:*?"<>|]/g, '-')}.xlsx`, source: defaultNewFileSource(grid), error: '' };
  render();
  app.querySelector<HTMLInputElement>('.new-file input')?.select();
}

/** Chooses a template file for new lists. It's kept on this computer, unless it has patients in it. */
function chooseTemplate(): void {
  pickFile('.xlsx,.xlsm', async (file) => {
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const book = await openWorkbook(bytes.buffer as ArrayBuffer);
      state.template = { name: file.name, book };
      const kept = !hasPatients(book);
      if (kept) {
        try {
          storage?.setItem(TEMPLATE_KEY, JSON.stringify({ name: file.name, data: toBase64(bytes) }));
        } catch {
          // Storage full or blocked: it's used this time only.
        }
      }
      if (state.newFile) {
        state.newFile.source = `template:${bestTemplateSheet(book)}`;
        state.newFile.error = kept ? '' : `${file.name} has patients in it, so it’s used this time but not kept on this computer.`;
      }
    } catch (e) {
      if (state.newFile) state.newFile.error = e instanceof WorkbookError ? `${file.name}: ${e.message}` : `Couldn’t read ${file.name}.`;
    }
    render();
  });
}

function forgetTemplate(): void {
  state.template = null;
  try {
    storage?.removeItem(TEMPLATE_KEY);
  } catch {
    // Nothing kept.
  }
  if (state.newFile) state.newFile.source = defaultNewFileSource(state.grid!);
  render();
}

/** Saves a new file (Save As where the browser can, else a download). */
async function writeNewFile(name: string, bytes: Uint8Array): Promise<{ how: 'as'; name: string; handle: FileSystemFileHandle } | { how: 'downloaded'; name: string } | { how: 'cancelled' }> {
  const picker = (window as unknown as { showSaveFilePicker?: SaveFilePicker }).showSaveFilePicker;
  if (picker) {
    const activation = (navigator as Navigator & { userActivation?: { isActive: boolean } }).userActivation;
    if (activation && !activation.isActive) {
      const next = await askInPage(`${name} is ready.`, [
        { label: 'Choose where to save…', value: 'pick' as const, primary: true },
        { label: 'Download it', value: 'download' as const },
        { label: 'Cancel', value: 'cancel' as const },
      ]);
      if (next === 'cancel') return { how: 'cancelled' };
      if (next === 'download') {
        downloadBlob(name, new Blob([bytes as BlobPart], { type: XLSX_TYPE }));
        return { how: 'downloaded', name };
      }
    }
    try {
      const handle = await picker({ suggestedName: name, types: [{ description: 'Excel workbook', accept: { [XLSX_TYPE]: ['.xlsx'] } }] });
      const problem = await writeTo(handle, bytes);
      if (!problem) return { how: 'as', name: handle.name, handle };
    } catch (e) {
      if (e instanceof DOMException && e.name === 'AbortError') return { how: 'cancelled' };
    }
  }
  downloadBlob(name, new Blob([bytes as BlobPart], { type: XLSX_TYPE }));
  return { how: 'downloaded', name };
}

/**
 * Saves this list in a new Excel file: a copy of the template (or the chosen sheet's file) with only
 * that sheet and any sheets without lists (like instructions), the list in place of its patients,
 * and its heading filled in. With no file loaded, the new file becomes this list's file.
 */
async function saveToNewFile(grid: Grid): Promise<void> {
  if (!(await allergiesChecked([grid], 'Save'))) return;
  try {
    await saveToNewFileNow(grid);
  } catch (e) {
    saveFailed(e);
  }
}

async function saveToNewFileNow(grid: Grid): Promise<void> {
  const form = state.newFile;
  if (!form) return;
  const source = newFileSources().find((s) => s.key === form.source);
  const name = form.name.trim().replace(/[\\/:*?"<>|]/g, '-');
  if (!source || !name) {
    form.error = !source ? 'Choose a template, or a sheet to copy the layout of.' : 'Give the new file a name.';
    render();
    return;
  }
  const fileName = /\.xls[xm]$/i.test(name) ? name : `${name}.xlsx`;
  savingNotice(`Making ${fileName}…`);
  const { book, sheet } = source;
  const template = book.sheets.find((s) => s.name === sheet)!;
  const { cell, oldLength } = newSheetPlace(grid, template, source.sameFile);
  const edit = blockEdit(grid, sheet, cell, oldLength);
  const heading = headingFor(grid, template, sheet, cell);
  const keep = book.sheets
    .filter((s) => {
      if (s.name === sheet) return true;
      try {
        return !findListBlock(s, state.settings.layout);
      } catch {
        return true;
      }
    })
    .map((s) => s.name);
  let result;
  try {
    result = await saveWorkbook(book, [edit, ...heading.edits], [], { keepOnly: keep, activeSheet: sheet });
  } catch (e) {
    form.error = e instanceof WorkbookError ? e.message : `Couldn’t make the new file${e instanceof Error ? ` (${e.message})` : ''}.`;
    state.saveReport = null;
    render();
    return;
  }
  const outcome = await writeNewFile(fileName, result.bytes);
  if (outcome.how === 'cancelled') {
    state.saveReport = { ok: false, text: 'Not saved.' };
    render();
    return;
  }
  state.newFile = null;
  const where = `It’s on sheet “${sheet}”, from ${cell}.${headingNote(heading.cells)}`;
  // With no file loaded, the new one becomes this list's file, so Save keeps it up to date.
  let adopted = '';
  if (!state.workbook) {
    const newBook = await openWorkbook(result.bytes.buffer as ArrayBuffer);
    const handle = outcome.how === 'as' ? outcome.handle : null;
    const wb: SheetFile = { name: outcome.name, sheets: [...newBook.sheets], book: newBook, handle, lastModified: handle ? ((await handle.getFile().catch(() => null))?.lastModified ?? null) : null };
    state.workbook = wb;
    for (const g of state.grids.values()) g.fromWorkbook = wb;
    grid.fromSheet = sheet;
    grid.pasteAt = cell;
    grid.padTo = edit.rows.length;
    adopted = ` From now on, Save updates ${outcome.name}${handle ? '' : ' (choose the downloaded file when asked)'}.`;
  }
  state.saveReport =
    outcome.how === 'as'
      ? { ok: true, text: `Saved this list as a new file, ${outcome.name}. ${where}${adopted}` }
      : { ok: false, text: `This list was saved as a new file, ${outcome.name}, in your Downloads folder (it may be named like “${outcome.name.replace(/\.xls([xm])$/i, ' (1).xls$1')}”). ${where}${adopted}` };
  render();
}

function openNewSheet(grid: Grid): void {
  const book = state.workbook?.book;
  if (!book) return;
  const names = book.sheets.map((s) => s.name);
  const withList = book.sheets.find((s) => findListBlock(s, state.settings.layout));
  const copyOf = grid.fromSheet && names.includes(grid.fromSheet) ? grid.fromSheet : (withList ?? book.sheets[0])?.name ?? '';
  state.saveMenuOpen = false;
  state.newFile = null;
  state.newSheet = { name: defaultSheetName(grid, names), copyOf, error: '' };
  render();
  app.querySelector<HTMLInputElement>('.new-sheet input')?.select();
}

/** Saves this list as a new sheet in your Excel file, copying another sheet's layout. */
async function saveAsNewSheet(grid: Grid): Promise<void> {
  if (!(await allergiesChecked([grid], 'Save'))) return;
  try {
    await saveAsNewSheetNow(grid);
  } catch (e) {
    saveFailed(e);
  }
}

/** Something unexpected went wrong while saving: say so, rather than nothing happening. */
function saveFailed(e: unknown): void {
  console.error(e);
  const message = e instanceof Error ? e.message : String(e);
  state.saveReport = { ok: false, text: `Saving didn’t work (${message}). Nothing was changed in your file. Try again, or use Copy for Excel instead.` };
  render();
}

async function saveAsNewSheetNow(grid: Grid): Promise<void> {
  const wb = state.workbook;
  const form = state.newSheet;
  if (!wb?.book || !form) return;
  const permission = askToSave(wb);
  savingNotice(`Saving a new sheet into ${wb.name}…`);
  const book = wb.book;
  const name = form.name.trim();
  const problem = sheetNameProblem(name, book.sheets.map((s) => s.name));
  const template = book.sheets.find((s) => s.name === form.copyOf);
  if (problem || !template) {
    form.error = problem ?? 'Choose a sheet to copy the layout from.';
    state.saveReport = null;
    render();
    return;
  }
  const { cell, oldLength } = newSheetPlace(grid, template);
  const edit = blockEdit(grid, name, cell, oldLength);
  const lost = overwrittenBelow(template, edit, oldLength);
  if (lost.length) {
    const go = await askInPage(
      `On the new sheet the list is longer than on “${template.name}”, so it writes over ${lost.length} cell${lost.length === 1 ? '' : 's'} copied from below it: ${lost.slice(0, 3).map((t) => `“${t.split('\n')[0]}”`).join(', ')}${lost.length > 3 ? '…' : ''}.`,
      [{ label: 'Save anyway', value: true, primary: true }, { label: 'Cancel', value: false }],
    );
    if (!go) {
      reportSave({ how: 'cancelled' }, '');
      return;
    }
  }
  let result;
  let heading: { cells: HeadingCell[]; edits: SheetEdit[] } = { cells: [], edits: [] };
  try {
    heading = headingFor(grid, template, name, cell);
    result = await saveWorkbook(book, [edit, ...heading.edits], [{ name, copyOf: template.name }]);
  } catch (e) {
    form.error = e instanceof WorkbookError ? e.message : `Couldn’t save into ${wb.name}${e instanceof Error ? ` (${e.message})` : ''}.`;
    state.saveReport = null;
    render();
    return;
  }
  const outcome = await writeOut(wb, result.bytes, permission);
  if (outcome.how === 'cancelled') {
    reportSave(outcome, '');
    return;
  }
  // From now on this list saves into its new sheet.
  if (grid.sheetOnly) {
    for (const [key, g] of state.grids) if (g === grid) state.grids.delete(key);
    grid.combined = { ...grid.combined, id: `sheet:${name}`, slotName: name };
    state.grids.set(`sheet:${name}`, grid);
    state.sheetSelected = name;
  }
  grid.fromWorkbook = wb;
  grid.fromSheet = name;
  grid.pasteAt = cell;
  grid.padTo = edit.rows.length;
  state.newSheet = null;
  reportSave(outcome, `This list was`, ` It’s on the new sheet “${name}”, from ${cell}.${headingNote(heading.cells)}`);
}

/** A list's rows (and each patient's IOL report pictures) to write into a sheet from `cell`. */
function blockEdit(grid: Grid, sheet: string, cell: string, padTo: number): SheetEdit {
  return { sheet, cell, rows: copyBlock(grid.rows, state.settings.layout, padTo), pictures: copyPictures(grid.rows, padTo) };
}

/** Lists from your prepped file, each with the block it writes back: the most recent wins per sheet. */
function workbookEdits(wb: SheetFile): { grid: Grid; edit: SheetEdit }[] {
  const bySheet = new Map<string, { grid: Grid; edit: SheetEdit }>();
  for (const grid of state.grids.values()) {
    if (grid.fromWorkbook !== wb || !grid.fromSheet || !grid.pasteAt) continue;
    bySheet.set(grid.fromSheet, { grid, edit: blockEdit(grid, grid.fromSheet, grid.pasteAt, grid.padTo) });
  }
  return [...bySheet.values()];
}

/**
 * Saves your lists into the prepped Excel file: only the cells that changed in each list's block.
 * Saves over the file itself when the browser allows it (with your permission), otherwise downloads
 * an updated copy with the same name.
 */
async function saveToWorkbook(): Promise<void> {
  if (state.workbook && !(await allergiesChecked(workbookEdits(state.workbook).map((e) => e.grid), 'Save'))) return;
  try {
    await saveToWorkbookNow();
  } catch (e) {
    saveFailed(e);
  }
}

async function saveToWorkbookNow(): Promise<void> {
  const wb = state.workbook;
  if (!wb?.book) return;
  const permission = askToSave(wb);
  savingNotice(`Saving into ${wb.name}…`);
  const book = wb.book;
  const edits = workbookEdits(wb);
  const headings = edits.map(({ grid, edit }) => headingFor(grid, book.sheets.find((s) => s.name === edit.sheet), edit.sheet, edit.cell));
  if (!edits.length) {
    state.saveReport = { ok: false, text: 'None of your lists came from this file yet, so there’s nothing to save into it. Use “Save as a new sheet…” (▾) instead.' };
    render();
    return;
  }
  const lost = edits.flatMap(({ grid, edit }) => {
    const sheet = book.sheets.find((s) => s.name === edit.sheet);
    return sheet ? overwrittenBelow(sheet, edit, grid.padTo).map((text) => `${edit.sheet}: ${text.split('\n')[0]}`) : [];
  });
  if (lost.length) {
    const go = await askInPage(
      `The list is longer than before, so saving writes over ${lost.length} cell${lost.length === 1 ? '' : 's'} below it: ${lost.slice(0, 3).map((t) => `“${t}”`).join(', ')}${lost.length > 3 ? '…' : ''}.`,
      [{ label: 'Save anyway', value: true, primary: true }, { label: 'Cancel', value: false }],
    );
    if (!go) {
      reportSave({ how: 'cancelled' }, '');
      return;
    }
  }

  let result;
  try {
    result = await saveWorkbook(book, [...edits.map((e) => e.edit), ...headings.flatMap((h) => h.edits)]);
  } catch (e) {
    state.saveReport = { ok: false, text: e instanceof WorkbookError ? e.message : `Couldn’t save into ${wb.name}${e instanceof Error ? ` (${e.message})` : ''}.` };
    render();
    return;
  }
  if (result.changed === 0) {
    state.saveReport = { ok: true, text: `Nothing to save: ${wb.name} already has everything on your list${edits.length > 1 ? 's' : ''}.` };
    render();
    return;
  }
  const outcome = await writeOut(wb, result.bytes, permission);
  if (outcome.how === 'cancelled') {
    reportSave(outcome, '');
    return;
  }
  for (const { grid, edit } of edits) grid.padTo = edit.rows.length;
  const changes = `${result.changed} change${result.changed === 1 ? '' : 's'}`;
  const sheets = edits.map((e) => `“${e.edit.sheet}”`).join(', ');
  const formulas = result.formulas ? ` Formula cells were left as they were (${result.formulas}).` : '';
  reportSave(outcome, `${changes[0].toUpperCase()}${changes.slice(1)} (sheet${edits.length === 1 ? '' : 's'} ${sheets})`, `${formulas}${headingNote(headings.flatMap((h) => h.cells))}`);
}

/**
 * Updates a list from your prepped sheet. Without a sheet named, the sheet that fits this list best
 * is used (most of its patients, fewest others); null means don't use the file.
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
  const isDefault = sheetName === undefined && workbook.sheets.length > 1;
  let found;
  try {
    const candidates = sheetName === undefined ? workbook.sheets : workbook.sheets.filter((s) => s.name === sheetName);
    found = findPreppedRows(candidates, grid.combined, state.settings.layout);
  } catch (e) {
    say(e instanceof MergeError ? e.message : `Couldn’t read ${workbook.name}.`, false);
    return;
  }
  if (!found) {
    say(
      workbook.sheets.length > 1
        ? isDefault
          ? `None of the patients on ${listName(grid.combined)} are in ${workbook.name}, so the list is as on the run sheet.`
          : `None of the patients on ${listName(grid.combined)} are on sheet “${sheetName}”, so the list is as on the run sheet.`
        : `None of the patients on ${listName(grid.combined)} are in ${workbook.name}, so the list is as on the run sheet.`,
      false,
    );
    return;
  }
  const error = mergeInto(grid, found.rows, found.cell, found.pictures);
  if (error) {
    say(error, false);
    return;
  }
  grid.fromSheet = found.sheet;
  const where = workbook.sheets.length > 1 ? `sheet “${found.sheet}”, from ${found.cell}` : `from ${found.cell}`;
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
  for (const [key, grid] of state.grids) if (grid.sheetOnly) state.grids.delete(key);
  if (state.grid?.sheetOnly) state.grid = null;
  state.sheetSelected = '';
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
function mergeInto(grid: Grid, rows: string[][], pasteAt: string | null, pictures: ReportImage[][] = []): string | null {
  try {
    const others = listedSessions().filter((s) => !grid.sessions.includes(s));
    const result = mergePrepped(rows, grid.combined, others, state.settings.layout, listName, pictures);
    // Pasted rows carry no pictures, so reports already added stay with their patients.
    keepReports(result.rows, grid.rows, state.settings.layout);
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


/** Ticked cataract patients with what's in their Implants box, or null if prompts/lens column aren't on your sheet. */
function lensItems(grid: Grid): LensItem[] | null {
  const { layout, prompts } = state.settings;
  const lensCol = prompts.enabled ? promptColumns(layout, prompts).lens : -1;
  if (lensCol === -1) return null;
  const eyeCol = layout.columns.findIndex((c) => c.source === 'field:eye');
  const keyCol = layout.columns.findIndex((c) => c.source === 'field:patient' || c.source === 'field:name' || c.source === 'field:mrn');
  return activeRows(grid)
    .filter((r) => r.included && isLensRow(r))
    .map((r) => ({
      who: r.patient ? [r.patient.mrn, r.patient.name].filter(Boolean).join(' ') : (keyCol >= 0 ? r.cells[keyCol] : '').split('\n')[0],
      eye: eyeCol >= 0 ? r.cells[eyeCol] : r.patient?.eye ?? '',
      cell: r.cells[lensCol] ?? '',
    }));
}

/** Copies the whole email to the theatre coordinator: a table of each ticked cataract patient's lens, then how many of each are needed. */
async function copyLensList(grid: Grid): Promise<void> {
  const items = lensItems(grid);
  if (!items) {
    say('Turn on the cataract prompts, with an Implants column, in Settings to copy a lens list.');
    return;
  }
  const { lines, missing } = lensLines(items, state.settings.prompts.lenses);
  if (!lines.length) {
    say(items.length ? 'No lenses picked yet: choose one under each patient (or type it in the Implants box).' : 'No cataract patients are ticked.');
    return;
  }
  const { combined } = grid;
  const info = {
    list: combined.consultant || combined.slotName || [combined.or, combined.ampm].filter(Boolean).join(' '),
    day: combined.date ? weekdayLabel(combined.date) : '',
    name: state.settings.email.name,
  };
  try {
    await copyHtml(lensEmailHtml(lines, info), lensEmailText(lines, info));
    say(
      `Copied the lens email (${lines.length} lens${lines.length === 1 ? '' : 'es'}) for the theatre coordinator: paste it into a new email.` +
        (missing.length ? ` Left out, no lens picked: ${missing.map((m) => m.who).join('; ')}.` : ''),
    );
  } catch {
    say('Your browser blocked copying. Try again.');
  }
}

/**
 * Before a list leaves the app (copy or save), asks if ticked patients still have an allergy with no
 * details. Asked in the page, like the save questions, so it can't be hidden by the browser.
 */
async function allergiesChecked(grids: Grid[], what: string): Promise<boolean> {
  const col = allergyColumn(state.settings.layout);
  const waiting = grids.reduce((n, g) => n + pendingAllergies(activeRows(g), col).length, 0);
  if (!waiting) return true;
  const go = await askInPage(
    `${waiting} ticked patient${waiting === 1 ? ' has' : 's have'} an allergy with no details yet (red in the list). ${what} anyway?`,
    [
      { label: 'Go back and check', value: false, primary: true },
      { label: `${what} anyway`, value: true },
    ],
  );
  state.saveReport = null;
  render();
  return go;
}

async function copyRows(grid: Grid): Promise<void> {
  if (!(await allergiesChecked([grid], 'Copy'))) return;
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
    ...[state.view === 'main' ? renderViewer() : null].filter((v): v is HTMLElement => v !== null),
    h(
      'div',
      { class: 'toast-area', role: 'status', 'aria-live': 'polite' },
      state.status
        ? h('div', { class: 'toast' }, h('span', {}, state.status), h('button', { class: 'toast-close', title: 'Dismiss', onclick: () => ((state.status = ''), render()) }, '×'))
        : null,
    ),
  );
  for (const area of app.querySelectorAll<HTMLTextAreaElement>('textarea.cell')) autoGrow(area);
  document.body.classList.toggle('viewing', Boolean(state.viewer));
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

interface OpenFilePicker {
  (options: { types?: { description: string; accept: Record<string, string[]> }[]; multiple?: boolean }): Promise<FileSystemFileHandle[]>;
}

/**
 * Chooses your prepped file. Where the browser can (Edge/Chrome), this keeps a handle to the file so
 * changes can be saved back into it; otherwise it's an ordinary file chooser.
 */
async function choosePrepped(): Promise<void> {
  const picker = (window as unknown as { showOpenFilePicker?: OpenFilePicker }).showOpenFilePicker;
  if (picker) {
    try {
      const [handle] = await picker({
        types: [{ description: 'Excel workbook or CSV', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'], 'text/csv': ['.csv'] } }],
      });
      await openFile(await handle.getFile(), 'prepped', handle);
      return;
    } catch (e) {
      if (e instanceof DOMException && e.name === 'AbortError') return;
      // Not allowed here (e.g. some file:// pages): use the ordinary chooser.
    }
  }
  pickFile(PREPPED_FILES, (file) => void openFile(file, 'prepped'));
}

/** "Previously prepped": your Excel sheet, so each list keeps your notes when the run sheet has changed. */
function renderPrepped(compact: boolean): HTMLElement {
  const choose = () => void choosePrepped();
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
    h('p', { class: 'hint small' }, 'Choose your cheat sheet (Excel) to work on it straight away. With the run sheet loaded too, each list is updated from it: your notes stay with each patient, new patients are marked New and cancellations go to the bottom. Read on this computer only.'),
    wb
      ? h('p', {}, '✓ ', h('strong', {}, wb.name), ' ', h('button', { class: 'link', onclick: choose }, 'Change'), ' · ', h('button', { class: 'link', onclick: forgetWorkbook }, 'Remove'))
      : h('button', { onclick: choose }, 'Choose prepped Excel sheet…'),
  );
}

/** Steps 1 and 2: the run sheet (and your prepped sheet), then the list or lists. Compact once loaded. */
function renderSetup(): HTMLElement {
  const fileError = state.fileError ? h('p', { class: 'error', role: 'alert' }, state.fileError) : null;
  if (!state.table) {
    const wb = state.workbook;
    const lists = wb ? sheetLists() : [];
    const sheetStep = lists.length
      ? step(
          2,
          lists.length === 1 ? 'List' : `Lists in ${wb!.name}`,
          h(
            'div',
            { class: 'list-cards' },
            lists.map((b) =>
              h(
                'div',
                { class: b.sheet === state.sheetSelected ? 'list-card no-tick on' : 'list-card no-tick' },
                h(
                  'button',
                  { class: 'list-open', 'aria-pressed': b.sheet === state.sheetSelected ? 'true' : 'false', onclick: () => chooseSheet(b.sheet) },
                  h('span', { class: 'list-title' }, b.sheet),
                  h('span', { class: 'list-sub' }, `${b.matches} patient${b.matches === 1 ? '' : 's'} · from ${b.cell}`),
                ),
              ),
            ),
          ),
          h('p', { class: 'hint small' }, 'From your prepped sheet. Load the latest run sheet above when you have it, to check for patients added, cancelled or moved: your edits here are kept.'),
        )
      : null;
    return h('div', {}, step(
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
            rows: state.workbook ? 3 : 5,
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
    ), sheetStep);
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
    renderSaveMenu(grid),
    grid.sheetOnly
      ? null
      : h(
          'button',
          {
            class: state.updateOpen ? 'on' : '',
            title: 'List changed? With the latest run sheet loaded, paste your prepped rows back in to keep your notes.',
            onclick: () => (state.updateOpen ? closeUpdate() : openUpdate()),
          },
          grid.summary ? 'Update again from my sheet…' : 'Update from my sheet…',
        ),
    renderEmailMenu(grid),
    (() => {
      const items = lensItems(grid);
      if (!items) return null;
      const ready = lensLines(items, state.settings.prompts.lenses).lines.length;
      return h(
        'button',
        {
          disabled: ready === 0,
          title: ready ? 'Copies each ticked cataract patient with their lens, to paste into an email to the theatre coordinator' : 'Pick a lens under a cataract patient first',
          onclick: () => void copyLensList(grid),
        },
        'Copy lens list',
      );
    })(),
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

  const allergyCol = allergyColumn(layout);
  const keyIndex = layout.columns.findIndex((c) => c.source === 'field:patient' || c.source === 'field:name' || c.source === 'field:mrn');
  const rowName = (r: GridRow) => r.patient?.name || (keyIndex >= 0 ? r.cells[keyIndex] : '') || 'this patient';
  const combinedLists = grid.sessions.length > 1;

  const whoOf = (r: GridRow) => {
    const p = r.patient;
    const list = p ? grid.listOf.get(p) : undefined;
    return [
      combinedLists && list ? `${list.or} ${list.ampm}` : '',
      p && list && p.consultantSurname && p.consultantSurname !== list.consultant ? p.consultantSurname : '',
    ].filter(Boolean).join(' · ');
  };
  const badgesOf = (r: GridRow) => {
    const isCancelled = r.status === 'cancelled';
    return [
      allergyPending(r, allergyCol)
        ? h('span', { class: 'badge badge-allergy', title: 'The run sheet says this patient has an allergy. Acknowledge it and type the details.' }, 'Allergy: add details')
        : null,
      r.status === 'new' ? h('span', { class: 'badge badge-new', title: 'Added to the list since your sheet: needs prep' }, 'New') : null,
      r.reinstated ? h('span', { class: 'badge badge-new', title: 'Was cancelled on your sheet and is back on the run sheet' }, 'Back on list') : null,
      r.movedFrom !== undefined ? h('span', { class: 'badge badge-moved', title: 'The order changed: this patient has moved' }, `Moved, was #${r.movedFrom}`) : null,
      isCancelled ? h('span', { class: 'badge badge-cancelled' }, 'Cancelled') : null,
      r.note ? h('div', { class: 'who' }, r.note) : null,
    ];
  };
  /** The ⠿ handle: drag to move (pointer, so fingers work too), or ↑ ↓ keys. */
  const gripOf = (r: GridRow, index: number, items: string) =>
    r.status === 'cancelled'
      ? null
      : h(
          'button',
          {
            class: 'grip',
            title: 'Drag to move (or use the ↑ ↓ keys)',
            'aria-label': `Move ${rowName(r)}: use the up and down arrow keys`,
            'data-key': `${rowKey(r)}:grip`,
            onclick: (e: Event) => e.stopPropagation(),
            onpointerdown: (e: Event) =>
              startDrag(e as PointerEvent, {
                items: [...app.querySelectorAll<HTMLElement>(items)],
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
  const tickOf = (r: GridRow) =>
    h('input', {
      type: 'checkbox',
      checked: r.included,
      title: r.status === 'cancelled' ? 'Untick to leave this cancelled patient off the copy' : 'Untick to leave this patient off the copy and the email',
      'aria-label': `Include ${rowName(r)}`,
      onclick: (e: Event) => e.stopPropagation(),
      onchange: (e: Event) => {
        r.included = (e.target as HTMLInputElement).checked;
        render();
      },
    });
  const classesOf = (r: GridRow) =>
    [
      r.status === 'cancelled' ? 'row-cancelled' : 'active',
      r.included ? '' : 'excluded',
      r.status === 'new' || r.reinstated ? 'row-new' : '',
    ].filter(Boolean);
  /** A cell's colour class. The allergy cell is red until the details are typed (or acknowledged). */
  const classFor = (r: GridRow, col: number, value: string) => {
    if (col === allergyCol && r.patient?.allergyPending && r.status !== 'cancelled') {
      if (allergyPending(r, allergyCol)) return 'allergy-pending';
      if (r.allergyAck && lacksAllergyDetails(value)) return 'allergy-ack';
    }
    return cellClass(layout.columns[col].source, value);
  };
  /** "Acknowledge" under a red allergy box: you've seen it, and will add the details later. */
  const acknowledgeOf = (r: GridRow, col: number) =>
    col === allergyCol && allergyPending(r, col)
      ? h(
          'button',
          {
            class: 'mini ack-btn',
            title: 'I have seen this allergy (type the details in the box above, or acknowledge it now)',
            'data-key': `${rowKey(r)}:ack`,
            onclick: () => {
              r.allergyAck = true;
              // An emptied box would lose the fact that there's an allergy at all.
              if (!r.cells[col].trim() && r.patient) r.cells[col] = r.patient.allergy;
              render();
            },
          },
          '✓ Acknowledge',
        )
      : null;
  /** A banner over the list while ticked patients still have an allergy to check. */
  const allergyBanner = () => {
    const pending = pendingAllergies(active, allergyCol);
    if (!pending.length) return null;
    const many = pending.length > 1;
    return h(
      'div',
      { class: 'allergy-banner', role: 'alert' },
      h('strong', {}, `⚠ ${pending.length} patient${many ? 's have' : ' has'} an allergy to check: `),
      pending.map(rowName).join('; '),
      '. Type the details in the red box, or press Acknowledge.',
    );
  };
  /** A column's text box for a row, with the run sheet's newer value to Use or Keep, if any. */
  const editorOf = (r: GridRow, col: number, hasPrompts: boolean, onEdit: (target: HTMLTextAreaElement) => void) => {
    const c = layout.columns[col];
    const area = h('textarea', {
          class: 'cell',
          rows: 1,
          value: r.cells[col],
          'aria-label': `${c.heading} for ${rowName(r)}`,
          'data-key': `${rowKey(r)}:cell${col}`,
          id: `cell-${rowKey(r)}-${col}`,
          // Leaving an edited cell redraws the prompts, so typing "PXF" by hand lights up its button too.
          onchange: hasPrompts || col === allergyCol ? () => setTimeout(render) : undefined,
          oninput: (e: Event) => {
            const target = e.target as HTMLTextAreaElement;
            r.cells[col] = target.value;
            onEdit(target);
            autoGrow(target);
          },
          // A screenshot pasted into a lens patient's box becomes their IOL report.
          onpaste: (e: Event) => {
            const files = pastedImages(e as ClipboardEvent);
            if (!files.length || !takesReports(r)) return;
            e.preventDefault();
            void addReports(r, files);
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
    return { area, change };
  };

  const renderRow = (r: GridRow, index: number) => {
    const p = r.patient;
    const who = whoOf(r);
    const prompts = r.status === 'cancelled' ? null : renderPrompts(r);
    const main = h(
      'tr',
      {},
      h('td', { class: 'handle' }, gripOf(r, index, 'tbody.patient.active'), tickOf(r)),
      h('td', { class: 'time', title: p?.consultant ?? '' }, h('div', { class: 'time-value' }, p?.time ?? ''), who ? h('div', { class: 'who' }, who) : null, ...badgesOf(r)),
      layout.columns.map((c, col) => {
        const { area, change } = editorOf(r, col, Boolean(prompts), (target) => (target.parentElement!.className = classFor(r, col, target.value)));
        return h('td', { class: classFor(r, col, r.cells[col]), style: `min-width:${columnWidth(c.source)}px` }, area, change, acknowledgeOf(r, col));
      }),
    );
    return h(
      'tbody',
      { class: ['patient', ...classesOf(r)].join(' ') },
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

  const phone = isPhone();
  const colOf = (source: string) => layout.columns.findIndex((c) => c.source === source);
  const eyeCol = colOf('field:eye');
  const procCol = Math.max(colOf('field:procedure'), colOf('field:procedureRaw'));
  const alertCol = colOf('field:alerts');

  /** Phone: one card per patient. Closed, it's a summary; tap it to edit each column in full width. */
  const renderCard = (r: GridRow, index: number) => {
    const key = rowKey(r);
    const open = state.openCards.has(key);
    const p = r.patient;
    const eye = eyeCol >= 0 ? r.cells[eyeCol].trim().toUpperCase() : '';
    const title = (keyIndex >= 0 ? r.cells[keyIndex] : '') || p?.patient || rowName(r);
    const procedure = procCol >= 0 ? r.cells[procCol] : p?.procedure ?? '';
    const alerts = alertCol >= 0 ? r.cells[alertCol].trim() : '';
    const sub = [procedure, p?.time, whoOf(r)].filter(Boolean).join(' · ');
    const toggle = () => {
      if (open) state.openCards.delete(key);
      else state.openCards.add(key);
      render();
    };
    const prompts = r.status === 'cancelled' || !open ? null : renderPrompts(r);
    const promptCols = promptColumns(layout, state.settings.prompts);
    const eyeButtons =
      eyeCol >= 0
        ? h(
            'div',
            { class: 'eye-choice', role: 'group', 'aria-label': `Eye for ${rowName(r)}` },
            [['RIGHT', 'Right'], ['LEFT', 'Left'], ['BOTH', 'Both']].map(([value, label]) =>
              h(
                'button',
                {
                  class: `eye-btn eye-btn-${value.toLowerCase()}${eye === value ? ' on' : ''}`,
                  'aria-pressed': eye === value ? 'true' : 'false',
                  'data-key': `${key}:eye:${value}`,
                  onclick: () => ((r.cells[eyeCol] = eye === value ? '' : value), render()),
                },
                label,
              ),
            ),
          )
        : null;
    const body = open
      ? h(
          'div',
          { class: 'pcard-body' },
          layout.columns.map((c, col) => {
            if (col === eyeCol) return h('div', { class: 'pfield' }, h('span', { class: 'pfield-label' }, c.heading || 'Eye'), eyeButtons, r.changes[col] !== undefined ? editorOf(r, col, false, () => undefined).change : null);
            const { area, change } = editorOf(r, col, Boolean(prompts), () => undefined);
            // The prompts sit under the box they fill in: findings under Diagnosis, lens and target under Implants.
            const near = prompts && col === promptCols.findings ? renderPrompts(r, 'findings') : prompts && col === promptCols.lens ? renderPrompts(r, 'lens') : null;
            return h('div', { class: `pfield ${classFor(r, col, r.cells[col])}` }, h('label', { class: 'pfield-label', for: `cell-${rowKey(r)}-${col}` }, c.heading || `Column ${col + 1}`), area, change, acknowledgeOf(r, col), near);
          }),
          // With no Implants column, the IOL report still has a place on the card.
          promptCols.lens === -1 && prompts ? renderReports(r) : null,
          r.status === 'cancelled'
            ? null
            : h(
                'div',
                { class: 'pcard-move' },
                h('button', { disabled: index <= 0, 'data-key': `${key}:up`, onclick: () => moveRow(grid, index, index - 1) }, '↑ Move up'),
                h('button', { disabled: index >= active.length - 1, 'data-key': `${key}:down`, onclick: () => moveRow(grid, index, index + 1) }, '↓ Move down'),
                h('button', { class: 'ghost', onclick: toggle }, 'Done'),
              ),
        )
      : null;
    return h(
      'div',
      { class: ['pcard', ...classesOf(r), open ? 'open' : ''].filter(Boolean).join(' ') },
      h(
        'div',
        {
          class: 'pcard-head',
          role: 'button',
          tabindex: 0,
          'aria-expanded': open ? 'true' : 'false',
          onclick: toggle,
          onkeydown: (e: Event) => {
            if ((e as KeyboardEvent).key === 'Enter' || (e as KeyboardEvent).key === ' ') {
              e.preventDefault();
              toggle();
            }
          },
        },
        gripOf(r, index, '.pcard.active'),
        h('span', { class: `eye-pill ${eye ? `eye-${eye.toLowerCase()}` : 'eye-missing'}` }, eye === 'RIGHT' ? 'R' : eye === 'LEFT' ? 'L' : eye === 'BOTH' ? 'B' : '?'),
        h(
          'div',
          { class: 'pcard-title' },
          h('div', { class: 'pcard-name' }, title.split('\n')[0]),
          sub ? h('div', { class: 'pcard-sub' }, sub.split('\n')[0]) : null,
          h('div', { class: 'pcard-badges' }, ...badgesOf(r)),
        ),
        tickOf(r),
      ),
      alerts && !open ? h('div', { class: 'pcard-alert' }, alerts.split('\n').join(' · ')) : null,
      body,
    );
  };

  const cards = h(
    'div',
    { class: 'pcards' },
    active.map(renderCard),
    cancelled.length ? h('div', { class: 'pcards-cancelled' }, `Cancelled: not on the latest run sheet (${cancelled.length}). Copied below the list under a CANCELLED row; untick any you don’t want to keep.`) : null,
    cancelled.map((r) => renderCard(r, -1)),
  );
  const allOpen = grid.rows.length > 0 && grid.rows.every((r) => state.openCards.has(rowKey(r)));
  const openAll = h(
    'div',
    { class: 'pcards-tools' },
    h('label', {}, h('input', { type: 'checkbox', checked: active.length > 0 && active.every((r) => r.included), onchange: (e: Event) => {
      const on = (e.target as HTMLInputElement).checked;
      for (const r of active) r.included = on;
      render();
    } }), ` ${count} of ${active.length} ticked`),
    h('button', { class: 'link', onclick: () => {
      if (allOpen) state.openCards.clear();
      else for (const r of grid.rows) state.openCards.add(rowKey(r));
      render();
    } }, allOpen ? 'Close all' : 'Open all'),
  );

  return h(
    'section',
    { class: phone ? 'step prep phone' : 'step prep' },
    h(
      'div',
      { class: 'prep-head' },
      grid.sheetOnly
        ? h('h2', {}, h('span', { class: 'num' }, '3'), combined.slotName, h('span', { class: 'muted heading-note' }, ` from ${grid.fromWorkbook?.name ?? 'your sheet'}`))
        : h('h2', {}, h('span', { class: 'num' }, '3'), listName(combined)),
      grid.sheetOnly ? null : factChips,
    ),
    phone ? renderPhoneBar(grid, count, cancelledCount) : toolbar,
    renderSaveReport(),
    renderNewSheet(grid),
    renderNewFile(grid),
    grid.sheetOnly
      ? h('p', { class: 'source-bar muted' }, 'Working from your prepped sheet, without a run sheet: times, New, Moved and the Medical Records email need the run sheet.')
      : renderSource(grid),
    renderUpdate(grid),
    allergyBanner(),
    phone ? [openAll, cards] : h('div', { class: 'table-wrap' }, table),
  );
}

/** Phone: Copy, Save and ⋯ (Update, Email) in a bar along the bottom, within reach of a thumb. */
function renderPhoneBar(grid: Grid, count: number, cancelledCount: number): HTMLElement {
  const ticked = tickedPatients(grid).length;
  const hasRecipients = medicalRecordsEmail(grid).to.length > 0;
  const item = (label: string, detail: string, onclick: () => void, disabled = false) =>
    h('button', { class: 'menu-item', role: 'menuitem', disabled, onclick }, h('span', {}, label), h('span', { class: 'menu-detail' }, detail));
  const run = (action: () => void) => () => {
    state.moreMenuOpen = false;
    action();
  };
  return h(
    'div',
    { class: 'phone-bar' },
    h('button', { class: 'primary', disabled: count + cancelledCount === 0, onclick: () => copyRows(grid) }, `Copy ${count}${cancelledCount ? ` + ${cancelledCount}` : ''}`),
    renderSaveMenu(grid),
    h(
      'div',
      { class: 'menu-wrap' },
      h(
        'button',
        {
          class: state.moreMenuOpen ? 'on' : '',
          'aria-haspopup': 'menu',
          'aria-expanded': state.moreMenuOpen ? 'true' : 'false',
          'aria-label': 'More: update from your sheet, email Medical Records',
          onclick: () => ((state.moreMenuOpen = !state.moreMenuOpen), (state.saveMenuOpen = false), (state.emailMenuOpen = false), render()),
        },
        '⋯',
      ),
      state.moreMenuOpen
        ? h(
            'div',
            { class: 'menu', role: 'menu' },
            grid.sheetOnly ? null : item(grid.summary ? 'Update again from my sheet…' : 'Update from my sheet…', 'List changed? Keep your notes', run(() => openUpdate())),
            h('p', { class: 'menu-note' }, `Email Medical Records for ${emailSummary(grid)}:`),
            item('Outlook on the web', 'Opens a draft in a new tab', run(() => void emailMedicalRecords(grid)), !ticked || !hasRecipients),
            item('Email app', 'The email app on this phone', run(() => void emailInApp(grid)), !ticked || !hasRecipients),
            canShare ? item('Share…', 'Text only: add the address yourself', run(() => void shareEmail(grid)), !ticked) : null,
            item('Copy email text', 'To paste into any email', run(() => void copyEmailText(grid)), !ticked),
            hasRecipients ? null : h('p', { class: 'menu-note' }, 'Add the Medical Records address in Settings to open a draft.'),
            lensItems(grid) ? item('Copy lens list', 'Patients and lenses, for the theatre coordinator', run(() => void copyLensList(grid))) : null,
          )
        : null,
    ),
  );
}

/**
 * Save to file.xlsx: saves straight into this list's sheet (or, for a list not in the file, opens
 * "Save as a new sheet"); ▾ has both.
 */
function renderSaveMenu(grid: Grid): HTMLElement | null {
  const wb = state.workbook;
  if (!wb?.book) {
    return h(
      'button',
      {
        class: state.newFile ? 'on' : '',
        disabled: Boolean(state.saveReport?.busy || state.saveReport?.actions),
        title: 'Saves this list in a new Excel file made from your template',
        onclick: () => openNewFile(grid),
      },
      h('span', { class: 'label-long' }, 'Save to a new Excel file…'),
      h('span', { class: 'label-short' }, 'Save…'),
    );
  }
  const linked = grid.fromWorkbook === wb && Boolean(grid.fromSheet);
  const others = workbookEdits(wb).length - 1;
  const how = wb.handle ? 'Saves over the file (your browser may ask first).' : 'Asks where to save: choose the file you opened to replace it.';
  const item = (label: string, detail: string, onclick: () => void) =>
    h('button', { class: 'menu-item', role: 'menuitem', onclick }, h('span', {}, label), h('span', { class: 'menu-detail' }, detail));
  const saveHere = () => {
    state.saveMenuOpen = false;
    void saveToWorkbook();
  };
  return h(
    'div',
    { class: 'menu-wrap split' },
    h(
      'button',
      {
        class: 'split-main',
        disabled: Boolean(state.saveReport?.busy || state.saveReport?.actions),
        title: linked ? `Saves the cells you changed into sheet “${grid.fromSheet}” of ${wb.name}. ${how}` : `This list isn’t in ${wb.name} yet: save it as a new sheet.`,
        onclick: () => (linked ? saveHere() : openNewSheet(grid)),
      },
      h('span', { class: 'label-long' }, linked ? `Save to ${wb.name}` : `Save as a new sheet in ${wb.name}…`),
      h('span', { class: 'label-short' }, linked ? 'Save' : 'Save…'),
    ),
    h(
      'button',
      {
        class: state.saveMenuOpen || state.newSheet || state.newFile ? 'split-more on' : 'split-more',
        title: 'More ways to save',
        'aria-label': 'More ways to save',
        'aria-haspopup': 'menu',
        'aria-expanded': state.saveMenuOpen ? 'true' : 'false',
        onclick: () => ((state.saveMenuOpen = !state.saveMenuOpen), (state.emailMenuOpen = false), render()),
      },
      '▾',
    ),
    state.saveMenuOpen
      ? h(
          'div',
          { class: 'menu', role: 'menu' },
          linked
            ? item(
                `Save to sheet “${grid.fromSheet}”`,
                `Only the cells you changed${others > 0 ? `, here and in ${others} other list${others === 1 ? '' : 's'} from this file` : ''}. Formatting and other sheets stay as they are.`,
                saveHere,
              )
            : null,
          item('Save as a new sheet…', 'Adds a sheet with this list, copying the headings and formatting of one of your sheets.', () => openNewSheet(grid)),
          item('Save to a new Excel file…', 'A separate file with just this list, made from your template or one of these sheets.', () => openNewFile(grid)),
          h('p', { class: 'menu-note' }, `${how} Close the file in Excel first.`),
        )
      : null,
  );
}

/** How the last save went: where the file is, or why it couldn't save over it. */
function renderSaveReport(): HTMLElement | null {
  const report = state.saveReport;
  if (!report) return null;
  const actions = report.actions ?? [];
  return h(
    'div',
    { class: `save-report${report.ok ? '' : ' warn'}${report.busy ? ' busy' : ''}`, role: report.ok ? 'status' : 'alert' },
    h('span', {}, report.busy ? h('span', { class: 'spinner', 'aria-hidden': 'true' }) : report.ok ? '✓ ' : actions.length ? '' : '⚠ ', report.text),
    actions.length
      ? h('span', { class: 'report-actions' }, actions.map((a) => h('button', { class: a.primary ? 'primary' : '', onclick: a.run }, a.label)))
      : report.busy
        ? null
        : h('button', { class: 'link', onclick: () => ((state.saveReport = null), render()) }, 'Dismiss'),
  );
}

/** The "Save to a new Excel file" form: its name, and the template (or sheet) to make it from. */
function renderNewFile(grid: Grid): HTMLElement | null {
  const form = state.newFile;
  if (!form) return null;
  const sources = newFileSources();
  const busy = Boolean(state.saveReport?.busy || state.saveReport?.actions);
  const source = h(
    'select',
    { 'aria-label': 'Make it from', onchange: (e: Event) => ((form.source = (e.target as HTMLSelectElement).value), render()) },
    sources.map((s) => h('option', { value: s.key }, s.label)),
  );
  source.value = form.source;
  return h(
    'div',
    { class: 'update-open new-sheet new-file' },
    h(
      'p',
      {},
      h('strong', {}, 'Save to a new Excel file. '),
      'Makes a separate file with just this list, laid out like your template: its headings, column widths, colours, dropdowns and print settings, with the list in place of its patients and the heading filled in. Sheets without patients (like instructions) come along too.',
    ),
    h(
      'div',
      { class: 'inline-form' },
      h(
        'label',
        {},
        'File name ',
        h('input', {
          type: 'text',
          value: form.name,
          size: 34,
          oninput: (e: Event) => (form.name = (e.target as HTMLInputElement).value),
          onkeydown: (e: Event) => {
            if ((e as KeyboardEvent).key === 'Enter') void saveToNewFile(grid);
          },
        }),
      ),
      sources.length ? h('label', {}, 'Make it from ', source) : null,
      h('button', { class: sources.length ? 'link' : '', onclick: chooseTemplate }, sources.length ? 'Choose a template file…' : 'Choose your template file…'),
      state.template ? h('button', { class: 'link', title: 'Stop keeping the template on this computer', onclick: forgetTemplate }, 'Forget template') : null,
    ),
    h(
      'div',
      { class: 'inline-form' },
      h('button', { class: 'primary', disabled: busy || !sources.length, onclick: () => void saveToNewFile(grid) }, state.saveReport?.busy ? 'Saving…' : 'Save new file'),
      h('button', { class: 'ghost', onclick: () => ((state.newFile = null), render()) }, 'Cancel'),
      h('span', { class: 'hint small' }, state.template ? `Template: ${state.template.name}, kept on this computer only.` : 'A template you choose is kept on this computer only (never if it has patients in it).'),
    ),
    form.error ? h('p', { class: 'error', role: 'alert' }, form.error) : null,
  );
}

/** The "Save as a new sheet" form: its name, and which sheet's layout to copy. */
function renderNewSheet(grid: Grid): HTMLElement | null {
  const form = state.newSheet;
  const book = state.workbook?.book;
  if (!form || !book) return null;
  const copyOf = h(
    'select',
    { 'aria-label': 'Copy the layout of', onchange: (e: Event) => (form.copyOf = (e.target as HTMLSelectElement).value) },
    book.sheets.map((s) => h('option', { value: s.name }, s.name)),
  );
  copyOf.value = form.copyOf;
  return h(
    'div',
    { class: 'update-open new-sheet' },
    h('p', {}, h('strong', {}, 'Save as a new sheet. '), `Adds a sheet to ${state.workbook!.name} with this list in place of the copied sheet’s patients. Its headings, column widths, formatting and formulas are copied too.`),
    h(
      'div',
      { class: 'inline-form' },
      h(
        'label',
        {},
        'Name ',
        h('input', {
          type: 'text',
          value: form.name,
          maxlength: 31,
          oninput: (e: Event) => (form.name = (e.target as HTMLInputElement).value),
          onkeydown: (e: Event) => {
            if ((e as KeyboardEvent).key === 'Enter') void saveAsNewSheet(grid);
          },
        }),
      ),
      h('label', {}, 'Copy the layout of ', copyOf),
      h('button', { class: 'primary', disabled: Boolean(state.saveReport?.busy || state.saveReport?.actions), onclick: () => void saveAsNewSheet(grid) }, state.saveReport?.busy ? 'Saving…' : 'Save new sheet'),
      h('button', { class: 'ghost', onclick: () => ((state.newSheet = null), render()) }, 'Cancel'),
    ),
    form.error ? h('p', { class: 'error', role: 'alert' }, form.error) : null,
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
        title: grid.sheetOnly ? 'Load the run sheet to email Medical Records' : '',
        'aria-haspopup': 'menu',
        'aria-expanded': state.emailMenuOpen ? 'true' : 'false',
        onclick: () => ((state.emailMenuOpen = !state.emailMenuOpen), (state.saveMenuOpen = false), render()),
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
      h('button', { class: 'link', onclick: () => void choosePrepped() }, 'choose your Excel file'),
      ' instead: the rows for this list are found in it for you.',
    ),
    state.updateError ? h('p', { class: 'error', role: 'alert' }, state.updateError) : null,
    h('button', { class: 'ghost', onclick: closeUpdate }, 'Cancel'),
  );
}

/** Identifies a row: its booking, or (for rows only on your sheet) its place there. */
function rowKey(r: GridRow): string {
  return r.patient?.id ?? r.id ?? '';
}

/** Cataract or lens surgery: from the booking, or from the Procedure cell for rows only on your sheet. */
function isLensRow(r: GridRow): boolean {
  if (r.status === 'cancelled') return false;
  if (r.patient) return isLensSurgery(r.patient);
  const col = state.settings.layout.columns.findIndex((c) => c.source === 'field:procedure' || c.source === 'field:procedureRaw');
  return col >= 0 && /\bCE\b|\bIOL\b|cataract|phaco|lens/i.test(r.cells[col] ?? '');
}

/** One-click findings, lens and target under a lens patient. Each click edits the cell itself. */
/** The prompts under a patient; `only` gives just the findings (Diagnosis) or the lens and target (Implants). */
function renderPrompts(r: GridRow, only?: 'findings' | 'lens'): HTMLElement | null {
  const { prompts, layout } = state.settings;
  const reports = only !== 'findings' ? renderReports(r) : null;
  const cols = promptColumns(layout, prompts);
  if (!prompts.enabled || !isLensRow(r) || (cols.findings === -1 && cols.lens === -1)) return reports ? h('div', { class: 'prompts' }, reports) : null;
  const key = rowKey(r);
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
  if (cols.findings !== -1 && only !== 'lens') {
    const col = cols.findings;
    const cell = r.cells[col];
    const order = readEyeOrder(cell);
    const eyeOrder = h(
      'span',
      { class: 'segmented' },
      h('span', { class: 'prompt-label' }, 'Eye'),
      chip('1st', order === 'first', () => set(col, writeEyeOrder(cell, order === 'first' ? null : 'first')), '1st eye: goes at the top of the box'),
      chip('2nd', order === 'second', () => set(col, writeEyeOrder(cell, order === 'second' ? null : 'second')), '2nd eye: goes at the top of the box'),
    );
    parts.push(
      h(
        'span',
        { class: 'prompt-set' },
        eyeOrder,
        findingGroups(prompts.findings).map((g) => {
          const chips = g.options.map((o) =>
            chip(o.short, hasToken(cell, o.token), () => set(col, toggleFinding(cell, g, o.token, prompts.findings)), o.token),
          );
          return g.options.length > 1 ? h('span', { class: 'segmented' }, g.label ? h('span', { class: 'prompt-label' }, g.label) : null, chips) : chips[0];
        }),
      ),
    );
  }
  if (cols.lens !== -1 && only !== 'findings') {
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
  if (reports) parts.push(reports);
  return parts.length ? h('div', { class: 'prompts' }, parts) : null;
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


// ---- IOL reports -----------------------------------------------------------------

/** A short name for a row: MRN and name, or the first line of its patient cell. */
function rowLabel(r: GridRow): string {
  if (r.patient) return [r.patient.mrn, r.patient.name].filter(Boolean).join(' ');
  const { layout } = state.settings;
  const col = layout.columns.findIndex((c) => c.source === 'field:patient' || c.source === 'field:name' || c.source === 'field:mrn');
  return (col >= 0 ? r.cells[col] : '').split('\n')[0].trim() || 'this patient';
}

/** Cataract patients (and anyone who already has one) can have IOL report screenshots. */
function takesReports(r: GridRow): boolean {
  return Boolean(r.reports?.length) || (r.status !== 'cancelled' && isLensRow(r));
}

const imageUrls = new WeakMap<ReportImage, string>();

function imageUrl(image: ReportImage): string {
  let url = imageUrls.get(image);
  if (!url) {
    url = URL.createObjectURL(new Blob([image.bytes as BlobPart], { type: image.type }));
    imageUrls.set(image, url);
  }
  return url;
}

/** Pictures on the clipboard, unless it also has text (as when copying cells from Excel). */
function pastedImages(e: ClipboardEvent): File[] {
  const data = e.clipboardData;
  if (!data || data.getData('text/plain').trim()) return [];
  return [...data.files].filter((f) => f.type.startsWith('image/'));
}

/** A screenshot as PNG, JPEG or GIF; anything else the browser can show (BMP, WebP…) becomes a PNG. */
async function imageFromFile(file: Blob): Promise<ReportImage | null> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const direct = reportImage(bytes, file.type);
  if (direct) return direct;
  try {
    const bitmap = await createImageBitmap(file);
    const canvas = h('canvas', { width: bitmap.width, height: bitmap.height });
    canvas.getContext('2d')!.drawImage(bitmap, 0, 0);
    const png = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
    return png ? reportImage(new Uint8Array(await png.arrayBuffer())) : null;
  } catch {
    return null;
  }
}

/** Where a new screenshot is kept: in your Excel file once you save. */
function reportSaveHint(): string {
  const grid = state.grid;
  const wb = grid?.fromWorkbook && grid.fromSheet ? grid.fromWorkbook : null;
  return wb ? `Save to keep it in ${wb.name}.` : 'Save the list to an Excel file to keep it.';
}

async function addReports(r: GridRow, files: Blob[]): Promise<void> {
  const images = (await Promise.all(files.map(imageFromFile))).filter((i): i is ReportImage => i !== null);
  if (!images.length) {
    say('That isn’t a picture the app can read. Paste or choose a screenshot (PNG or JPEG).');
    return;
  }
  r.reports = [...(r.reports ?? []), ...images];
  const what = images.length === 1 ? 'an IOL report screenshot' : `${images.length} IOL report screenshots`;
  say(`Added ${what} for ${rowLabel(r)}. ${reportSaveHint()}`);
}

async function removeReport(r: GridRow, index: number): Promise<void> {
  const go = await askInPage(`Remove this IOL report screenshot for ${rowLabel(r)}?`, [
    { label: 'Remove', value: true, primary: true },
    { label: 'Keep it', value: false },
  ]);
  state.saveReport = null;
  if (go && r.reports) {
    r.reports = r.reports.filter((_, i) => i !== index);
    if (state.viewer && state.viewer.key === rowKey(r)) state.viewer = r.reports.length ? { ...state.viewer, index: Math.min(index, r.reports.length - 1) } : null;
    say(`Removed. ${reportSaveHint().replace('keep it in', 'take it out of')}`);
    return;
  }
  render();
}

function openViewer(r: GridRow, index: number): void {
  state.viewer = { key: rowKey(r), index, actualSize: false };
  render();
}

/** Thumbnails of the patient's IOL report screenshots, with a way to add more. */
function renderReports(r: GridRow): HTMLElement | null {
  if (!takesReports(r)) return null;
  const key = rowKey(r);
  const reports = r.reports ?? [];
  const input = h('input', {
    type: 'file',
    accept: 'image/*',
    multiple: true,
    hidden: true,
    onchange: (e: Event) => {
      const files = [...((e.target as HTMLInputElement).files ?? [])];
      if (files.length) void addReports(r, files);
    },
  });
  return h(
    'span',
    {
      class: reports.length ? 'reports has-reports' : 'reports',
      'data-report-row': key,
      title: 'Paste a screenshot of the IOL report here (click, then Ctrl+V), or drop one on it',
    },
    h('span', { class: 'prompt-label' }, 'IOL report'),
    reports.map((image, i) =>
      h(
        'span',
        { class: 'report-thumb' },
        h(
          'button',
          { class: 'report-open', title: 'Open full screen, e.g. to check at timeout', 'data-key': `${key}:report:${i}`, onclick: () => openViewer(r, i) },
          h('img', { src: imageUrl(image), alt: `IOL report ${i + 1} for ${rowLabel(r)}` }),
        ),
        h('button', { class: 'report-remove', title: 'Remove this screenshot', 'aria-label': `Remove IOL report ${i + 1} for ${rowLabel(r)}`, onclick: () => void removeReport(r, i) }, '×'),
      ),
    ),
    h(
      'button',
      { class: 'chip', 'data-key': `${key}:report:add`, title: 'Choose a screenshot of the IOL report, or click here and paste one with Ctrl+V', onclick: () => input.click() },
      reports.length ? '+ Add' : '+ Add screenshot',
    ),
    input,
    reports.length ? renderReportStatus(r) : h('span', { class: 'hint small report-hint' }, 'or paste one (Ctrl+V) into the Implants box'),
  );
}

/** The lens picked for a row, as at timeout: eye, lens and power, target. */
function lensFacts(r: GridRow): { eye: string; lens: string; target: string } {
  const { layout, prompts } = state.settings;
  const eyeCol = layout.columns.findIndex((c) => c.source === 'field:eye');
  const eye = (eyeCol >= 0 ? r.cells[eyeCol] : r.patient?.eye ?? '').trim().toUpperCase();
  const lensCol = promptColumns(layout, prompts).lens;
  const line = lensCol >= 0 ? lensLines([{ who: '', eye, cell: r.cells[lensCol] ?? '' }], prompts.lenses).lines[0] : undefined;
  return { eye, lens: line ? [line.model, line.power].filter(Boolean).join(' ') : '', target: line?.target ?? '' };
}

/** A patient's IOL report full screen, with the lens you picked above it, to check at timeout. */
function renderViewer(): HTMLElement | null {
  const v = state.viewer;
  if (!v) return null;
  const r = state.grid?.rows.find((row) => rowKey(row) === v.key);
  const reports = r?.reports ?? [];
  const image = reports[v.index];
  if (!r || !image) {
    state.viewer = null;
    return null;
  }
  const close = () => ((state.viewer = null), render());
  const go = (index: number) => ((state.viewer = { ...v, index, actualSize: false }), render());
  const { eye, lens, target } = lensFacts(r);
  const eyeLabel = eye === 'RIGHT' ? 'Right eye' : eye === 'LEFT' ? 'Left eye' : eye === 'BOTH' ? 'Both eyes' : 'Eye not set';
  return h(
    'div',
    { class: 'viewer', role: 'dialog', 'aria-modal': 'true', 'aria-label': `IOL report for ${rowLabel(r)}` },
    h(
      'div',
      { class: 'viewer-head' },
      h(
        'div',
        { class: 'viewer-facts' },
        h('div', { class: 'viewer-name' }, rowLabel(r)),
        h(
          'div',
          { class: 'viewer-lens' },
          h('span', { class: `viewer-eye viewer-eye-${eye ? eye.toLowerCase() : 'missing'}` }, eyeLabel),
          h('span', { class: lens ? 'viewer-iol' : 'viewer-iol missing' }, lens || 'No lens picked'),
          target ? h('span', { class: 'viewer-target' }, `Target ${target}`) : null,
        ),
        renderViewerCheck(r),
      ),
      h(
        'div',
        { class: 'viewer-actions' },
        reports.length > 1
          ? [
              h('button', { 'data-key': 'viewer:prev', disabled: v.index === 0, 'aria-label': 'Previous screenshot', onclick: () => go(v.index - 1) }, '‹'),
              h('span', { class: 'viewer-count' }, `${v.index + 1} of ${reports.length}`),
              h('button', { 'data-key': 'viewer:next', disabled: v.index >= reports.length - 1, 'aria-label': 'Next screenshot', onclick: () => go(v.index + 1) }, '›'),
            ]
          : null,
        h(
          'button',
          { 'data-key': 'viewer:size', onclick: () => ((state.viewer = { ...v, actualSize: !v.actualSize }), render()) },
          v.actualSize ? 'Fit to width' : 'Actual size',
        ),
        h('button', { class: 'primary', 'data-key': 'viewer:close', onclick: close }, 'Close'),
      ),
    ),
    h(
      'div',
      { class: v.actualSize ? 'viewer-body actual' : 'viewer-body' },
      h('img', { src: imageUrl(image), alt: `IOL report for ${rowLabel(r)}`, onclick: () => ((state.viewer = { ...v, actualSize: !v.actualSize }), render()) }),
    ),
  );
}


// ---- Reading the IOL report (OCR) -------------------------------------------------

type ReadState = ReportReading | 'reading' | 'unreadable' | 'failed';
const readings = new WeakMap<ReportImage, ReadState>();
let readQueue: Promise<void> = Promise.resolve();

/** Reads a screenshot's report in the background (one at a time), then redraws. */
function readingOf(image: ReportImage): ReadState {
  const known = readings.get(image);
  if (known) return known;
  readings.set(image, 'reading');
  readQueue = readQueue.then(async () => {
    try {
      const words = await readWords(new Blob([image.bytes as BlobPart], { type: image.type }));
      readings.set(image, readReport(words) ?? 'unreadable');
    } catch (e) {
      console.error('Couldn’t read the IOL report', e);
      readings.set(image, 'failed');
    }
    render();
  });
  return 'reading';
}

/** The lens picked for a row, as the report check needs it. */
function pickedLens(r: GridRow): PickedLens {
  const { layout, prompts } = state.settings;
  const eyeCol = layout.columns.findIndex((c) => c.source === 'field:eye');
  const eye = (eyeCol >= 0 ? r.cells[eyeCol] : r.patient?.eye ?? '').trim().toUpperCase();
  const lensCol = promptColumns(layout, prompts).lens;
  const cell = lensCol >= 0 ? (r.cells[lensCol] ?? '') : '';
  const pick = readLens(cell, prompts.lenses);
  const number = (text: string) => {
    const m = text.match(/[+-]?\d+(?:\.\d+)?/);
    return m ? Number(m[0]) : null;
  };
  let code = '';
  let power: number | null = null;
  let family: PickedLens['family'];
  if (pick) {
    code = lensCode(pick);
    power = number(pick.power);
    if (pick.lens.includes('x')) family = { base: pick.lens.split('x')[0], number: pick.extra };
  } else {
    // A lens typed in the box, e.g. "AR40 +19.5D".
    const line = cell.split('\n').find((l) => l.trim() && !/^target\b/i.test(l.trim())) ?? '';
    const m = line.match(/\b([A-Z]{2,}[A-Z0-9]*\d[A-Z0-9]*)\b\s*([+-]?\d{1,2}(?:\.\d{1,2})?)?/i);
    if (m) {
      code = m[1].toUpperCase();
      power = m[2] ? Number(m[2]) : null;
    }
  }
  const target = readTarget(cell);
  const label = rowLabel(r);
  const mrn = r.patient?.mrn ?? label.split(/\s+/)[0] ?? '';
  const surname = (r.patient?.name ?? label.replace(/^\S+\s+/, '')).split(',')[0].trim();
  return { eye, code, family, power, target: target === null ? null : target === 'plano' ? 0 : number(target), mrn, surname };
}

interface ReportStatus {
  state: 'reading' | 'unreadable' | 'checked';
  check?: LensCheck;
}

/** What the screenshots say about the lens picked, once they've been read. */
function reportStatus(r: GridRow): ReportStatus | null {
  const reports = r.reports ?? [];
  if (!reports.length) return null;
  const states = reports.map(readingOf);
  if (states.includes('reading')) return { state: 'reading' };
  const read = states.filter((s): s is ReportReading => typeof s === 'object');
  const combined = combineReadings(read);
  if (!combined) return { state: 'unreadable' };
  return { state: 'checked', check: checkLens(combined, pickedLens(r)) };
}

const CHECK_ICON: Record<string, string> = { ok: '✓', warn: '⚠', error: '⛔', none: '' };

/** One line under the thumbnails: what the report says for this lens, or what doesn't match. */
function renderReportStatus(r: GridRow): HTMLElement | null {
  const status = reportStatus(r);
  if (!status) return null;
  if (status.state === 'reading') return h('span', { class: 'report-check reading' }, 'Reading the report…');
  if (status.state === 'unreadable' || !status.check) return h('span', { class: 'report-check none' }, 'Couldn’t read it as an IOL calculation: check it by eye.');
  const { level, facts, flags, notes } = status.check;
  const first = flags.find((f) => f.level === 'error') ?? flags[0];
  // Errors stand alone; a warning comes after what the report says.
  const summary = level === 'error' ? first!.text : [...facts, ...(first ? [first.text] : [])].join(' · ');
  const all = [...facts, ...flags.map((f) => f.text), ...notes].join('\n');
  return h('span', { class: `report-check ${level}`, title: all }, `${CHECK_ICON[level]} ${summary}`.trim());
}

/** In the full-screen view: the check in full, and the report's own warnings. */
function renderViewerCheck(r: GridRow): HTMLElement | null {
  const status = reportStatus(r);
  if (!status) return null;
  if (status.state === 'reading') return h('div', { class: 'viewer-check' }, h('div', { class: 'report-check reading' }, 'Reading the report…'));
  if (status.state === 'unreadable' || !status.check) return h('div', { class: 'viewer-check' }, h('div', { class: 'report-check none' }, 'Couldn’t read it as an IOL calculation: check it by eye.'));
  const { facts, flags, notes } = status.check;
  return h(
    'div',
    { class: 'viewer-check' },
    facts.map((f) => h('div', { class: 'report-check ok' }, `✓ Report: ${f}`)),
    flags.map((f) => h('div', { class: `report-check ${f.level}` }, `${CHECK_ICON[f.level]} ${f.text}`)),
    notes.map((n) => h('div', { class: 'report-check note' }, `IOLMaster: ${n}`)),
    h('div', { class: 'report-check-hint' }, 'Read from the screenshot by the app. Always check against the report itself.'),
  );
}

/** The row a report box belongs to, for pasting or dropping a screenshot on it. */
function reportRowAt(target: EventTarget | null): GridRow | null {
  const key = (target as HTMLElement | null)?.closest?.<HTMLElement>('[data-report-row]')?.dataset.reportRow;
  return key ? (state.grid?.rows.find((r) => rowKey(r) === key) ?? null) : null;
}

// Drop a run sheet or settings file anywhere on the page.
const overlay = h('div', { class: 'drop-overlay', hidden: true }, 'Drop a run sheet (.csv) or a settings file (.json)');
document.body.append(overlay);
let dragDepth = 0;
const hasFiles = (e: DragEvent) => Boolean(e.dataTransfer?.types.includes('Files')) && !isImageDrag(e);
/** Pictures dragged in are for a patient's IOL report box, not a run sheet. */
const isImageDrag = (e: DragEvent) => {
  const items = [...(e.dataTransfer?.items ?? [])];
  return items.length > 0 && items.every((i) => i.kind === 'file' && i.type.startsWith('image/'));
};
document.addEventListener('dragover', (e) => {
  if (!isImageDrag(e)) return;
  e.preventDefault();
  if (e.dataTransfer) e.dataTransfer.dropEffect = reportRowAt(e.target) ? 'copy' : 'none';
});
document.addEventListener('drop', (e) => {
  if (!isImageDrag(e)) return;
  e.preventDefault();
  const r = reportRowAt(e.target);
  const files = [...(e.dataTransfer?.files ?? [])];
  if (r) void addReports(r, files);
  else say('Drop the screenshot on a patient’s IOL report box (under their lens).');
});
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
  if (!file) return;
  // Edge/Chrome can hand over the file itself, so changes can be saved back into it.
  const item = e.dataTransfer?.items[0] as (DataTransferItem & { getAsFileSystemHandle?: () => Promise<FileSystemHandle | null> }) | undefined;
  const handle = item?.getAsFileSystemHandle?.();
  void (async () => {
    const h = handle ? await handle.catch(() => null) : null;
    await openFile(file, 'any', h && h.kind === 'file' ? (h as FileSystemFileHandle) : null);
  })();
});


// Close the email menu on a click elsewhere or Escape.
document.addEventListener('pointerdown', (e) => {
  if ((state.emailMenuOpen || state.saveMenuOpen || state.moreMenuOpen) && !(e.target as HTMLElement).closest('.menu-wrap')) {
    state.emailMenuOpen = false;
    state.saveMenuOpen = false;
    state.moreMenuOpen = false;
    render();
  }
});
document.addEventListener('paste', (e) => {
  if ((e.target as HTMLElement | null)?.closest?.('textarea, input')) return;
  const r = reportRowAt(e.target);
  const files = pastedImages(e);
  if (!r || !files.length) return;
  e.preventDefault();
  void addReports(r, files);
});
document.addEventListener('keydown', (e) => {
  const v = state.viewer;
  if (v) {
    const count = state.grid?.rows.find((r) => rowKey(r) === v.key)?.reports?.length ?? 0;
    if (e.key === 'Escape') state.viewer = null;
    else if (e.key === 'ArrowLeft' && v.index > 0) state.viewer = { ...v, index: v.index - 1, actualSize: false };
    else if (e.key === 'ArrowRight' && v.index < count - 1) state.viewer = { ...v, index: v.index + 1, actualSize: false };
    else return;
    e.preventDefault();
    render();
    return;
  }
  if (e.key === 'Escape' && (state.emailMenuOpen || state.saveMenuOpen || state.moreMenuOpen)) {
    state.emailMenuOpen = false;
    state.saveMenuOpen = false;
    state.moreMenuOpen = false;
    render();
  }
});
render();
void loadStoredTemplate();

// Newer settings from another device (if signed in to sync) replace these once they arrive.
pullBeforeStart(SYNC_APP, STORAGE_KEY).then((result) => {
  if (result === 'pulled') applySyncedSettings();
});

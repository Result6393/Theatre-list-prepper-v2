import { downloadText, h } from './dom';
import { DEFAULT_MESSAGE, DEFAULT_SUBJECT, isEmailAddress } from './email';
import { FIELDS, RUN_SHEET_COLUMNS, defaultLayout } from './layout';
import type { Table } from './parse';
import { defaultPrompts, promptColumns } from './prompts';
import { defaultSettings, exportSettings, importSettings, parseList, type Settings } from './settings';

interface Props {
  settings: Settings;
  table: Table | null;
  /** Persist after an in-place edit that doesn't need a redraw. */
  save: () => void;
  rerender: () => void;
  replace: (settings: Settings) => void;
  done: () => void;
}

export function renderSettings(props: Props): HTMLElement {
  return h(
    'main',
    { class: 'settings' },
    renderColumns(props),
    renderPrompts(props),
    renderHospitals(props),
    renderEmail(props),
    renderBackup(props),
    h('div', { class: 'copy-bar' }, h('button', { class: 'primary', onclick: props.done }, 'Done')),
  );
}

function section(title: string, ...body: (Node | string | null)[]): HTMLElement {
  return h('section', { class: 'step' }, h('h2', {}, title), ...body);
}

function changed(props: Props): void {
  props.save();
  props.rerender();
}

// ---- Columns ----------------------------------------------------------------------

function renderColumns(props: Props): HTMLElement {
  const { layout } = props.settings;
  const rawColumns = [...new Set([...RUN_SHEET_COLUMNS, ...(props.table?.headers ?? [])])].filter((c) => c !== 'TYPE');

  const sourceSelect = (value: string, onChange: (v: string) => void) => {
    const option = (v: string, label: string) => h('option', { value: v, selected: v === value }, label);
    const select = h(
      'select',
      { 'aria-label': 'Column source', onchange: (e: Event) => onChange((e.target as HTMLSelectElement).value) },
      option('blank', 'Blank (fill in by hand)'),
      h('optgroup', { label: 'From the run sheet, cleaned up' }, FIELDS.map((f) => option(`field:${f.id}`, f.label))),
      h('optgroup', { label: 'Run sheet column, as listed' }, rawColumns.map((c) => option(`raw:${c}`, c))),
    );
    select.value = value;
    return select;
  };

  const move = (from: number, to: number) => {
    if (to < 0 || to >= layout.columns.length) return;
    const [col] = layout.columns.splice(from, 1);
    layout.columns.splice(to, 0, col);
    changed(props);
  };

  return section(
    'Columns',
    h(
      'p',
      { class: 'hint' },
      'Match these to the columns of your own sheet, left to right. Use Blank for columns you fill in by hand. Changing columns resets edits on the list you’re prepping.',
    ),
    h(
      'ol',
      { class: 'columns' },
      layout.columns.map((col, i) =>
        h(
          'li',
          {},
          h('span', { class: 'col-letter' }, columnLetter(layout.pasteStart, i)),
          h('input', {
            type: 'text',
            value: col.heading,
            'aria-label': 'Column heading',
            onchange: (e: Event) => {
              col.heading = (e.target as HTMLInputElement).value;
              props.save();
            },
          }),
          sourceSelect(col.source, (v) => {
            col.source = v;
            props.save();
          }),
          h('button', { class: 'mini', title: 'Move left', disabled: i === 0, onclick: () => move(i, i - 1) }, '↑'),
          h('button', { class: 'mini', title: 'Move right', disabled: i === layout.columns.length - 1, onclick: () => move(i, i + 1) }, '↓'),
          h(
            'button',
            {
              class: 'mini',
              title: 'Remove column',
              onclick: () => {
                layout.columns.splice(i, 1);
                changed(props);
              },
            },
            '×',
          ),
        ),
      ),
    ),
    h(
      'div',
      { class: 'inline-form' },
      h(
        'button',
        {
          onclick: () => {
            layout.columns.push({ heading: '', source: 'blank' });
            changed(props);
          },
        },
        'Add column',
      ),
      h(
        'button',
        {
          class: 'ghost',
          onclick: () => {
            if (!confirm('Reset columns to the default cheat-sheet layout?')) return;
            props.settings.layout = defaultLayout();
            changed(props);
          },
        },
        'Reset to default columns',
      ),
    ),
    h(
      'div',
      { class: 'inline-form' },
      h(
        'label',
        {},
        'Paste starts at cell ',
        h('input', {
          type: 'text',
          class: 'short',
          value: layout.pasteStart,
          onchange: (e: Event) => {
            layout.pasteStart = (e.target as HTMLInputElement).value.trim().toUpperCase();
            changed(props);
          },
        }),
      ),
      h(
        'label',
        {},
        h('input', {
          type: 'checkbox',
          checked: layout.includeHeadings,
          onchange: (e: Event) => {
            layout.includeHeadings = (e.target as HTMLInputElement).checked;
            props.save();
          },
        }),
        ' Copy the headings as the first row',
      ),
    ),
  );
}

/** Spreadsheet column letter for the i-th pasted column, e.g. start "B5", i 2 → "D". */
export function columnLetter(pasteStart: string, i: number): string {
  const letters = pasteStart.match(/^[A-Z]+/i)?.[0].toUpperCase() ?? 'A';
  let n = [...letters].reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) + i;
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

// ---- Cataract prompts -------------------------------------------------------------

function renderPrompts(props: Props): HTMLElement {
  const { prompts, layout } = props.settings;
  const resolved = promptColumns(layout, prompts);
  const lines = (label: string, hint: string, value: string[], onChange: (v: string[]) => void) =>
    h(
      'label',
      { class: 'stacked' },
      label,
      h('span', { class: 'hint small' }, hint),
      h('textarea', {
        rows: Math.max(3, value.length + 1),
        value: value.join('\n'),
        onchange: (e: Event) => {
          onChange((e.target as HTMLTextAreaElement).value.split('\n').map((l) => l.trim()).filter(Boolean));
          changed(props);
        },
      }),
    );
  const columnSelect = (label: string, value: string, auto: number, onChange: (v: string) => void) => {
    const autoName = auto >= 0 ? layout.columns[auto].heading || `column ${columnLetter(layout.pasteStart, auto)}` : 'none found';
    const select = h(
      'select',
      { onchange: (e: Event) => (onChange((e.target as HTMLSelectElement).value), changed(props)) },
      h('option', { value: '' }, `Automatic (${autoName})`),
      layout.columns.map((c, i) => h('option', { value: c.heading }, `${columnLetter(layout.pasteStart, i)}: ${c.heading || '(no heading)'}`)),
    );
    select.value = value;
    return h('label', {}, `${label} `, select);
  };
  return section(
    'Cataract prompts',
    h(
      'p',
      { class: 'hint' },
      'Clickable prompts under each cataract or lens patient. A click adds it to the cell; click again to take it out. Nothing is added unless you pick it.',
    ),
    h(
      'label',
      {},
      h('input', {
        type: 'checkbox',
        checked: prompts.enabled,
        onchange: (e: Event) => {
          prompts.enabled = (e.target as HTMLInputElement).checked;
          changed(props);
        },
      }),
      ' Show cataract prompts',
    ),
    lines('Findings', 'One per line. Separate choices with | to pick just one, e.g. Good dilation | Mod dilation | Poor dilation.', prompts.findings, (v) => (prompts.findings = v)),
    lines('Lenses', 'One per line. Put x where you type a number, e.g. CNA0Tx for the toric power (CNA0T4).', prompts.lenses, (v) => (prompts.lenses = v)),
    lines(
      'A-constants (SRK/T)',
      'One per line: the lens and its SRK/T A-constant as on your IOLMaster, e.g. DCB00 119.3. A lens that isn’t on a patient’s IOL report is then estimated from one that is (put x where a number goes, e.g. CNA0Tx 119.1).',
      prompts.constants,
      (v) => (prompts.constants = v),
    ),
    h('h4', { class: 'settings-sub' }, 'Biometry alerts'),
    h('p', { class: 'hint' }, 'From the IOL report, for the eye being operated on: flagged in amber under the report and on the patient (an eye that’s already pseudophakic is red). Change the cutoffs to suit your practice.'),
    h(
      'div',
      { class: 'limits-grid' },
      ([
        ['alVeryShort', 'Very short eye: AL under (mm)'],
        ['alShort', 'Short eye: AL under (mm)'],
        ['alLong', 'Long eye: AL over (mm)'],
        ['alVeryLong', 'Very long eye: AL over (mm)'],
        ['acdShallow', 'Shallow AC: ACD under (mm)'],
        ['ltThick', 'Thick lens: LT over (mm)'],
        ['kFlat', 'Flat cornea: K under (D)'],
        ['kSteep', 'Steep cornea: K over (D)'],
        ['toricCyl', 'Consider a toric: ΔK from (D)'],
        ['kAgreeCyl', 'ΔK and ΔTK differ by over (D)'],
        ['kAgreeAxis', '… or their axes by over (°)'],
        ['alDiff', 'Eyes’ AL differ by over (mm)'],
        ['kDiff', 'Eyes’ K differ by over (D)'],
        ['wtwSmall', 'Small cornea: WTW under (mm)'],
        ['wtwLarge', 'Large cornea: WTW over (mm)'],
      ] as [keyof typeof prompts.limits, string][]).map(([key, label]) =>
        h(
          'label',
          { class: 'limit' },
          h('span', {}, label),
          h('input', {
            type: 'number',
            step: '0.1',
            value: String(prompts.limits[key]),
            onchange: (e: Event) => {
              const n = Number((e.target as HTMLInputElement).value);
              if (Number.isFinite(n)) {
                prompts.limits[key] = n;
                changed(props);
              }
            },
          }),
        ),
      ),
    ),
    h(
      'div',
      { class: 'inline-form' },
      columnSelect('Findings go in', prompts.findingsColumn, promptColumns(layout, { ...prompts, findingsColumn: '' }).findings, (v) => (prompts.findingsColumn = v)),
      columnSelect('Lens and target go in', prompts.lensColumn, promptColumns(layout, { ...prompts, lensColumn: '' }).lens, (v) => (prompts.lensColumn = v)),
    ),
    resolved.findings === -1 || resolved.lens === -1
      ? h('p', { class: 'error' }, `No column chosen for the ${resolved.findings === -1 ? 'findings' : 'lens and target'}, so ${resolved.findings === -1 && resolved.lens === -1 ? 'they aren’t' : 'it isn’t'} shown.`)
      : null,
    h(
      'div',
      { class: 'inline-form' },
      h(
        'button',
        {
          class: 'ghost',
          onclick: () => {
            props.settings.prompts = defaultPrompts();
            changed(props);
          },
        },
        'Reset prompts',
      ),
    ),
  );
}

// ---- Hospitals --------------------------------------------------------------------

function renderHospitals(props: Props): HTMLElement {
  const { settings } = props;
  const listInput = (value: string[], label: string, onChange: (v: string[]) => void) =>
    h(
      'label',
      { class: 'stacked' },
      label,
      h('input', {
        type: 'text',
        value: value.join(', '),
        onchange: (e: Event) => {
          onChange(parseList((e.target as HTMLInputElement).value).map((v) => v.toUpperCase()));
          props.save();
        },
      }),
    );
  return section(
    'Hospitals',
    listInput(settings.flagHospitals, 'Flag in Plan & notes when found anywhere on a patient’s row (e.g. “POWH patient”)', (v) => {
      settings.flagHospitals = v;
    }),
    listInput(settings.hospitalCodes, 'Hospital codes to drop from comments when they appear on their own', (v) => {
      settings.hospitalCodes = v;
    }),
  );
}

// ---- Email ------------------------------------------------------------------------

function renderEmail(props: Props): HTMLElement {
  const { email } = props.settings;
  const addresses = (value: string[], label: string, onChange: (v: string[]) => void) => {
    const warning = h('span', { class: 'error' });
    const showWarning = (list: string[]) => {
      const bad = list.filter((a) => !isEmailAddress(a));
      warning.textContent = bad.length ? `Not a valid address, so it will be ignored: ${bad.join(', ')}` : '';
    };
    showWarning(value);
    return h(
      'label',
      { class: 'stacked' },
      label,
      h('input', {
        type: 'text',
        value: value.join(', '),
        placeholder: 'name@example.org, another@example.org',
        onchange: (e: Event) => {
          const list = parseList((e.target as HTMLInputElement).value);
          onChange(list);
          showWarning(list);
          props.save();
        },
      }),
      warning,
    );
  };
  const textField = (label: string, value: string, onChange: (v: string) => void, placeholder = '') =>
    h(
      'label',
      { class: 'stacked' },
      label,
      h('input', {
        type: 'text',
        value,
        placeholder,
        onchange: (e: Event) => {
          onChange((e.target as HTMLInputElement).value);
          props.save();
        },
      }),
    );

  return section(
    'Email to Medical Records',
    h(
      'p',
      { class: 'hint' },
      'Used by the “Email Medical Records” button to open a draft in Outlook on the web asking for the paper files of the ticked patients. These addresses are saved only in this browser (and your exported settings file).',
    ),
    addresses(email.to, 'To', (v) => (email.to = v)),
    addresses(email.cc, 'Cc (optional)', (v) => (email.cc = v)),
    textField('Your name (for the sign-off)', email.name, (v) => (email.name = v.trim())),
    textField('Subject', email.subject, (v) => (email.subject = v.trim() || DEFAULT_SUBJECT)),
    h(
      'label',
      { class: 'stacked' },
      'Message',
      h('textarea', {
        rows: 9,
        value: email.message,
        onchange: (e: Event) => {
          email.message = (e.target as HTMLTextAreaElement).value.trim() || DEFAULT_MESSAGE;
          props.save();
        },
      }),
    ),
    h(
      'p',
      { class: 'hint small' },
      'Placeholders: {day} (e.g. Tuesday 13/10), {date}, {patients} (one “- MRN NAME” line each), {name}, {theatre} (e.g. OR 3 PM), {count}.',
    ),
    textField('Outlook on the web address', email.outlookUrl, (v) => (email.outlookUrl = v.trim()), 'https://outlook.cloud.microsoft/mail/'),
    h(
      'div',
      { class: 'inline-form' },
      h(
        'button',
        {
          class: 'ghost',
          onclick: () => {
            email.subject = DEFAULT_SUBJECT;
            email.message = DEFAULT_MESSAGE;
            changed(props);
          },
        },
        'Reset wording',
      ),
    ),
  );
}

// ---- Backup -----------------------------------------------------------------------

function renderBackup(props: Props): HTMLElement {
  const fileInput = h('input', {
    type: 'file',
    accept: '.json,application/json',
    hidden: true,
    onchange: async (e: Event) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;
      try {
        props.replace(importSettings(await file.text()));
      } catch {
        alert('That file isn’t a Theatre List Prepper settings file.');
      }
    },
  });
  return section(
    'Back up or move your settings',
    h(
      'p',
      { class: 'hint' },
      'Settings are saved in this browser only. Export them to a file to keep a backup, use them on another computer, or share them with a colleague. No patient details are ever saved.',
    ),
    h(
      'div',
      { class: 'inline-form' },
      h('button', { onclick: () => downloadText('theatre-list-prepper-settings.json', exportSettings(props.settings)) }, 'Export settings'),
      h('button', { onclick: () => fileInput.click() }, 'Import settings'),
      fileInput,
      h(
        'button',
        {
          class: 'ghost danger',
          onclick: () => {
            if (!confirm('Clear all settings, including your columns?')) return;
            props.replace(defaultSettings());
          },
        },
        'Clear all settings',
      ),
    ),
  );
}

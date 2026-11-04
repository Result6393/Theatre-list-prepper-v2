# Theatre List Prepper

Turns a whole day's theatre run sheet (SurgiNet SN006A export) into rows for one list's consultant cheat sheet, ready to paste into Excel.

1. **Paste the run sheet** copied from Excel, or choose/drop the file (CSV or Excel). Optionally, also choose your **previously prepped** Excel sheet (see below).
2. **Click the list you're prepping** (theatre + AM/PM, e.g. "OR 3 · PM · Smith").
   - **Lists that run on together:** tick more than one list (e.g. OR 3 AM and PM without a break) to prep them as one list. Each patient shows which list they're on. Edits you've already made on a list carry over when you tick it in with another.
   - The run sheet can cover several days; lists are grouped by day.
3. **Check and edit** the rows: untick patients you don't want, type in any cell, and **drag** the ⠿ handle to reorder (or click the handle and use the ↑ ↓ keys). Edits and unticks are remembered while you switch between lists.
4. **Copy for Excel**, click the first cell of your sheet (B5 by default) and press Ctrl+V.
5. **Email Medical Records ▾** (optional) asks for the paper files of the ticked patients on all the ticked lists, grouped under a heading per list when there are several (e.g. "…patients having surgery on Tuesday 13/10" and a list of `MRN NAME`). Check the draft and press Send. Choose how to open it:
   - **Outlook on the web** opens a draft in a new tab.
   - **Email app** opens a draft in the email app set up on your computer (e.g. installed Outlook), which uses your normal compose font.
   - **Share…** (phones and tablets, when the browser supports it) shares the text to Outlook or another app. Recipients can't be passed this way, so add the address in the email app.
   - **Copy email text** copies it to paste into any email.

Click any of Consultant, Date, OR or Session above the list to copy it for your sheet's header.

## Cataract prompts

Under each cataract or lens patient there's a row of one-click prompts. They're reminders only: nothing is mandatory, and nothing goes on the sheet unless you pick it. They stay faded until you point at (or type in) that patient, and anything you've picked always shows.

- **Findings:** `PXF`, `Guttata`, dilation (`Good` / `Mod` / `Poor`, pick one), `Phacodonesis`, `Tamsulosin`. Each click adds it to *Diagnosis & relevant history* on one line (e.g. `PXF, Poor dilation`); click again to take it out.
- **Lens:** `DCB00`, `DIB00`, `ZCU…`, `CNA0T…`. Toric lenses get a box for the number (e.g. `4` → `CNA0T4`), and any lens a box for the power (`21.5` → `+21.5D`). It goes on the first line of *Implants / to order*.
- **Target:** one click for `Plano`, or type one with its sign (e.g. `-0.5` → `Target -0.50`). Goes under the lens.

The prompts just edit the cell text, so you can still type in the cell, and anything you type yourself (like `PXF`) lights up its button. Picks also come back when you update a list from your sheet. Change the findings, lenses and columns in **Settings → Cataract prompts** (put `x` where a lens code takes a number, and separate pick-one choices with `|`), or turn the prompts off.

## Your prepped Excel file

Choose your cheat sheet (`.xlsx`) under **Previously prepped?** in step 1. Download it from OneDrive first if needed. It's read on this computer and never uploaded.

- **Straight away, without a run sheet:** each sheet with patient rows becomes a list in step 2. Open one to see the patients and edit, reorder, tick and use the prompts as usual. Times, New, Moved and the Medical Records email need the run sheet.
- **With the run sheet too** (load it before or after): each list you open is updated from the file, see below. Edits you made before loading the run sheet are kept.
- **Save to *file*.xlsx** writes your lists back into the same file (**▾** next to it has both options):
  - **Save to sheet "…"**: only the cells you changed in each list's block are written. Formatting, formulas, other sheets and everything else stay as they were, and new rows take the formatting of the row above. If a list has grown and would write over something below it, it asks first.
  - **Save as a new sheet…**: adds a sheet (name it, e.g. "OR 4 AM 14-10") that starts as a copy of one of your sheets, keeping its headings, column widths, formatting and formulas, with this list in place of that sheet's patients. Works for any list, including one whose patients aren't in the file yet (e.g. tomorrow's list). Anything else on the copied sheet, like a consultant or date at the top, is copied as it was, so update it in Excel if it's changed. After that, the list saves into its new sheet.
  - **Where it saves:** over the file you opened, when you chose it with **Choose prepped Excel sheet…** in Edge or Chrome (the browser may ask first). If it can't (you dropped the file on the page, didn't allow it, or the file is open in Excel), it asks where to save: choose the original file to replace it, and later saves go straight there. Only if the browser can't do that either does it download an updated copy to your Downloads folder.
  - A message under the toolbar says exactly where it saved, or why it couldn't save over the file, and stays until you dismiss it.
  - **Close the file in Excel before saving**: Excel locks it on Windows, and an open window won't show the changes until you reopen it. If the file was changed (e.g. saved in Excel) since you opened it here, the app asks before replacing it.

## When the list changes on the day

Patients dropped, added or reordered after you prepped? Keep your notes without retyping them.

**From your Excel file (easiest):**

1. **Re-export the run sheet** and load it in step 1.
2. Under **Previously prepped?** choose your cheat sheet. You can do this before or after the run sheet.
3. Click the list. It's updated from the file straight away: the rows are found from the first MRN to the last, so headings and notes around them don't matter.
   - **Several sheets** (e.g. one per theatre list): each list uses the sheet that fits it best (most of its patients, fewest others). To change it, use **Sheet** above the list; each option says how many of this list's patients it has. **Don't use** puts the list back as on the run sheet.
   - Copying puts the rows back where they were found in the file (e.g. B5), or use **Save to *file*.xlsx**.

**Or by copying the rows:** in your cheat sheet, **select all your patient rows** (from B5 down to the last one, including any cancelled rows), copy them, click **Update from my sheet…** and press Ctrl+V in the box. (That box also has a "choose your Excel file" link.)

Then check the list. Patients are matched by MRN:
- **Patients still on the list** keep everything you typed, in the new order. Anyone whose place changed is marked **Moved, was #n** (both patients, when two swap).
- **New patients** are filled in from the run sheet and marked **New**, so you can see who still needs prep.
- **Patients no longer on the list** are kept at the bottom, marked **Cancelled** (with the list they're on now, if they moved to another one). Untick any you don't want to keep. Cancelled patients are never put in the Medical Records email.
- If the run sheet now says something different from your sheet (e.g. the eye or procedure changed), the cell shows **Now: …** with **Use** or **Keep**.

Then click **Copy for Excel**, click B5 (or wherever the rows were) and press Ctrl+V. This replaces the old rows. Cancelled patients go under a `CANCELLED` row at the bottom, and blank rows clear anything left over.

You can do this as many times as the list changes. **Undo** puts the list back as it was before the update. Nothing is saved anywhere: your sheet stays the record.

The Patient (or MRN) column must be one of your columns, so rows can be matched.

## What it fills in

With the default columns (B–K of the consultant cheat sheet):

| Column | Filled with |
|---|---|
| Eye | `RIGHT` / `LEFT` / `BOTH`, from the procedure |
| Patient | `MRN SURNAME, Given names AgeSex` |
| Procedure | Shortened, e.g. `LE CE/IOL`, `RE CE/IOL + iStent inject`, `RE CE/toric IOL`, `RE DCR`, `LE trabeculectomy + MMC`. Any part of a procedure the app doesn't recognise is kept on the end (e.g. `LE CE/IOL + capsular tension ring`); procedures it doesn't recognise at all are kept as listed |
| Diagnosis & relevant history | Interpreter language, and `Diabetic` (with Ozempic etc.) from the comments |
| Exam (VA / IOP) | blank |
| Drops / allergies | `NKA`, `Allergy: Y` or `Allergy not recorded` |
| Implants / to order, Ordered | blank |
| Alerts | Clinical alerts, one per line |
| Plan & notes | `POWH patient` / `STG patient` if that hospital appears anywhere on the row, plus any other comment text |

## Settings

- **Columns:** match your own sheet. Add, remove, rename and reorder columns. Each one is blank, a cleaned-up field, or a run-sheet column as listed. You can also set where the paste starts and whether to copy a heading row.
- **Email to Medical Records:** the To/Cc addresses, your name for the sign-off, and the subject and wording (with placeholders such as `{day}` and `{patients}`). Nothing is built in, so the addresses stay out of the repo.
- **Hospitals:** which hospitals to flag, and which hospital codes to drop from comments.
- **Sync across devices:** sign in with your sync password at the top of Settings and your settings follow you to every device where you've signed in. They're encrypted in your browser before upload, so the sync server (a Cloudflare Worker, see [Result6393.github.io](https://github.com/Result6393/Result6393.github.io)) only stores scrambled data. Works on the GitHub Pages version; the downloaded single-file copy can't reach it and just uses its own settings.
- **Back up / move:** export your settings to a file and import them on another computer, or share them with a colleague. To import, just drag the settings (`.json`) file onto the page, anywhere (or use the Import settings button). If you've already changed settings, the app asks before replacing them.

## Privacy

An exported settings file can contain email addresses (such as the Medical Records recipient), so it is never kept in this repository: `.gitignore` excludes `*settings*.json`. Share it with colleagues directly instead.

Everything happens in your browser. The run sheet and your prepped Excel file are read on your computer and never uploaded, stored or sent anywhere, and the app makes no network requests apart from settings sync, if you turn it on. Only your settings (columns, hospital flags, email addresses and wording) are saved in the browser. If you sign in to sync, those settings (encrypted) are also the only thing ever sent to the sync server.

The one exception is the **Email Medical Records** button: you press it, and the patients' MRNs and names are placed in a link that opens a draft in your own email (or the share sheet on your phone). Nothing is sent until you press Send there.

Never commit real run sheets or theatre lists to this repository. `.gitignore` excludes spreadsheet and CSV files, and the tests use invented patients and surgeons.

## Getting the app

- **Web:** the GitHub Pages site (built from `main`).
- **Offline:** download `theatre-list-prepper-vX.Y.Z.html` from the latest release and double-click it. It's a single file that works without internet.

## Versions

The version number is in `package.json` and appears in the page title and the top bar. To release: bump `version` in `package.json` (e.g. `npm version minor --no-git-tag-version`) and merge to `main`. The release workflow then publishes `theatre-list-prepper-vX.Y.Z.html` as release `vX.Y.Z`, creating the tag. A version that already has a release is skipped.

## Development

```sh
npm install
npm run dev    # local dev server
npm test       # unit tests
npm run build  # single self-contained dist/index.html
```

The extraction rules (eye, procedure abbreviations, comment parsing) are in `src/extract.ts`. Add a procedure abbreviation to `PROCEDURE_RULES` (its `words` list stops those words being repeated on the end), with a test in `tests/extract.test.ts`.

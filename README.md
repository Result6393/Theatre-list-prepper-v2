# Theatre List Prepper

Turns a whole day's theatre run sheet (SurgiNet SN006A export) into rows for one list's consultant cheat sheet, ready to paste into Excel.

1. **Paste the run sheet** copied from Excel, or choose/drop the CSV file.
2. **Click the list you're prepping** (theatre + AM/PM, e.g. "OR 3 · PM · Smith").
3. **Check and edit** the rows: untick patients you don't want, reorder, type in any cell.
4. **Copy for Excel**, click the first cell of your sheet (B5 by default) and press Ctrl+V.
5. **Email Medical Records** (optional), asking for the paper files of the ticked patients (e.g. "…patients having surgery on Tuesday 13/10" and a list of `MRN NAME`). Check the draft and press Send. Choose how to open it:
   - **Email Medical Records (Outlook on the web)** opens a draft in a new tab.
   - **Open in email app** opens a draft in the email app set up on your computer (e.g. installed Outlook), which uses your normal compose font.
   - **Share…** (phones and tablets, when the browser supports it) shares the text to Outlook or another app. Recipients can't be passed this way, so add the address in the email app.
   - **Copy email text** copies it to paste into any email.
   - **One email for several lists:** under the table, tick other lists in "Include in the Medical Records email". Their ticked patients are added, grouped under a heading per list. The run sheet can cover several days; lists are grouped by day and the email names each day. Edits and unticks on each list are remembered while you switch between them. (Copy for Excel is still one list at a time.)

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

Everything happens in your browser. The run sheet is never uploaded, stored or sent anywhere, and the app makes no network requests apart from settings sync, if you turn it on. Only your settings (columns, hospital flags, email addresses and wording) are saved in the browser. If you sign in to sync, those settings (encrypted) are also the only thing ever sent to the sync server.

The one exception is the **Email Medical Records** button: you press it, and the patients' MRNs and names are placed in a link that opens a draft in your own email (or the share sheet on your phone). Nothing is sent until you press Send there.

Never commit real run sheets or theatre lists to this repository. `.gitignore` excludes spreadsheet and CSV files, and the tests use invented patients and surgeons.

## Getting the app

- **Web:** the GitHub Pages site (built from `main`).
- **Offline:** download `theatre-list-prepper-vX.Y.Z.html` from the latest release and double-click it. It's a single file that works without internet.

## Versions

The version number is in `package.json` and appears in the page title and the top bar. To release: bump `version` in `package.json` (e.g. `npm version minor --no-git-tag-version`), merge, then push a matching tag (`v1.1.0`). The release workflow refuses a tag that doesn't match.

## Development

```sh
npm install
npm run dev    # local dev server
npm test       # unit tests
npm run build  # single self-contained dist/index.html
```

The extraction rules (eye, procedure abbreviations, comment parsing) are in `src/extract.ts`. Add a procedure abbreviation to `PROCEDURE_RULES` (its `words` list stops those words being repeated on the end), with a test in `tests/extract.test.ts`.

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

## Standard and Compact views

On a computer, the **Standard | Compact** switch in the toolbar (which stays at the top as you scroll) chooses how the list looks (remembered on that computer). **⇅ Reorder** beside *Copy lens list* opens a short list of the patients to put in order (drag ⠿, or ↑ ↓), without scrolling through the cards. While you drag (there, or the ⠿ on a card or row), the others slide apart to open a gap where the patient will go, outlined in dashed blue.

- **Standard** (a card per patient): the patient's number on the list (top left), name, eye (**R / L / B**) and time along the top, with a **To do** badge listing what's still missing (1st/2nd eye, cataract grade, VA/IOP, lens) or **✓ Ready**; *Exam (VA / IOP)* and *Drops / allergies* on one line; then *Diagnosis* with 1st / 2nd eye, the findings and dilation, beside the *Lens* panel: the lens in one dropdown (with its number for one like `ZCU…`), the power with **−** / **+** in 0.5 D steps (starting from the report's best power for the target, once a report has been read), the target (**Plano**, **−0.25**, **−0.50** or other), the *Implants* box, and the IOL report (click the box under it and paste, or drop a screenshot on it). *Procedure*, *Ordered*, *Alerts*, *Plan & notes* and any other columns are along the bottom.
- **Compact**: the columns as a table, like your sheet, with the prompts under each patient.

Both edit the same cells, so your Excel file is the same whichever you use.

## Locking

Every patient **starts locked** against changes (🔒 beside their number): their boxes are read-only (you can still select and copy the text), and the buttons, lens dropdown, tick, drag handle and pasting a screenshot are off. Opening their IOL report still works (Apply in it is off). Tap **🔒** to unlock one (and again to lock it), or **🔓 Unlock all / 🔒 Lock all** (in the toolbar on a computer, beside the tick count on a phone). Patients you unlock stay unlocked while the page is open (and through updates from the run sheet); reopening the file locks everyone again.

## On a phone

On a narrow screen the list becomes a stack of patient cards instead of a wide table, so there's no scrolling sideways:

- Each card shows the eye (R / L / B), the patient, procedure and time, any New / Moved / Cancelled badge, and alerts. Tap it to open it: every column is a full-width box, the eye is three big buttons, and the cataract prompts sit under the box they fill in (findings under Diagnosis, lens and target under Implants). One patient is open at a time, and their name stays at the top of the screen as you scroll through them.
- Move a patient by dragging its ⠿ handle with your finger, or with **↑ Move up** / **↓ Move down** in the open card.
- **Copy**, **Save** and **⋯** (Update from my sheet, Email Medical Records) are in a bar at the bottom of the screen.

## Allergies

When the run sheet says a patient has an allergy (`Y`) but not what it is, their *Drops / allergies* box turns **red**, with an **Allergy: add details** badge, an **Acknowledge** button, and a banner over the list naming them. It stops being red when you type the details in the box (or press Acknowledge to say you've seen it and will add them later; the box then goes amber). Copying for Excel or saving into a file asks first if any ticked patient still has an allergy with no details. Nothing about allergies is stored: it's worked out from the run sheet and your box each time.

## Cataract prompts

Under each cataract or lens patient there's a row of one-click prompts. They're reminders only: nothing is mandatory, and nothing goes on the sheet unless you pick it. They stay faded until you point at (or type in) that patient, and anything you've picked always shows.

- **Cataract grade:** a box beside 1st / 2nd eye (shaded yellow until filled in). The diagnosis box starts with which eye and the grade on one line, e.g. `1st eye: NS2+ C1` (or `Cataract: NS2+ C1` until the eye is picked), then the findings and history.
- **1st / 2nd eye:** two buttons (shaded yellow until you pick one), `1st` and `2nd`, put `1st eye` or `2nd eye` on its own line at the top of *Diagnosis & relevant history* (picking the other one switches it; click again to take it out). It also recognises `1st eye` / `2nd eye` you typed yourself, anywhere in the box.
- **Findings:** `PXF`, `Guttata`, `Phacodonesis`, `Tamsulosin` take three clicks: present (`PXF`), then not present (`No PXF`), then back to saying nothing. Dilation (`Good` / `Mod` / `Poor`) is pick one (click again to take it out). They go in *Diagnosis & relevant history* on one line (e.g. `PXF, No guttata, Poor dilation`), and `no …` or `nil …` you type yourself lights up the not-present state.
- **VA / IOP:** an empty *Exam (VA / IOP)* box is shaded yellow until it's filled in.
- **Lens:** `DCB00`, `DIB00`, `ZCU…`, `DIU…`, `CNA0T…`. Toric lenses get a box for the number (e.g. `4` → `CNA0T4`), and any lens a box for the power (`21.5` → `+21.5D`). It goes on the first line of *Implants / to order*.
- **Target:** one click for `Plano`, or type one with its sign (e.g. `-0.5` → `Target -0.50`). Goes under the lens.

- **IOL report:** a screenshot of the patient's IOL calculation (e.g. from the IOLMaster 700), kept with them to check the lens against at timeout. Click in the patient's *Implants* box (or on **+ Add screenshot**) and paste it with Ctrl+V, drop it on the box, or choose the file. It shows as a thumbnail; click it to open it full screen, with the eye, lens, power and target you picked in large text above it (**Actual size** zooms in, **‹ ›** or the arrow keys go between screenshots, Esc closes). **×** removes one (it asks right beside the thumbnail). **Re-read** (under the report, and in the full-screen view) reads a screenshot again from scratch. Screenshots are saved as pictures on the patient's row in your Excel file, just right of the list, and they move with the patient when the list is reordered or updated. A picture you paste on a patient's row in Excel shows up here too. Click one in Excel and drag a corner to see it bigger. They're only kept by saving to Excel; **Copy for Excel** copies text only.
- **Report check:** each screenshot is read by the app (OCR, on this computer, offline; it takes a few seconds) and compared with the lens you picked. Under the thumbnail, and in the full-screen view, it shows what the report says for your lens in the eye being operated on, e.g. `✓ Right eye: DCB00 +20.50 → predicted −0.33`, with the axis for a toric (`ZCU100 +22.00, axis 167°`). It flags in **red** a power that's in the other eye's column, or a report with another patient's name, and in **amber** a power that isn't on the report, a lens that isn't on it, or a target other than plano or −0.50 that the report wasn't calculated for (the IOLMaster is usually set to −0.50 to show the most useful range of lenses, so plano and −0.50 never get a note). The report's own warnings (like a difference in axial length) are shown too. It's only a cross-check: it never changes your lens, and you should always read the report itself. What OCR read is saved in your Excel file along with the picture (a hidden part of the file that Excel keeps but doesn't show), so a reopened file is ready at once instead of being read again; it's never kept in the browser. If a screenshot can't be read, the message says whether the reading itself failed or no IOL calculation was found in it, with **Read again**. **The name must match:** a report with another patient's name on it (surname, or given name) shows a red **NAME MISMATCH** banner on the patient, under the report and across the top of the full-screen view. A different patient ID is only a note, as the IOLMaster often uses another system's number.
- **Biometry alerts:** from the report, for the eye being operated on, flagged in amber under the report, in the full-screen view, and as badges on the patient: a short or long eye (AL under 22.0 / over 26.0 mm, *very* under 21.0 / over 28.0), shallow AC (ACD under 2.5 mm), thick lens (LT over 5.0 mm), flat or steep cornea (K under 41 / over 47 D), corneal astigmatism of 1.0 D or more with a non-toric lens (consider a toric), ΔK and ΔTK not agreeing (by over 0.5 D or 15°), a small or large cornea (WTW under 11.0 / over 12.5 mm), values the IOLMaster marks borderline (!), previous laser vision correction, and the two eyes' AL or K differing (over 0.3 mm / 1.0 D). An eye that's already **pseudophakic** (or aphakic) is red. The values are boxed on the screenshot. Change the cutoffs in **Settings → Cataract prompts → Biometry alerts**.
- **Highlights:** in the full-screen view, the screenshot has boxes over the patient's name, date of birth and MRN, the eye's heading (OD / OS), ΔK, the lens's name and the power you picked: red for a right eye, blue for a left (purple if the eye isn't set). Dashed boxes mark what was looked at but isn't a match: the other eye's column when your power is there, or the lens an estimate is worked out from.
- **A lens that isn't on the report** (e.g. CNA0T0 when the report has DCB00): it's estimated from one that is, using the A-constants in **Settings → Cataract prompts → A-constants** (`DCB00 119.3`, `CNA0Tx 119.1` to start with; use the values on your IOLMaster). The report's Barrett numbers for that lens are moved by the power difference between the two lenses' A-constants in this eye, worked out with SRK/T (a published formula) from the report's AL and K, or as the A-constants' difference if those can't be read. The full-screen view shows the working step by step: the A-constants used, that a lower A-constant needs a lower power (and that keeping the same power gives a more myopic result, by how much), the SRK/T powers for each lens, the power it's equivalent to on the report and the report's rows either side, the best power for the target, and a table of powers with their estimated refractions. It's an estimate, in amber: always check it. For a toric it's the spherical equivalent only.
- **Try a lens:** beside the screenshot in the full-screen view (below it on a narrow screen), choose any of your lenses (with its number for one like `ZCU…`: the report's toric model is filled in), a target (the report's, plano or any other) and, if you like, a power. It lists the powers around the best one for that target, highest first as on the report, with the predicted refraction for each, from the report (Barrett) when the lens is on it, or estimated as above, with the working, when it isn't. For a toric on the report's toric calculation (e.g. ZCU100) it also shows the cylinder, axis and predicted residual cylinder; for a toric that isn't (or another model than the report chose), it's the spherical equivalent only, and says so. Powers marked ≈ are beyond the report's and less certain. Click a power to choose it, then **Apply** puts that lens, power and target in the patient's *Implants* box. The highlights on the screenshot follow the lens you're trying.
- **Summary ▾** (toolbar, or ⋯ on a phone): the ticked patients with every column that has something in it and what each IOL report check found, to send on (e.g. to the consultant). **Copy for email** pastes as a table in Outlook (text anywhere else); **Print / save as PDF** opens the print dialog with the list as a landscape page: choose *Save as PDF* there for a PDF.
- **Copy lens list:** one click (in the toolbar, or in the ⋯ menu on a phone) copies the whole email to the theatre coordinator, ready to paste into a new email: *Hi All*, a line naming the list and day, a table of each ticked cataract patient (eye, patient, lens, power, target), then a summary table of the lenses needed, sorted by model then power, with the number of each, and *All the best* signed with your name from the email settings. It reads the lens (and target) from the Implants box, or whatever you typed there. Patients with no lens yet are left out, and the message names them.

The prompts just edit the cell text, so you can still type in the cell, and anything you type yourself (like `PXF`) lights up its button. Picks also come back when you update a list from your sheet. Change the findings, lenses and columns in **Settings → Cataract prompts** (put `x` where a lens code takes a number, and separate pick-one choices with `|`), or turn the prompts off.

## Your prepped Excel file

Choose your cheat sheet (`.xlsx`) under **Previously prepped?** in step 1. Download it from OneDrive first if needed. It's read on this computer and never uploaded.

- **Straight away, without a run sheet:** each sheet with patient rows becomes a list in step 2. Open one to see the patients and edit, reorder, tick and use the prompts as usual. Times, New, Moved and the Medical Records email need the run sheet.
- **With the run sheet too** (load it before or after): each list you open is updated from the file, see below. Edits you made before loading the run sheet are kept.
- **Save to *file*.xlsx** writes your lists back into the same file (**▾** next to it has both options):
  - **Save to sheet "…"**: only the cells you changed in each list's block are written. Formatting, formulas, other sheets and everything else stay as they were, and new rows take the formatting of the row above. If a list has grown and would write over something below it, it asks first.
  - **The sheet's heading** is filled in when you save a list from the run sheet (to its own sheet or a new one): cells above the list labelled **Consultant** (or Surgeon), **Date**, **OR** (or Theatre), **Hospital / OR** (e.g. "SSEH OR 1"), **AM/PM**, and **Session** (AM/PM with the session times, e.g. "PM 13:00–17:30") get the list's values, in the cell to the right of the label (after it, if the label is merged across cells) or after the colon in a cell like "Consultant: Smith". A real Excel date stays a date, and the OR is written the way the sheet already writes it ("4" or "OR 4"). Column headings and other labels are never written over; the save message lists what it filled in.
  - **Save as a new sheet…**: adds a sheet (name it, e.g. "OR 4 AM 14-10") that starts as a copy of one of your sheets, keeping its headings, column widths, formatting and formulas, with this list in place of that sheet's patients. The new sheet is added as the first (leftmost) tab. Works for any list, including one whose patients aren't in the file yet (e.g. tomorrow's list). A blank template sheet works too: the list then goes at your paste cell (B5). After that, the list saves into its new sheet.
  - **Save to a new Excel file…**: a separate file with just this list, laid out like your template: headings, column widths, colours, dropdowns, notes and print settings, with the list in place of its patients and the heading filled in. Sheets without patients (like a "How to use" sheet) come along; other lists' sheets don't. Make it from a sheet of the file you have open, or from a **template file** you choose once: it's kept on this computer only (and never if it has patients in it), so it's ready next time. With no Excel file open (just the run sheet), this is the Save button, and the new file then becomes the list's file, so Save keeps it up to date.
  - **Longer lists than the sheet's rows**: rows added below take the row above's height, formatting and simple formulas (like the "#" column's `=ROW()-4`), and the sheet's colours, dropdowns and print area are extended to cover them.
  - **Where it saves:** over the file you opened, when you chose it with **Choose prepped Excel sheet…** in Edge or Chrome (the browser may ask first). If it can't (you dropped the file on the page, didn't allow it, or the file is open in Excel), it asks where to save: choose the original file to replace it, and later saves go straight there. Only if the browser can't do that either does it download an updated copy to your Downloads folder.
  - A message pinned to the bottom of the screen (so you see it however far down you are) shows **Saving…** straight away, then says exactly where it saved, or why it couldn't save over the file, and stays until you dismiss it. If the browser is waiting for you to allow it to edit the file (its question can be a small box by the address bar), the message says so and offers **Download a copy instead**. Questions like "Save anyway?" appear in that message too, not as pop-ups the browser might hide.
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

Everything happens in your browser. The run sheet and your prepped Excel file are read on your computer and never uploaded, stored or sent anywhere, and the app makes no network requests apart from settings sync, if you turn it on. Only your settings (columns, hospital flags, email addresses and wording) are saved in the browser. IOL report screenshots are held only while the page is open, and kept by saving them into your Excel file; they're read (OCR) inside the page, with the OCR engine built into the app, so nothing is downloaded or sent, and what's read is saved only in your Excel file, with the picture. If you sign in to sync, those settings (encrypted) are also the only thing ever sent to the sync server.

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

The IOL report check is in `src/iolreport.ts` (reading the report's lens tables from OCR'd words by where they sit on the page, and comparing them with the lens picked), with tests in `tests/iolreport.test.ts` built from invented reports. OCR is [Tesseract.js](https://github.com/naptha/tesseract.js) (Apache-2.0) in `src/ocr.ts`: its engine, worker and English data are built into the offline file (about 8 MB), so it works with no network.

# Runbook — migrate an existing Automated Testing Template

Goal: replace the macros + ribbon in a live operator workbook with the
canonical sources from this folder, without touching any data sheets.

---

## Recommended path: clean rebuild with `build_clean_template.py`

`build_clean_template.py` produces a clean `.xlsm` by copying the source file,
opening the copy, deleting the non-kept sheets, importing the canonical VBA, and
re-saving (Excel rewrites the whole OOXML package — no inherited calcChain rot);
it then strips the embedded web add-in and swaps in the ribbon at the zip level.
Run on the work machine (Excel + pywin32 required).

**Prerequisites:**

- `pip install pywin32` (one-time).
- Excel Trust Center → Macro Settings → **Trust access to the VBA project
  object model** = ENABLED.

**Steps:**

1. Close the Automated Testing Template in Excel.

2. Run the headless gates (no Excel needed — confirm all-PASS before building):

       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/check_sources.py
       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/test_build_helpers.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.1.xlsm"

3. Build the clean workbook — **for the v1.2 build, `--source` is the live
   v1.1 template and `--out` is the v1.2 name:**

       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/build_clean_template.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.1.xlsm" --out "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.2.xlsm"

   Guards: the source is backed up automatically to a timestamped
   `.bak-YYYYmmdd_HHMMSS` on every run; the script refuses `--out` == `--source`
   outright, and refuses an existing `--out` unless you pass `--force` (which
   first takes the same timestamped backup of the old `--out`). It also aborts
   if the source reads as MIP ciphertext (wrong interpreter).

4. Drift-verify the output (every line should be MATCH / OK):

       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/verify_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.2.xlsm"

   Expected: all modules `MATCH`, `customUI14.xml == repo`,
   `no web-extension/add-in parts`, the structure checks OK (sheet presence +
   visibility, named-range anchors incl. the literal
   `'Test Selection'!$B$4:$C$16`, featurePropertyBag, banner, `DV_LastUpload`),
   `RESULT: all modules match`. The file is still unsigned at this point —
   do not pass `--require-signature` yet.

5. **Sign the VBA project** (full flow in `signing/README.md`; one-time
   prerequisite: the owner has run `signing/make_cert.ps1`): open the built
   `.xlsm` → `Alt+F11` → **Tools** → **Digital Signature...** → **Choose** →
   **SDR DataViewer Templates** → **OK** → save the workbook → close Excel.

6. Final gate — re-verify **with the signature enforced**:

       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/verify_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.2.xlsm" --require-signature

   Any VBA edit inside the workbook strips the signature, so a failure here
   means the file was never signed or was modified after signing.

7. Open the rebuilt file and run the **operator acceptance checklist** (§ below).

8. Once accepted: rename the clean file to the canonical name (replacing the
   old one). Transfer to Synology only after the user approves — never automated.

   **After v1.2 ships, re-make any parallel-project copies from it.** Older
   copies carry stale macros; re-cloning each parallel-project template from
   the shipped v1.2 file brings them all onto the current behavior. Each
   tester runs `signing/tester-setup.ps1` **once per Windows account** so
   signed builds open with zero macro prompts.

**Rollback:** restore the timestamped `.bak-YYYYmmdd_HHMMSS` file the script
created alongside the source (and alongside the old `--out`, when `--force`
overwrote one).

---

## Manual fallback (always works — no pywin32, no Excel COM)

Use this if `build_clean_template.py` fails (COM trust issue, pywin32 missing,
etc.).

1. In Excel, open the old template. Create a new blank workbook.
2. Right-click each tab → Move or Copy → move to the new workbook ("Create a
   copy" checked), preserving order: `Test SOP's`, `Test Selection`, `_Settings`,
   the 12 canonical sheets, `_Template_Master`, `_Template_00`–`_Template_11`.
   (`_Settings` and `Test Selection` are normally hidden — unhide them first to
   copy, then re-hide `_Settings` as very-hidden afterwards.)
3. Delete the new workbook's original default sheet.
4. Alt+F11 → delete any old `DataViewerUpload`, `Module1`/`TestingTools`,
   `SampleNav` modules. File → Import File → import `DataViewerUpload.bas`,
   `TestingTools.bas`, `SampleNav.bas`. Double-click `ThisWorkbook` in the
   Project pane; replace its body with the contents of `ThisWorkbook.cls.txt`
   (keep the existing `Option Explicit` header or paste the whole file).
5. Open the file in the **Custom UI Editor** (free Office tool), paste the
   contents of `customUI14.xml` into the `customUI14` node, save.
6. Save the file as `.xlsm` (Excel Macro-Enabled Workbook). Close and reopen.
7. Run `verify_sidecar.py` against the result to confirm drift is zero.

---

## Operator acceptance checklist

Run this in Excel against the rebuilt (or manually migrated) file before
treating it as production-ready. Each item must pass before shipping.

1. **Test Selection layout.** The **Test Selection** sheet shows **only** the
   title (B2) + one-line hint (B3) + the 13-row checkbox table (B4:C16), in
   the documented order (Custom Test Template, Lifetime Test, Long Puff
   Lifetime Test, Rapid Puff Lifetime Test, Intense Test, User Test
   Simulation, Big Headspace Serial Test, Viscosity Compatibility, Various Oil
   Compatibility, Temperature Cycling Test #1, Temperature Cycling Test #2,
   Negative Pressure Test, Test SOP's) + the three-line guidance banner below
   the table (rows 18–20). No settings / status remain on the sheet.

2. **B4:B16 render as checkboxes.** Column B shows native in-cell checkboxes,
   not `TRUE`/`FALSE` text. (If they render as text, a COM value-write stripped
   the control — apply the **Option-B fallback** below.)

3. **Selection toggles visibility; SOP's still uploaded.** Ticking / unticking
   a test shows / hides its sheet. **Test SOP's** toggles its own sheet's
   visibility, but an **Upload All still includes Test SOP's even when its box
   is unticked**.

4. **Ribbon layout.** Order is Sample Blocks · Sample Navigation · **Help** ·
   **DataViewer Upload** · **Active Folders** — Help before DataViewer Upload.
   Upload All / **Upload Checkpoint** / Specify Test Name are text-only and
   stacked; the three Pick buttons are stacked; Delete All Review Sheets is its
   own column; **no group exceeds 3 rows** and the tab is not crowded.

5. **Active Folders updates live.** The **Active Folders** group shows the three
   current paths (full value on hover) and updates immediately after each Pick,
   including the new **Pick DataViewer File**.

6. **Instructions popup.** The **Instructions** button opens its MsgBox; no
   instructions remain on any sheet.

7. **Upload All — prompt + receipt + reset.** Upload All on ≥1 populated test →
   - Prompts for a descriptive file name (pre-filled with the last one; a name
     with no date gets today's date appended).
   - On success the **delivery receipt** popup appears (see item 14).
   - The `.xlsx` lands in Synology + Local + opens in DataViewer.
   - Each uploaded sheet has a `… - Review` copy kept; the canonical sheet is
     blank/fresh.
   - Selection resets to **Lifetime Test + Test SOP's**.

8. **Specify Test Name — rename + revert; Upload All restores.** Click
   **Specify Test Name**, enter a project name → the workbook on disk is renamed
   to `<project> - <original file name> (do not send).xlsm` (the clean rename
   removed the un-prefixed copy, so **only one file exists at a time**; an
   existing file at the target name gets an overwrite confirm, never a silent
   replace). Enter a **blank** name → the original file name is restored (the
   "(do not send)" marker disappears). Then **Upload All** → the original file
   name is restored automatically before the upload runs.

9. **Clog column is automatic.** On a populated test sheet, enter a **Draw
   Pressure (kPa)** value for a block → its **Clog** cell fills with the correct
   text + color, with no operator typing. Spot-check: Draw = `15` → **"Heavy
   Clog"** (red highlight, white text); a value in (5, 15) → **"Light Clog"**
   (yellow); blank or ≤ 5 → empty.

10. **Add/Remove Sample works structurally.** Add Sample / Remove Sample succeed
    on **every checkbox test sheet EXCEPT `Temperature Cycling Test #1` and
    `Test SOP's`** — including any renamed or copied test sheet (detection is by
    layout, not sheet name).

11. **Hidden plumbing + clean distributed copy.** `_Settings` is invisible to the
    operator. Open a distributed `.xlsx` → it contains **neither** `_Settings`
    **nor** `Test Selection`, and **no** `… - Review` sheets.

12. **All prior acceptance items still pass.** Re-confirm the pre-redesign
    behavior, none of which this change should disturb:
    - **No save prompt on close** (change nothing, close → no "Do you want to
      save?"; confirms the dirty-on-open save-loop fix).
    - **Navigation / sample edits:** Add Sample, Remove Sample, Reset Formulas,
      First/Prev/Next/Last, and Ctrl+Shift+. / Ctrl+Shift+, all work.
    - **Puffs step-picker crash safety:** type `20` in a puff seed cell → column
      fills; force an error mid-edit → events stay alive (toggles/pickers still
      respond; `EnableEvents` restored at `PuffDone:`).
    - **Review-copy accumulation:** uploading the same test 3–4× without deleting
      reviews appends `… - Review 2`, `… - Review 3`, … each at the end of the
      workbook (`UniqueReviewName` counter).
    - **Delete All Review Sheets** removes all `… - Review` sheets in one click;
      canonical / utility / template sheets untouched.
    - **Forensics (optional):** `verify_sidecar.py` against the live file shows no
      `xl/webextensions/*` parts and no `Workbook_SheetBeforeDoubleClick` handler.
    - **Reset safety — missing snapshot (optional):** on a scratch copy, remove
      one `_Template_NN` snapshot, put a sample on its matching test sheet, Upload
      All → that sheet is **left intact** (not reset, no Review copy) and the log
      notes "no internal snapshot"; nothing is lost.

13. **Upload Checkpoint — delivery without reset.** On a populated test, click
    **Upload Checkpoint** → same delivery as Upload All (Synology + Local +
    DataViewer) but the sheets are **NOT reset** and **no** `- Review` copies
    appear. Immediately run a second checkpoint with the **same** file name →
    it overwrites silently (no collision prompt: own-stream overwrite is the
    design). Then enter a name that collides with a **foreign** `.xlsx` already
    in the folder → an overwrite confirm appears (default **No**).

14. **Delivery receipt.** After any successful upload, the receipt popup states
    the data is **already delivered**, shows the full Synology path of the
    uploaded copy, says **"NEVER email or share THIS workbook"**, and offers
    **Open the Synology folder now?** → Yes opens Explorer with the file
    selected.

15. **Receiver test — distributed copy is ribbon-free.** Open the distributed
    `.xlsx` (ideally on a machine/profile without the template): **no
    "TPM Testing" tab** appears and **no "Cannot run the macro" popups** fire on
    open, on tab hover, or anywhere else. A tester-created **chart sheet** in
    the workbook must also be absent from the `.xlsx`.

16. **Lockbox restore.** (a) Delete a canonical test sheet → re-tick its box on
    Test Selection → a **fresh blank sheet** is restored from its snapshot.
    (b) **Rename** a populated canonical sheet, run Upload All → a warning lists
    the orphan sheet ("looks like a test sheet… will NOT be uploaded"); after
    the upload the canonical sheet exists again, fresh.

17. **Puffs v1.2 behavior.** Type any non-list number (e.g. `7`) into a puffs
    cell at row 7 → rows below fill with `prev+7`. Type `25` into **row 5** →
    row 5 holds `25` and rows below fill with `prev+25`. Pick `custom` → the
    column clears; **paste** an irregular sequence → it survives and uploads
    unchanged (typing single numbers re-triggers the auto-fill — paste is the
    literal-entry path). Type a number into a **`- Review`** sheet's puffs
    column → nothing auto-fills (archives are immune). Clearing a puffs cell or
    typing text must NOT raise a VBA error.

18. **Clog text-safety.** Enter `n/a` (text) as Draw Pressure → the Clog cell
    stays **empty** (not "Heavy Clog").

19. **Macros-disabled banner.** Open a copy without enabling macros → the
    Test Selection sheet shows the three-line banner (enable-macros hint,
    "never email or upload this file", Upload All/Checkpoint deliver
    automatically) and it is legible.

20. **Signing + close nudge.** On a machine that ran
    `signing/tester-setup.ps1`, the signed build opens with **zero macro
    prompts** (even a copy that arrived by email/Teams); on a machine that
    didn't, the standard prompt appears (not the red block, for local files).
    Close the workbook with un-uploaded data → a **one-button reminder**
    appears (close is never blocked); close again after an upload → no
    reminder.

---

## Option-B fallback — if B4:B16 shows `TRUE`/`FALSE` text instead of checkboxes

Native in-cell checkboxes **cannot be created via COM** (pywin32/Excel
automation) — only via the Office.js `range.control = {type:"Checkbox"}` API
(ExcelApi 1.18+) or the desktop ribbon's Insert ▸ Checkbox. The build therefore
**preserves** the source workbook's existing `B4:B16` checkbox cells and only
*value-writes* into them when re-laying the table. The risk (caught here at
acceptance item 2): if a COM value-write is ever found to strip the native
control, the rebuilt sheet shows plain `TRUE`/`FALSE` text.

Recovery — **describe-only here; build it only if acceptance item 2 fails:**

1. Open the rebuilt `…v1.2.xlsm`. Apply native checkboxes to `B4:B16`
   either via **Excel on the web** with an Office Script
   (`range.control = {type:"Checkbox"}`), or via the desktop ribbon
   **Insert ▸ Checkbox** over the selected range.
2. Snapshot that now-correct sheet once (e.g. a one-time `SeedSelectionTemplate`
   macro) into a very-hidden `_Template_Sel` template, the same way the
   `_Template_NN` data snapshots are baked.
3. Switch `_relay_test_selection` in `build_clean_template.py` to **stamp the
   `_Template_Sel` snapshot** (full `Worksheet.Copy` from the baked template)
   instead of value-writing into the live cells, so every rebuild reproduces the
   real checkboxes.

---

## Re-seeding blank-template snapshots

If a sheet's layout changes: update the blank `Automated Testing Template -
DVE.xlsm` to match, then run `SeedBlankTemplatesFromFile` (Alt+F8) to refresh
the `_Template_NN` snapshots. On a fully blank workbook you can instead use
`RebuildBlankTemplates`.

---

## What changes / what doesn't

- **Changed:** VBA modules; ribbon (5-group TPM Testing tab: Sample Blocks ·
  Sample Navigation · Help · DataViewer Upload · Active Folders); the upload
  sheet is renamed `DataViewer Upload` → **Test Selection** and stripped to just
  the 13-row checkbox table; paths / file name / status / log moved onto a new
  very-hidden **`_Settings`** sheet (the six named ranges move with them); file
  name is now an InputBox prompt at Upload, results shown via MsgBox; on-sheet
  upload buttons removed (ribbon hosts them now); post-upload reset is
  non-destructive (Review copy instead of delete+recreate); `_Template_NN`
  snapshots rebuilt as blank.
- **Unchanged:** every visible data sheet and all your data. Test SOP's data
  untouched (now a visibility toggle, still always uploaded). Named ranges
  (`DV_FileName`, `DV_SynologyPath`, `DV_LocalPath`, `DV_Status`, `DV_Log`,
  `DV_TestSelection`, `DV_DataViewerExe`) preserved — six now resolve onto
  `_Settings`, `DV_TestSelection` onto `Test Selection`.

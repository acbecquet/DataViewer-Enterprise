# Runbook — migrate an existing Automated Testing Template

Goal: replace the macros + ribbon in a live operator workbook with the
canonical sources from this folder, without touching any data sheets.

---

## Recommended path: clean rebuild with `build_clean_template.py`

`build_clean_template.py` produces a **brand-new** `.xlsm` (fresh OOXML
package — no inherited corruption, no web add-in) by copying the source's
sheets into a new workbook, importing the canonical VBA, and injecting the
ribbon at the zip level. Run on the work machine (Excel + pywin32 required).

**Prerequisites:**

- `pip install pywin32` (one-time).
- Excel Trust Center → Macro Settings → **Trust access to the VBA project
  object model** = ENABLED.

**Steps:**

1. Close the Automated Testing Template in Excel.

2. Run the headless gates (no Excel needed — confirm all-PASS before building):

       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/check_sources.py
       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/test_build_helpers.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm"

3. Build the clean workbook (backs up source automatically):

       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/build_clean_template.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm" --out "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1 (clean).xlsm"

4. Drift-verify the output (every line should be MATCH / OK):

       "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/verify_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1 (clean).xlsm"

   Expected: all modules `MATCH`, `customUI14.xml == repo`,
   `no web-extension/add-in parts`, `RESULT: all modules match`.

5. Open the rebuilt file and run the **operator acceptance checklist** (§ below).

6. Once accepted: rename the clean file to the canonical name (replacing the
   old one). Transfer to Synology only after the user approves — never automated.

**Rollback:** restore the `.bak` file the script created alongside the source.

---

## Manual fallback (always works — no pywin32, no Excel COM)

Use this if `build_clean_template.py` fails (COM trust issue, pywin32 missing,
etc.).

1. In Excel, open the old template. Create a new blank workbook.
2. Right-click each tab → Move or Copy → move to the new workbook ("Create a
   copy" checked), preserving order: `Test SOP's`, `DataViewer Upload`, the 12
   canonical sheets, `_Template_Master`, `_Template_00`–`_Template_11`.
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

1. **No save prompt on close.** Open the rebuilt file, change nothing, close →
   Excel must not prompt "Do you want to save?". (Confirms the dirty-on-open
   save-loop fix.)

2. **Sheet visibility and navigation.** Toggle a test TRUE/FALSE on the
   DataViewer Upload sheet → the matching sheet shows/hides correctly. Verify:
   Add Sample, Remove Sample, Reset Formulas, First/Prev/Next/Last sample
   navigation, and Ctrl+Shift+. / Ctrl+Shift+, all work.

3. **Puffs step-picker — crash safety.** Type `20` in a puff seed cell → the
   column fills. Force an error mid-edit (e.g. type into a protected cell) →
   after the error, events are still alive: toggles and pickers still respond.
   (Confirms `EnableEvents` is always restored at `PuffDone:`.)

4. **Dry-Run Checklist.** Click Dry-Run Checklist on the ribbon (or the button)
   on a populated sheet → checklist passes with no errors.

5. **Upload All — Review copy created.** Upload All on ≥1 populated test →
   - The `.xlsx` lands in Synology + Local + opens in DataViewer.
   - Each uploaded sheet now has a `… - Review` copy at the end of the
     workbook; the canonical sheet is blank/fresh.
   - Selection resets to Lifetime-only.

6. **Review copy accumulation — no distributed copies.** Upload a second test
   without deleting existing reviews → a second Review sheet appears. Then
   upload the **same** test 3–4 times in a row (still without deleting reviews)
   → each run adds `… - Review 2`, `… - Review 3`, … with a distinct counter
   suffix, each landing at the end of the workbook. (This exercises the
   `UniqueReviewName` counter + end-positioning that the headless checks can't
   confirm.) Open a distributed `.xlsx` → it contains **no** Review sheets.

7. **Delete All Review Sheets.** Click Delete All Review Sheets on the ribbon →
   all `… - Review` sheets are gone in one click; canonical, utility, and
   template sheets are untouched.

8. **Forensics (optional).** Re-run the forensic dump or `verify_sidecar.py`
   against the live file → no `xl/webextensions/*` parts, no
   `Workbook_SheetBeforeDoubleClick` handler.

9. **Reset safety — missing snapshot (optional).** To confirm the "never clear
   without a known-good source" poka-yoke: on a scratch copy, delete one
   `_Template_NN` snapshot (VBE → set its `.Visible` then delete, or skip
   seeding one), put a sample on its matching selected test sheet, and Upload
   All → that sheet is **left intact** (not reset, no Review copy) and the log
   notes "no internal snapshot"; nothing is lost.

---

## Re-seeding blank-template snapshots

If a sheet's layout changes: update the blank `Automated Testing Template -
DVE.xlsm` to match, then run `SeedBlankTemplatesFromFile` (Alt+F8) to refresh
the `_Template_NN` snapshots. On a fully blank workbook you can instead use
`RebuildBlankTemplates`.

---

## What changes / what doesn't

- **Changed:** VBA modules; ribbon (3-group TPM Testing tab); on-sheet upload
  buttons removed (ribbon hosts them now); post-upload reset is now
  non-destructive (Review copy instead of delete+recreate); `_Template_NN`
  snapshots rebuilt as blank.
- **Unchanged:** every visible data sheet and all your data. Test SOP's
  untouched. Named ranges (`DV_FileName`, `DV_SynologyPath`, `DV_LocalPath`,
  `DV_Status`, `DV_Log`, `DV_TestSelection`, `DV_DataViewerExe`) preserved.

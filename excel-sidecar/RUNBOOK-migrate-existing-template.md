# Runbook -- migrate an existing Automated Testing Template

Goal: make the new post-upload reset work in a workbook that **already has real
data**, without changing any of your data sheets. Only the macros are updated and
12 very-hidden `_Template_NN` blank-snapshot sheets are added. Your visible
sheets and data (including **Lifetime Test**) are left exactly as they are.

## Prerequisites

- `pywin32` installed (for `install_sidecar.py --apply`).
- Excel Trust Center -> Macro Settings -> **Trust access to the VBA project
  object model** = ENABLED.
- A **blank** source workbook that has every test sheet -- e.g.
  `resources/templates/Automated Testing Template - DVE.xlsm` (verified clean).

## Steps

1. **Close** the Automated Testing Template in Excel.
   (`install_sidecar.py` drives Excel via COM and needs the file closed.)

2. **Update the macros** (auto-backs-up first; never touches sheets/data):

   ```bat
   cd "C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar"
   py -3.13 install_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template.xlsm" --apply
   ```

   This writes `Automated Testing Template.xlsm.bak-<timestamp>`, removes the old
   modules, imports the updated ones (incl. `RebuildBlankTemplates` and
   `SeedBlankTemplatesFromFile`), and sets the `ThisWorkbook` code.

3. **Reopen** the Automated Testing Template (enable macros if prompted).

4. **Seed the snapshots** from the blank copy:
   - Press **Alt+F8**, choose **`SeedBlankTemplatesFromFile`**, click Run.
   - In the file picker, choose **`Automated Testing Template - DVE.xlsm`**.
   - It reports *"Created 12 blank-template snapshot(s)... Your data sheets were
     not touched."*

5. **Save** the workbook (Ctrl+S).

6. **Verify** (optional):

   ```bat
   py -3.13 install_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template.xlsm"
   py -3.13 verify_sidecar.py  --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template.xlsm"
   ```

   `verify_sidecar.py` should report every module `MATCH`.

## What changed / what didn't

- Changed: the VBA modules; 12 very-hidden `_Template_NN` snapshot sheets added.
- Unchanged: every visible sheet and all your data. **Lifetime Test is not
  touched.** (The snapshots come from the blank DVE copy, not from your sheets.)

## How the reset behaves now

On **Upload All**, each uploaded sheet is reverted to its `_Template_NN` snapshot
(blank, at the DVE copy's block count). Re-grow a sheet with Add Sample for the
next campaign. The snapshots are auto-hidden and auto-excluded from the
distributed `.xlsx` copies.

## Re-seeding later

If you change a sheet's layout: update the blank `Automated Testing Template -
DVE.xlsm` to match, then re-run `SeedBlankTemplatesFromFile`. (On a fully blank
workbook you can instead use `RebuildBlankTemplates`, which snapshots the
workbook's own sheets.)

## Rollback

Close Excel and restore `Automated Testing Template.xlsm.bak-<timestamp>` over
the file.

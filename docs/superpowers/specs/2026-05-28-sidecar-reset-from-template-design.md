# Sidecar Post-Upload Reset -- Revert From Internal Blank Snapshots (Design)

**Date:** 2026-05-28
**Branch:** `feature/xlsm-vba-sidecar` (sidecar worktree)
**Status:** Approved + implemented. FINAL design. Supersedes the surgical-clear
approach (`2026-05-27-sidecar-post-upload-reset.md`) and the external-packaged-
template approach that was the first cut of this doc.

## Problem

`Btn_UploadAll` resets the live operator workbook after a successful upload so it
is fresh for the next session. The surgical per-cell clear (`ClearSheetEntries`)
wiped the `A5` puff seed (the "row 5 keeps getting deleted" bug), could not know
per-sheet seeds/intervals/notes, and skipped non-standard sheets. It needed a
blank *source* to revert to.

## Why internal snapshots (not an external template file)

The first cut opened an external packaged template (`Standardized Test
Template*.xlsx` next to `DataViewer.exe`) and copied sheets over. That coupled
the reset to a second file that drifted out of date with the operator template --
exactly the failure that surfaced in testing -- and to the DataViewer install
path. (DataViewer's own C++ also loads that packaged template for other features,
so it keeps its own copy regardless.)

Instead, each canonical test sheet gets a **blank snapshot stored inside the
workbook itself** -- a very-hidden `_Template_NN` sheet. The reset copies the
snapshot over the live sheet. Self-contained: no external file, no install-path
dependency, no drift. This is what the original deployed code *tried* to do (it
referenced `_Template_<X>` sheets that were never created).

## Snapshot storage

- One very-hidden sheet per canonical data sheet, named `_Template_NN` where `NN`
  is the zero-padded index into `CanonicalDataSheets()` (`_Template_00` ..
  `_Template_11`). Positional rather than name-based to stay within Excel's
  31-char sheet-name limit and dodge special-char issues.
- The `_Template_` prefix is load-bearing: `ApplySheetVisibility` already
  auto-very-hides any `_Template_*` sheet, and `TrimSheetsInWorkbook` already
  drops every non-kept sheet -- so snapshots are auto-hidden and auto-excluded
  from the distributed `.xlsx` copies with no extra code.
- `TemplateSheetName(idx)` centralizes the naming.

## Building snapshots: `RebuildBlankTemplates`

A `Public Sub` the operator runs (Alt+F8) once at setup and again whenever a
sheet's layout changes:
- **Poka-yoke:** refuses (MsgBox) if any canonical sheet has entered samples, so
  real data can never be baked into a template. Run it on a BLANK workbook.
- For each canonical sheet: drop any existing `_Template_NN`, copy the (blank)
  sheet via Excel's native `Worksheet.Copy` (preserves formatting, validation,
  conditional formatting, merges -- unlike openpyxl), rename to `_Template_NN`.
- Hidden sources (most are, per `DV_TestSelection`) are made visible for the copy
  so the copy becomes the active sheet, then restored.
- Snapshots are hidden at the end, after activating a non-snapshot sheet (Excel
  refuses to very-hide the active sheet).

## Reset flow

`ResetLiveWorkbookAfterUpload(keep)` -- no external file open:
- For each canonical sheet (by index `i`) in `keep` that exists live, call
  `RestoreSheetFromTemplate ThisWorkbook, sheetName, i`.
- Then `ResetSelectionToDefault` + `ApplySheetVisibility` (re-hide; default
  selection = Lifetime Test).

`RestoreSheetFromTemplate(liveWb, sheetName, idx)` -- per-sheet, transactional:
- Look up `_Template_NN`. If missing -> log + skip (never clear without a source).
- **A1 sanity:** if the snapshot's title and the live sheet's title both exist
  and differ, the positional mapping is stale -> log + skip (don't restore the
  wrong template).
- Make the very-hidden snapshot visible, copy it Before the live sheet (copy
  becomes active), re-hide the snapshot, delete the original, rename the copy.
  The original is deleted only AFTER the replacement is in place -> no data-loss
  window.

## Removed / superseded

- External-file machinery: `ResolveTemplatePath`, the `DV_TemplatePath` override,
  opening a template workbook. Gone.
- Surgical clear: `ClearSheetEntries`, `BlockCount`, `LAST_DATA_ROW`. Gone.

## Behavior

- Reset sheets return to the snapshot's block count (the blank default); grow
  again with Add Sample. (Per the user's "full revert" choice.)
- Every canonical sheet can be reset, including `Custom Test Template` and the
  non-standard `Temperature Cycling Test #1` (the external approach could not
  reset `Custom Test Template` -- it had no packaged counterpart).
- Distributed `.xlsx` copies never contain the snapshots (trimmed out).

## First-time setup

1. Open a BLANK template workbook (e.g. the blank `Automated Testing Template -
   DVE.xlsm` copy).
2. Import the updated macros (`DataViewerUpload.bas`, etc.).
3. Run `RebuildBlankTemplates` once. The workbook is now self-contained.
4. Use it as the operator template going forward. No external template file is
   needed at runtime.

## Verification

- **Structural (here):** Sub/Function balance; `ResolveTemplatePath` /
  `DV_TemplatePath` / `ClearSheetEntries` / `BlockCount` / `LAST_DATA_ROW`
  absent; `TemplateSheetName` / `RebuildBlankTemplates` present; CRLF + plaintext.
- **Manual (Excel):** run `RebuildBlankTemplates` on a blank workbook (confirm it
  creates snapshots and refuses on a data-laden one); enter data incl. an
  expanded sheet + Temp Cycling #1; Upload All; confirm each sheet reverts to its
  blank snapshot (block count default, `A5` seed + formulas intact, milestone
  notes present); confirm distributed `.xlsx` copies contain no `_Template_`
  sheets; confirm selection reset + re-hide.

## Constraints honored

Commits allowed on this branch (per the user). Tooling never modifies the live
operator `.xlsm` (the operator runs `RebuildBlankTemplates` themselves). No
`src/` / `.pro` / Qt changes. `%USERPROFILE%\SynologyDrive\` untouched.

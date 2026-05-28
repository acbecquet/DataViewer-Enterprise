# Sidecar Post-Upload Reset -- Revert From Packaged Template (Design)

**Date:** 2026-05-28
**Branch:** `feature/xlsm-vba-sidecar` (sidecar worktree)
**Status:** Approved design (brainstorming output). Supersedes the surgical-clear
approach in `2026-05-27-sidecar-post-upload-reset.md`. Implementation plan to
follow via writing-plans.

## Problem

`Btn_UploadAll` resets the live operator workbook after a successful upload so
it is fresh for the next session. The current reset, `ClearSheetEntries` (a
surgical per-cell clear added 2026-05-27), has three problems:

1. **It wipes row 5.** It clears `A5` (the per-sheet puff *seed*) plus `B5`/`C5`.
   The puff chain is `A6 = A5 + interval`, so blanking `A5` breaks the whole puff
   sequence -- the "row 5 keeps getting deleted" symptom the user reported.
2. **It cannot know per-sheet truth.** Seeds, puff intervals, the `B5`
   before-weight seed, and intentional milestone notes differ per sheet type; a
   hand-maintained cell map will always risk drifting from the real template.
3. **It skips non-standard sheets.** `Temperature Cycling Test #1` has no row-4
   headers (`BlockCount = 0`), so `ClearSheetEntries` does `Exit Sub` and never
   resets it.

## Decision

Replace the surgical clear with a **whole-sheet revert from the
DataViewer-packaged template**. After upload, each uploaded data sheet becomes an
exact copy of that sheet in the packaged template -- correct headers, seeds,
intervals, formulas, formatting, data-validation, and milestone notes -- at the
template's default sample-block count. The operator re-grows blocks with
AddSample for the next campaign, exactly as when starting from a fresh template.

Confirmed with the user: **full revert, including resetting block count to the
template default.**

## Template source

`resources\templates\Standardized Test Template - December 2025.xlsx`, shipped
inside the DataViewer install. Verified against the live operator workbook:

- Contains **every** `CanonicalDataSheets()` entry **except `Custom Test
  Template`** (operator-only). `Test SOP's` is not a canonical data sheet and is
  never reset.
- Matching `A1` titles and identical row-4 headers on all shared sheets.
- Correct per-sheet puff seed + interval baked into rows 5-6 (Lifetime `A5=5`,
  `A6=A5+5`; Viscosity `A5=20`; Negative Pressure `A5=1`) plus the I/J/K/L
  formula scaffolding, and intentional template content (e.g. Negative Pressure
  milestone notes at `H5/H15/H25/H35`).

### Locating it at runtime

`templatesDir = parentFolder( ResolveDataViewerExe() ) & "\resources\templates\"`,
then the `Standardized Test Template*.xlsx` inside it (newest by modified date if
several -- survives the annual filename roll to "...2026"). `ResolveDataViewerExe()`
already reads the `DV_DataViewerExe` named range (cell `I12`); on this machine
that is `C:\Users\S1134987\DataViewer Enterprise\DataViewer.exe`, giving
`C:\Users\S1134987\DataViewer Enterprise\resources\templates\` (confirmed to
exist, holding the one template file). Optional `DV_TemplatePath` named range
overrides the whole lookup.

## Why whole-sheet revert is safe

No defined name or cross-sheet formula references any data sheet -- every `DV_*`
name lives on the `DataViewer Upload` sheet, and a workbook-wide scan found zero
cross-sheet formula references into the data sheets. So deleting and replacing a
data sheet cannot create `#REF!`s or break a named range.

## Approach

Rewrite `ResetLiveWorkbookAfterUpload(keep)`:

1. Resolve the template path. **If not found -> log and `Exit Sub`, leaving all
   data intact.** (Never clear without a source.)
2. Open the template **read-only, once** (`UpdateLinks:=0`).
3. For each sheet in `CanonicalDataSheets()` that is in `keep` and exists live,
   call `RestoreSheetFromTemplate liveWb, tplWb, name`.
4. Close the template (`SaveChanges:=False`).
5. `ResetSelectionToDefault`; `ApplySheetVisibility` (re-hide; default selection
   = Lifetime Test).

`RestoreSheetFromTemplate(liveWb, tplWb, name)` -- per-sheet, transactional:

1. If the template lacks `name` -> log "no packaged template -- not auto-reset"
   and `Exit Sub` (this covers `Custom Test Template`).
2. `Set orig = liveWb.Worksheets(name)`; remember its tab index.
3. Copy the template sheet **Before** `orig` (the copy lands at the original tab
   position and becomes the active sheet -- capture it by object reference, not
   by name, to be robust against any stale "name (2)" leftovers).
4. Delete `orig` (`DisplayAlerts=False` around the delete).
5. Rename the copy to `name`.

The original is deleted only **after** its replacement is in place, so a
mid-failure can never leave the sheet missing -- worst case a temporarily
double-named sheet, never data loss. `Application.EnableEvents` and
`ScreenUpdating` are off for the whole reset and restored in `Cleanup`.

## Removals

- `ClearSheetEntries`, `BlockCount`, and the `LAST_DATA_ROW` constant (all exist
  only for the surgical clear). Keep `FIRST_DATA_ROW` (still used by validation
  and trim).
- Update the module header comment (step "e") to describe the template revert
  instead of the hard-clear wording.

## Edge cases / fail-safes (poka-yokes)

- Template folder/file missing or unopenable -> skip the whole reset + log; data
  intact.
- `Custom Test Template`, or any other sheet absent from the template -> skip
  that sheet + log; its data is left intact.
- Reset only touches **uploaded** sheets (the `keep` list), never sheets the
  operator did not select.

## Out of scope (works -- do not touch)

Save / folder-copy / DB / DataViewer-launch path (confirmed good by the user and
the upload log). The duplicate-sample-ID hint (already shipped this session).
Any alternative handling for `Custom Test Template` beyond skip-and-log.

## Verification

- **Structural (here):** Sub/End Sub + Function/End Function balanced; no
  remaining references to `ClearSheetEntries` / `BlockCount` / `LAST_DATA_ROW`;
  CRLF + plaintext; `verify_sidecar.py` parses the module.
- **Manual (Excel, by the user):** re-import `DataViewerUpload.bas`; upload a
  selection that includes an expanded sheet (e.g. Lifetime grown to 12 blocks)
  and `Temperature Cycling Test #1`; confirm each reverts to the template
  (default block count, `A5` seed present, formulas + milestone notes intact);
  confirm `Custom Test Template` (if uploaded) is left intact + logged; confirm
  the selection resets to Lifetime and everything re-hides except Lifetime /
  Test SOP's / DataViewer Upload. Fail-safe check: set `DV_TemplatePath` to a
  missing file -> reset skips + logs, data intact.

## Constraints honored

NO commits on this branch (this design doc is left uncommitted for review, per
the standing instruction -- which overrides the brainstorming skill's
commit-the-doc step). Live operator `.xlsm` read-only (inspection only).
`%USERPROFILE%\SynologyDrive\` untouched. No `src/` / `.pro` / Qt changes.

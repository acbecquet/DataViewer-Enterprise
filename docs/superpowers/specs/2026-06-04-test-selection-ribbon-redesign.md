# Test Selection + Ribbon Redesign — Design Spec

**Date:** 2026-06-04
**Branch:** `feature/excel-sidecar-rebuild` (continuation; builds on the clean-rebuild work already shipped to the live template)
**Status:** Approved (design); pending implementation plan
**Supersedes/extends:** `2026-06-04-excel-sidecar-clean-rebuild.md` (same sidecar; this is the next iteration)

---

## Goal

Reduce the operator's upload sheet to a single, well-formatted **checkbox table**
("Test Selection"), and move every other affordance — folder/exe paths, the file
name, status/log, and the instructions — off the sheet and onto the **TPM Testing
ribbon** (or into popups). The ribbon becomes the single home for all actions; the
sheet becomes pure test selection.

## Background — current state (post clean-rebuild)

- One sheet named **"DataViewer Upload"** hosts: a TRUE/FALSE checkbox table
  (`DV_TestSelection`, `$A$3:$B$16`), an instructions block (6 numbered steps),
  an "Upload Settings" block (File Name, Synology/Local folders, DataViewer.exe),
  and a "Last Run Status" + "Last Log" area.
- Six named ranges live **on that sheet**: `DV_FileName=$I$6`,
  `DV_SynologyPath=$I$8`, `DV_LocalPath=$I$10`, `DV_DataViewerExe=$I$12`,
  `DV_Status=$H$16`, `DV_Log=$H$17`. The selection range `DV_TestSelection` is
  `$A$3:$B$16` (col A = TRUE/FALSE, col B = sheet name).
- The selection table **already contains a "Test SOP's" row**, but it is inert:
  `ApplySheetVisibility` force-shows `Test SOP's` regardless, and `BuildKeepList`
  always uploads it.
- Ribbon group **"DataViewer Upload"**: Upload All (large), Dry-Run Checklist
  (large), Pick Synology Folder (normal), Pick Local Folder (normal), Delete All
  Review Sheets (large).
- VBA: `OnWorkbookSheetChange`, `ApplySheetVisibility`, `BuildKeepList`,
  `RunChecklist`, the `Btn_*` handlers, and the `Ribbon_*` wrappers all key off the
  sheet-name constant `UPLOAD_SHEET_NAME = "DataViewer Upload"` and the named
  ranges via `GetNamed`/`SetNamed`.
- Build: `build_clean_template.py` hard-codes `UPLOAD_SHEET`, the `KEEP` list, the
  `NAMED` dict (name → cell on the upload sheet), and `_set_default_visibility`'s
  visible set.

## Scope

**In scope**

1. Rename the sheet to **"Test Selection"** and strip it to only the checkbox table.
2. Reorder the checkbox table to the operator's order (13 rows; see §2.2).
3. Make **Test SOP's** a real visibility toggle (still always uploaded).
4. Relocate the six non-selection named ranges to a new very-hidden **`_Settings`**
   sheet (values carried over).
5. Ribbon redesign: 2-column upload/pick layout, **Delete All Review Sheets** as its
   own column, a new **Pick DataViewer File** button, a new **Active Folders**
   read-only path display, and a new **Instructions** button.
6. File name entered via **InputBox at upload**; results surfaced via **MsgBox**;
   full log retained on `_Settings`.
7. Extend the headless gates + build + verify accordingly.

**Out of scope / non-goals**

- No change to the data-sheet (12-col block) layout, the upload/trim/reset
  pipeline semantics, the snapshot (`_Template_NN`) mechanism, or the Review-sheet
  feature.
- No change to `CanonicalDataSheets()` membership or order (and therefore no
  snapshot re-indexing). The on-sheet table order is **decoupled** from
  `CanonicalDataSheets()` (visibility/keep-list look tests up by name).
- No Synology interaction; no auto-merge. Same ship pipeline as before.

---

## Modification 1 — "DataViewer Upload" → "Test Selection"

### 1.1 Sheet rename + content strip

- Rename `"DataViewer Upload"` → `"Test Selection"`. One VBA constant
  (`UPLOAD_SHEET_NAME`) and the build script's `UPLOAD_SHEET`/`KEEP`/`NAMED`/
  visibility set change; everything else keys off the constant.
- Remove from the sheet: the instructions block, the Upload Settings block (File
  Name + the three paths), and the Status/Last-Log area. The sheet retains **only**
  the title, a one-line hint, and the checkbox table.

### 1.2 Checkbox table (new order — 13 rows)

Layout on the **Test Selection** sheet:

| Row | Col A (checkbox / bool) | Col B (test name) |
|----:|--------------------------|-------------------|
| 1   | *title*: **TEST SELECTION** (merged A1:B1) | |
| 2   | *hint*: "Check the tests you're running." (merged A2:B2) | |
| 3   | ☐ | Custom Test Template |
| 4   | ☑ | Lifetime Test  *(default TRUE)* |
| 5   | ☐ | Long Puff Lifetime Test |
| 6   | ☐ | Rapid Puff Lifetime Test |
| 7   | ☐ | Intense Test |
| 8   | ☐ | User Test Simulation |
| 9   | ☐ | Big Headspace Serial Test |
| 10  | ☐ | Viscosity Compatibility |
| 11  | ☐ | Various Oil Compatibility |
| 12  | ☐ | Temperature Cycling Test #1 |
| 13  | ☐ | Temperature Cycling Test #2 |
| 14  | ☐ | Negative Pressure Test |
| 15  | ☑ | Test SOP's  *(default TRUE)* |

- `DV_TestSelection = 'Test Selection'!$A$3:$B$15` (13 rows).
- **Names in col B must use the canonical sheet spellings** so the dict lookup
  matches: `Negative Pressure Test` (not "Negative pressure Test"),
  `Temperature Cycling Test #1` / `#2`, `Test SOP's` (straight apostrophe;
  `NormalizeSheetName` already folds curly↔straight).
- Col A holds a real boolean (TRUE/FALSE) rendered as a checkbox. Defaults:
  **Lifetime Test** and **Test SOP's** TRUE; all others FALSE.
- Formatting: title row styled (bold, fill), hint row muted, table with light
  borders/banding, column A narrow (checkbox), column B wide enough for the
  longest name. Exact styling is cosmetic; the plan specifies concrete values.

### 1.3 Ribbon layout (TPM Testing tab)

**Hard constraint: every group is at most 3 rows tall.** Five groups, left→right:

1. **Sample Blocks** — unchanged (Add / Remove / Reset Formulas).
2. **Sample Navigation** — unchanged (First / Prev / Next / Last).
3. **Help** — **Instructions** button, `size="large"`, info icon
   (`imageMso="Info"`) (§ Modification 2). Positioned **immediately before**
   DataViewer Upload.
4. **DataViewer Upload** — three columns:
   - Col 1 (vertical box): **Upload All**, **Dry-Run Checklist** — `size="normal"`,
     **no image** (text-only), stacked.
   - Col 2 (vertical box): **Pick Synology Folder**, **Pick Local Folder**,
     **Pick DataViewer File** — `size="normal"`, small `imageMso` icons, stacked.
   - Col 3: **Delete All Review Sheets** — `size="large"` (its own column).
5. **Active Folders** — three read-only path rows (§1.4): Synology, Local,
   DataViewer (exactly 3 rows).

`<customUI>` gains `onLoad="Ribbon_OnLoad"` to capture the `IRibbonUI` for dynamic
refresh of the path rows.

### 1.4 Active Folders — read-only path display

Three controls, one per path, showing the **full** stored value with the complete
path available on hover; they refresh immediately after any Pick.

- **Primary control: read-only `editBox`** — fixed width (`sizeString`), `getText`
  returns the full path, `getSupertip` returns the full path, and `onChange`
  re-invalidates the control (reverting any keystroke ⇒ effectively read-only).
  An editBox bounds its width and clips/scrolls long text, which is what makes long
  paths usable in the ribbon.
- **Decided: read-only `editBox`** (see §"Resolved mechanics"). `labelControl`
  remains a possible fallback if the editBox proves awkward in acceptance.
- Rows: `Synology:`, `Local:`, `DataViewer:`.

### 1.5 New "Pick DataViewer File" button

Adds a file picker for the DataViewer executable, writing `DV_DataViewerExe`.
New handler `Btn_PickDataViewerExe` + a `PickFileInto` helper
(`Application.GetOpenFilename`, filter `*.exe`). Wrapper `Ribbon_PickDataViewer`.

---

## Modification 2 — Instructions button

A **large Instructions button** (info icon) on the ribbon shows a `MsgBox` with
rewritten, ribbon-centric guidance. No instructions remain on any sheet. The text:

```
How to upload test data to DataViewer

1. Test Selection sheet: tick the tests you're running. Each ticked test's
   sheet appears — fill it in. (Untick to hide a sheet again.)

2. Set your destinations once, from the TPM Testing ribbon:
      Pick Synology Folder   •   Pick Local Folder   •   Pick DataViewer File
   The current paths show in the 'Active Folders' box on the ribbon.

3. (Optional) Dry-Run Checklist validates your data without uploading.

4. Upload All: enter a descriptive file name when prompted (Product + Test + Date).
   The data is copied to the Synology and Local folders, opened in DataViewer,
   and each uploaded sheet is reset — a '<name> - Review' copy is kept so you can
   see what was sent.

5. A summary pops up when it finishes. Use 'Delete All Review Sheets' to clear
   the review copies once you're done with them.

Tip: 'Test SOP's' is always included in every upload, whether or not its box is ticked.
```

Handler `Btn_ShowInstructions`; wrapper `Ribbon_Instructions`; the text lives in a
single VBA constant.

---

## Data model changes

### New very-hidden `_Settings` sheet

Holds the six relocated cells; their **named ranges move with them**, so all VBA
keeps working unchanged through `GetNamed`/`SetNamed`.

| `_Settings` cell | Label (col A) | Named range (col B) |
|------------------|---------------|---------------------|
| `B1` | File name (last used) | `DV_FileName` |
| `B2` | Synology folder | `DV_SynologyPath` |
| `B3` | Local folder | `DV_LocalPath` |
| `B4` | DataViewer.exe | `DV_DataViewerExe` |
| `B5` | Status | `DV_Status` |
| `B6` | Log | `DV_Log` |

- Visibility: **`xlSheetVeryHidden`**. Added to the very-hidden enforcement in both
  VBA (`ApplySheetVisibility`) and the build (`_set_default_visibility`), since
  `_Settings` matches neither the `_Template_` nor `_Macro` prefix.
- **Excluded from distributed copies**: `_Settings` is not in the trim keep-list, so
  `TrimSheetsInWorkbook` drops it from the macro-free `.xlsx`. (As does the Test
  Selection sheet — neither is needed in a distributed copy.)
- **Values carried over** during the build: existing `DV_FileName` /
  `DV_SynologyPath` / `DV_LocalPath` / `DV_DataViewerExe` values are read from the
  source workbook before restructuring and written into `_Settings`, so the
  operator does not re-pick folders after a rebuild.

### Named-range targets after the change

| Named range | New RefersTo |
|-------------|--------------|
| `DV_TestSelection` | `'Test Selection'!$A$3:$B$15` |
| `DV_FileName` | `'_Settings'!$B$1` |
| `DV_SynologyPath` | `'_Settings'!$B$2` |
| `DV_LocalPath` | `'_Settings'!$B$3` |
| `DV_DataViewerExe` | `'_Settings'!$B$4` |
| `DV_Status` | `'_Settings'!$B$5` |
| `DV_Log` | `'_Settings'!$B$6` |

---

## Behavior changes (VBA — `DataViewerUpload.bas`, `ThisWorkbook.cls.txt`)

### B1. Sheet rename

`UPLOAD_SHEET_NAME = "Test Selection"`. All references (visibility, change handler,
delete-reviews active-sheet guard, snapshot-copy anchor, rebuild activate) follow
the constant — no other literal changes.

### B2. Test SOP's visibility toggle

- `ApplySheetVisibility`: replace the unconditional `EnsureVisible
  SOPS_SHEET_NAME, xlSheetVisible` with a selection-driven show/hide — if the
  selection dict has a `Test SOP's` entry, honor it; otherwise default **visible**.
  The Test Selection sheet itself stays always-visible.
- `ResetSelectionToDefault`: set **both** `Lifetime Test` **and** `Test SOP's` to
  TRUE; everything else FALSE. (Introduce a small default-visible test rather than
  the single `DEFAULT_SELECTED_SHEET` compare.)
- `BuildKeepList`: **unchanged** — still adds `Test SOP's` unconditionally (always
  uploaded, even when its box is unticked).
- `CanonicalDataSheets()`: **unchanged** (12 sheets, no SOPs) ⇒ SOPs is never reset
  or snapshotted.
- `OnWorkbookSheetChange`: no special-casing — the SOP row is just another
  `DV_TestSelection` row, so ticking it triggers `ApplySheetVisibility`.

### B3. File name prompted at upload

- New `PromptForFileName()` — `Application.InputBox(..., Type:=2)` pre-filled from
  `GetNamed("DV_FileName")`; returns "" on Cancel.
- `Btn_UploadAll`: before the checklist, prompt; on "" abort with a status of
  "Cancelled (no file name)"; else `SetNamed "DV_FileName"` and continue.
- `RunChecklist`: **remove** the `DV_FileName is empty` check (filename is an
  upload-time concern now). Paths + data checks remain, so **Dry-Run** stays
  friction-free and never prompts.

### B4. Results via MsgBox

- `Btn_DryRunChecklist`: pass ⇒ `MsgBox "Checklist passed — ready to upload."`;
  fail ⇒ `MsgBox` the issue list. Still `StampLog` to `DV_Log` on `_Settings`.
- `Btn_UploadAll`: checklist failure, missing path, and "no data" branches each
  `MsgBox` their reason (in addition to the existing `SetNamed "DV_Status"`);
  success ⇒ a `MsgBox` summary (file name, destinations, Review copy kept); the
  `Failed:` handler ⇒ `MsgBox` the error.
- Helper `ShowFailures(title, failures)` formats a capped list (first ~20 + "…and N
  more — see the log on the hidden _Settings sheet").

### B5. Dynamic ribbon + new callbacks (all `Public` in `DataViewerUpload.bas`)

- Module-level `Public gRibbon As IRibbonUI`.
- `Ribbon_OnLoad(ribbon As IRibbonUI)` ⇒ `Set gRibbon = ribbon`.
- `RefreshPathLabels()` ⇒ if `gRibbon` is set, invalidate the three path controls
  (`ebSynPath`/`ebLocPath`/`ebExePath`). Guarded against a lost reference.
- `GetSynPathText` / `GetLocPathText` / `GetExePathText` (`getText`) ⇒ return the
  stored path. `GetSynPathTip` / `GetLocPathTip` / `GetExePathTip` (`getSupertip`)
  ⇒ return the full path. `PathReadOnly(control, text)` (`onChange`) ⇒ no-op +
  `RefreshPathLabels` (revert edits).
- `Btn_PickDataViewerExe` + `PickFileInto`; `Ribbon_PickDataViewer`.
- `Btn_ShowInstructions`; `Ribbon_Instructions`.
- `Btn_PickSynologyFolder` / `Btn_PickLocalFolder` call `RefreshPathLabels` after a
  successful pick.

`ThisWorkbook.cls.txt`: no change required (it calls public subs and references no
sheet-name literal); the ribbon `onLoad` is independent of `Workbook_Open`.

---

## Build script changes (`build_clean_template.py`)

- `UPLOAD_SHEET = "Test Selection"`; `KEEP` swaps `"DataViewer Upload"` →
  `"Test Selection"` and adds `"_Settings"`; `NAMED` updated to the §"Named-range
  targets" table (six names on `_Settings`, selection on `Test Selection`).
- New build steps (single-workbook, in place — no cross-workbook copy):
  1. **Carry-over read** — read existing `DV_FileName`/`Syn`/`Loc`/`Exe` values.
  2. **Rename** the upload sheet (`"DataViewer Upload"` if present, else already
     `"Test Selection"`) → `"Test Selection"`. Idempotent.
  3. **Create `_Settings`** (if missing): col-A labels, col-B carried-over values;
     `Visible = xlSheetVeryHidden`.
  4. **Re-lay the Test Selection sheet IN PLACE — preserve the native checkboxes.**
     COM cannot create native in-cell checkboxes (that needs the Office.js
     `range.control` API), so the build keeps the source's existing checkbox cells
     (`A3:A15`) and rewrites only around them — **never `Clear`s `A3:A15`**:
     - Write title (`A1`) + hint (`A2`); rewrite `B3:B15` with the 13 names in the
       new order and `A3:A15` with the boolean defaults (Lifetime + SOPs `True`,
       else `False`) — **value writes only**, so each cell keeps its checkbox.
     - Clear the clutter only: `A16:B16` (old stray row) + the instructions/settings/
       status regions (columns `C`+ and rows `16`+). Apply formatting to `A1:B15`.
     - Acceptance visually confirms `A3:A15` still render as checkboxes; if a COM
       value-write ever strips the control, fall back to the snapshot-seed procedure
       (Option B) documented in the runbook.
  5. **Recreate all seven named ranges** at the new refs.
  6. `_set_default_visibility`: visible set `{"Test Selection", "Test SOP's",
     "Lifetime Test"}`; very-hide names starting `_Template_`/`_Macro` **or equal to
     `_Settings`**.
  7. Existing steps unchanged: break stray links, stamp canonical sheets from
     snapshots, remove `Btn_*` shapes, import VBA, Save, then zip-level
     add-in-strip + ribbon inject.

## Validation / verification changes

- **`check_sources.py`** — extend invariants:
  - `customUI14.xml` declares `onLoad="Ribbon_OnLoad"` and the new control ids;
    **every** `onAction`/`getText`/`getSupertip`/`onChange`/`onLoad` callback in the
    ribbon has a matching `Public Sub` in `DataViewerUpload.bas` (extend the existing
    onAction↔handler cross-check to the new callback attributes).
  - `UPLOAD_SHEET_NAME = "Test Selection"` (no stray `"DataViewer Upload"` literal).
  - `PromptForFileName`/InputBox present and called from `Btn_UploadAll`; the
    `DV_FileName` empty-check is gone from `RunChecklist`.
  - `MsgBox` present in `Btn_DryRunChecklist` and `Btn_UploadAll`.
  - `ResetSelectionToDefault` defaults `Test SOP's` TRUE; `BuildKeepList` adds SOPs
    unconditionally.
  - Instructions text constant present; `gRibbon` + `RefreshPathLabels` present.
- **`test_build_helpers.py`** — `NAMED` targets `_Settings`/`Test Selection`,
  `_Settings` in `KEEP`, `UPLOAD_SHEET == "Test Selection"`.
- **`verify_sidecar.py`** — the customUI compare already covers the new ribbon. Add
  a light **workbook-structure** check (openpyxl): `"Test Selection"` exists,
  `"_Settings"` is very hidden, no `"DataViewer Upload"` sheet, and the seven named
  ranges resolve to the expected sheets.
- Then the usual: build → `verify_sidecar.py` → **operator acceptance** (§below).

## Resolved mechanics (settled by inspecting the live workbook)

1. **Checkboxes = native in-cell checkboxes, PRESERVED (not recreated).** The live
   sheet's column A holds real Booleans rendered by Excel's native cell-checkbox
   feature (cell metadata; confirmed: no Form-Control `ctrlProps`, no ActiveX, only
   `xl/metadata.xml`). Native checkboxes can only be *created* via the Office.js
   `range.control = {type:"Checkbox"}` API (ExcelApi 1.18+) — **not** via pywin32
   COM. The build therefore **preserves** the source's existing `A3:A15` checkbox
   cells (reorders by value-write, never clears them) rather than recreating them.
   **Fallback** if a COM value-write is ever found to strip the control (caught at
   acceptance): seed a baked `_Template_Sel` snapshot once and stamp it via a full
   `Worksheet.Copy` (Option B), documented in the runbook.
2. **Path rows = read-only `editBox`** (fixed `sizeString`, `getText` = full path,
   `getSupertip` = full path, `onChange` reverts ⇒ read-only). An editBox bounds its
   width and clips long paths, matching the "full path, hover for the rest" choice.

## Operator acceptance criteria (additions to the existing checklist)

1. Test Selection sheet shows **only** the title + hint + the 13-row checkbox table,
   in the exact order of §2.2; no instructions/settings/status remain.
2. Ticking/unticking a test shows/hides its sheet. Ticking **Test SOP's**
   shows/hides the SOP sheet; an **Upload All still includes Test SOP's even when its
   box is unticked**.
3. Ribbon: Upload All / Dry-Run are text-only and stacked; the three Pick buttons are
   stacked; Delete All Review Sheets is its own column; **no group exceeds 3 rows**
   and the tab is not crowded.
4. **Active Folders** shows the three current paths (full value on hover) and updates
   immediately after each Pick (incl. the new **Pick DataViewer File**).
5. **Instructions** opens the popup; the sheet carries no instructions.
6. **Upload All** prompts for a file name (pre-filled), and on success a popup
   summarizes; the `.xlsx` lands in Synology + Local + opens in DataViewer; a
   `… - Review` copy is kept; selection resets to Lifetime + SOPs.
7. **Dry-Run** shows a pass/fail popup and never prompts for a file name.
8. `_Settings` is invisible to the operator; a distributed `.xlsx` contains **neither**
   `_Settings` **nor** `Test Selection` (and no Review sheets).
9. All prior acceptance items (no save-prompt on close, nav/add/remove/reset,
   puff-picker crash-safety, Review accumulation, Delete-All) still pass.

## Risks

- **Checkbox creation via COM** is the main unknown; mitigated by the native/Form
  fallback and a live-workbook inspection before coding the build step.
- **Lost `IRibbonUI` on a VBA reset** would stop path-row refresh until the next file
  open; `RefreshPathLabels` is guarded, and `getText` re-reads on open, so paths are
  always correct after reopen.
- **Ribbon crowding** at five groups; Help/Instructions placement is trivially
  movable, and all groups are ≤3 rows by rule.

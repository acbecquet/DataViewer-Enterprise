# DataViewer Upload Sidecar — canonical macros for the testing template

This folder is the **single source of truth** for the VBA macros + ribbon that
live inside the operator's `Automated Testing Template.xlsm`. If you change the
workbook's behaviour, change it **here** and re-import — never edit the macros
only inside the `.xlsm`, or the repo and the deployed file drift apart (which
is exactly the bug that produced this folder; see
`../docs/superpowers/specs/2026-05-28-sidecar-reset-from-template-design.md`).

> The deployed `.xlsm` is MIP/IRM-encrypted at rest. The user's Python
> interpreter is on the MIP allowlist, so `verify_sidecar.py` and
> `build_clean_template.py` (run with that Python) read it as plaintext;
> ordinary tools (`cat`, `git`, the Edit/Read tools) see
> `%TSD-Header-...` ciphertext.

## What's in here

| File | Role | Where it goes in the workbook |
|------|------|-------------------------------|
| `DataViewerUpload.bas` | Upload, selection-driven sheet visibility, trim, **post-upload reset**, folder pickers | Standard module `DataViewerUpload` |
| `TestingTools.bas` | Add/Remove Sample, Reset Formulas (TPM Testing ribbon) | Standard module `TestingTools` (was `Module1`) |
| `SampleNav.bas` | Sample-block navigation hotkeys + ribbon | Standard module `SampleNav` |
| `ThisWorkbook.cls.txt` | `Workbook_Open` / `SheetChange` / `BeforeClose` (thin dispatch; `SheetBeforeDoubleClick` removed — selection is TRUE/FALSE dropdowns, not a double-click toggle) | Paste into the `ThisWorkbook` class module |
| `customUI14.xml` | The "TPM Testing" ribbon tab | Workbook `customUI` part (Custom UI Editor) |
| `*.png` | Ribbon button icons referenced by `customUI14.xml` | Workbook `customUI/images` |
| `verify_sidecar.py` | **Drift detector** — diffs the deployed `.xlsm` against these files | run from a shell |
| `build_clean_template.py` | **Clean rebuild** — copy source → open → drop non-kept sheets → re-save (clean package) → strip add-in + inject ribbon, via Excel COM + pywin32 | run on the work machine |
| `check_sources.py` | Headless invariant checker for the VBA/ribbon sources (no Excel needed) | run from a shell |

## How the workbook works (operator's view)

1. On open, only **Lifetime Test**, **Test Selection**, and **Test SOP's**
   are visible. The **Test Selection** sheet is a single checkbox table and
   nothing else: a title, a one-line hint, and 13 rows
   (`DV_TestSelection = 'Test Selection'!$A$3:$B$15`) — col A a native in-cell
   checkbox (boolean), col B the test name. The rows are in operator order:
   Custom Test Template, Lifetime Test, Long Puff Lifetime Test, Rapid Puff
   Lifetime Test, Intense Test, User Test Simulation, Big Headspace Serial Test,
   Viscosity Compatibility, Various Oil Compatibility, Temperature Cycling
   Test #1, Temperature Cycling Test #2, Negative Pressure Test, Test SOP's.
   Ticking a box unhides that test's sheet (defaults TRUE: Lifetime Test +
   Test SOP's). **Test SOP's** is a visibility toggle only — it is **always
   uploaded** whether or not its box is ticked.
2. The tech fills out the selected sheets (Add/Remove Sample + Reset Formulas
   on the TPM Testing ribbon; Ctrl+Shift+. / , to jump between samples).
3. **Upload All** prompts for a descriptive file name (InputBox, pre-filled with
   the last one), validates, stages a trimmed copy (selected + populated sheets +
   Test SOP's), converts it to a **macro-free `.xlsx`**, writes that `.xlsx` to
   `DV_SynologyPath` and `DV_LocalPath` and ingests the same file into DataViewer,
   then **resets the live sheets for the next session** (selection resets to
   Lifetime Test + Test SOP's) and re-hides everything. Outcomes (success,
   checklist failures, missing paths, errors) are reported in a **MsgBox**. Only
   the source template keeps macros; the distributed copies never do.

The paths, file name, and status/log no longer live on a visible sheet: they sit
on a **very-hidden `_Settings`** sheet (`DV_FileName`, `DV_SynologyPath`,
`DV_LocalPath`, `DV_DataViewerExe`, `DV_Status`, `DV_Log` in `B1:B6`), reachable
only through the ribbon and popups. Neither `_Settings` nor `Test Selection` is
included in a distributed `.xlsx`.

Required named ranges: `DV_FileName`, `DV_SynologyPath`, `DV_LocalPath`,
`DV_Status`, `DV_Log`, `DV_TestSelection`, `DV_DataViewerExe` (all but
`DV_TestSelection` now resolve onto `_Settings`).

## Clean rebuild (the canonical way to update the workbook)

The deployed `.xlsm` is rebuilt by `build_clean_template.py`, which follows this
codebase's proven single-workbook pattern: copy the source file, open the copy,
delete the non-kept sheets, stamp canonical sheets blank from their snapshots,
import the canonical VBA, and re-save (Excel rewrites the whole package — fresh,
no calcChain rot), then strip the embedded web add-in and swap in the repo ribbon
at the zip level. Run on a machine with Excel + pywin32:

    python excel-sidecar/build_clean_template.py --source "C:\path\Automated Testing Template v1.xlsm" --out "C:\path\Automated Testing Template v1 (clean).xlsm"
    python excel-sidecar/verify_sidecar.py --file "C:\path\Automated Testing Template v1 (clean).xlsm"

The second command should report all modules `MATCH`, `customUI14.xml == repo`,
and `no web-extension/add-in parts`. Requires Excel Trust Center → "Trust
access to the VBA project object model".

Headless checks that gate the sources (run anywhere, no Excel needed):

    python excel-sidecar/check_sources.py
    python excel-sidecar/test_build_helpers.py --source "C:\path\Automated Testing Template v1.xlsm"

Both must report `RESULT: ALL PASS` before a rebuild.

## Upload All keeps a Review copy (non-destructive reset)

On Upload All, after the data is distributed, each uploaded sheet is **renamed
into a `<name> - Review` copy** (data/formatting/formulas intact) and a fresh
blank sheet from the internal `_Template_NN` snapshot takes its place. Review
sheets persist (moved to the end of the workbook) until you click **Delete All
Review Sheets** on the ribbon. They never enter the distributed `.xlsx` copies.

## Ribbon

Every action lives on the **TPM Testing** ribbon tab. Groups, left→right:
**Sample Blocks** · **Sample Navigation** · **Help** · **DataViewer Upload** ·
**Active Folders**. Every group is at most 3 rows tall.

- **Help** — a single large **Instructions** button (info icon). Opens a MsgBox
  with ribbon-centric guidance; no instructions remain on any sheet. Placed
  immediately before DataViewer Upload.
- **DataViewer Upload** — three columns: a stacked text-only pair (**Upload All**,
  **Dry-Run Checklist**); a stacked picker trio (**Pick Synology Folder**, **Pick
  Local Folder**, **Pick DataViewer File**); and **Delete All Review Sheets** as
  its own large column.
- **Active Folders** — three read-only `editBox` path rows (Synology, Local,
  DataViewer) showing the current stored paths (full value on hover). They refresh
  immediately after any Pick. The ribbon captures `IRibbonUI` via
  `onLoad="Ribbon_OnLoad"` so the rows can be invalidated on demand.

## The post-upload reset (non-destructive)

After a successful upload, each uploaded sheet is reset **non-destructively**:
the live sheet is **renamed into a `<name> - Review` copy** (all data,
formatting, and formulas intact — a rename, not a delete) and a fresh blank
from the internal `_Template_NN` snapshot takes its place at the same tab
position. Review sheets are moved to the end of the workbook and stay visible
until you click **Delete All Review Sheets** on the ribbon; they never appear in
the distributed `.xlsx` copies (the staging-trim step excludes them).

If the same test is uploaded again before you delete the reviews, a counter is
appended: `… - Review 2`, `… - Review 3`, etc. Review-sheet names are capped
at Excel's 31-character limit using curated short base labels (§6.6 of the
spec).

The snapshots are **self-contained** -- no external template file, no DataViewer
install-path dependency, nothing to keep in sync. Build/refresh them with
`RebuildBlankTemplates` (Alt+F8) on a BLANK workbook; it refuses if any sheet
contains data, so you can't bake real samples into a template. The `_Template_`
prefix means snapshots are auto-very-hidden (ApplySheetVisibility) and
auto-excluded from the distributed `.xlsx` copies (the trim step).

**Fail-safe:** a sheet whose snapshot is missing, or whose copy adds no sheet,
is left fully intact (the rename is reversed) and logged. Nothing is ever lost.

Design:
`../docs/superpowers/specs/2026-06-04-excel-sidecar-clean-rebuild.md` (§6.4).
Cell map: `../docs/superpowers/specs/template-cell-map.md`.

## Install / update the workbook

Rebuild with `build_clean_template.py` (see **Clean rebuild** above); the
result matches this folder by construction.

### Manual fallback (always works — no pywin32 needed)

1. Open a new blank workbook in Excel.
2. Move/Copy each sheet from the source into the new workbook (right-click tab
   → Move or Copy, "Create a copy" checked), maintaining order.
3. Alt+F11 (VBE). Delete any old `DataViewerUpload`, `Module1`/`TestingTools`,
   `SampleNav` modules. File → Import File → import `DataViewerUpload.bas`,
   `TestingTools.bas`, `SampleNav.bas`. Double-click `ThisWorkbook`; paste the
   body of `ThisWorkbook.cls.txt` (after `Option Explicit`).
4. Apply the ribbon: open the file in the **Custom UI Editor**, paste the
   contents of `customUI14.xml` under the `customUI14` node, save.
5. Save as `.xlsm`. Close and reopen.

## Blank-template snapshots (the reset source)

The post-upload reset restores each sheet from a very-hidden `_Template_NN`
snapshot stored inside the workbook. Two macros build them (Alt+F8):

- **`RebuildBlankTemplates`** - snapshots the workbook's CURRENT sheets. Run on
  a BLANK workbook; it refuses if any sheet has entered samples.
- **`SeedBlankTemplatesFromFile`** - snapshots from an EXTERNAL blank workbook
  you pick (e.g. `Automated Testing Template - DVE.xlsm`). Use this to add
  snapshots to a workbook that already has data -- the data sheets are not
  touched.

To migrate an existing operator template, follow
`RUNBOOK-migrate-existing-template.md`.

## Verify the deployed file matches this folder

```bash
python excel-sidecar/verify_sidecar.py --file "C:\path\to\Automated Testing Template.xlsm"
```

Reports each module as `MATCH` or `DIFFERS` (with a short diff). Run it any
time you suspect drift, and after installing (everything should be `MATCH`).
Before re-importing the fix, expect `DataViewerUpload` to report `DIFFERS` —
that is the detector telling you the deployed reset is out of date.

## Known issues

- **`MakeTempXlsx`** reopens a same-VBA-codename copy and `SaveAs`. It runs on
  a temp staging copy (not `ThisWorkbook`), so it is not the data-loss cause,
  but it's a latent COM concern. Left unchanged.
- **`TestingTools.ResetEquations`** ("Reset Formulas" button) restores col-A
  formulas from `_Template_Master`, which imposes puff interval 20 on any
  sheet. Fine for Lifetime; wrong interval for Intense (10) / Negative
  Pressure (1). Confirm before fixing — it's operator-invoked, with a prompt.

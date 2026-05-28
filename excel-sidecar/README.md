# DataViewer Upload Sidecar — canonical macros for the testing template

This folder is the **single source of truth** for the VBA macros + ribbon that
live inside the operator's `Automated Testing Template.xlsm`. If you change the
workbook's behaviour, change it **here** and re-import — never edit the macros
only inside the `.xlsm`, or the repo and the deployed file drift apart (which
is exactly the bug that produced this folder; see
`../docs/superpowers/specs/2026-05-28-sidecar-reset-from-template-design.md`).

> The deployed `.xlsm` is MIP/IRM-encrypted at rest. The user's Python
> interpreter is on the MIP allowlist, so `verify_sidecar.py` /
> `install_sidecar.py` (run with that Python) read it as plaintext; ordinary
> tools (`cat`, `git`, the Edit/Read tools) see `%TSD-Header-...` ciphertext.

## What's in here

| File | Role | Where it goes in the workbook |
|------|------|-------------------------------|
| `DataViewerUpload.bas` | Upload, selection-driven sheet visibility, trim, **post-upload reset**, folder pickers | Standard module `DataViewerUpload` |
| `TestingTools.bas` | Add/Remove Sample, Reset Formulas (TPM Testing ribbon) | Standard module `TestingTools` (was `Module1`) |
| `SampleNav.bas` | Sample-block navigation hotkeys + ribbon | Standard module `SampleNav` |
| `ThisWorkbook.cls.txt` | `Workbook_Open` / `SheetChange` / `SheetBeforeDoubleClick` / `BeforeClose` | Paste into the `ThisWorkbook` class module |
| `customUI14.xml` | The "TPM Testing" ribbon tab | Workbook `customUI` part (Custom UI Editor) |
| `*.png` | Ribbon button icons referenced by `customUI14.xml` | Workbook `customUI/images` |
| `verify_sidecar.py` | **Drift detector** — diffs the deployed `.xlsm` against these files | run from a shell |
| `install_sidecar.py` | Optional one-shot importer (dry-run default, backs up first) | run from a shell |

## How the workbook works (operator's view)

1. On open, only **Lifetime Test**, **DataViewer Upload**, and **Test SOP's**
   are visible. The `DataViewer Upload` sheet has a `DV_TestSelection` range of
   TRUE/FALSE checkboxes (double-click col A rows 20-33 to toggle). Toggling
   TRUE unhides that test sheet.
2. The tech fills out the selected sheets (Add/Remove Sample + Reset Formulas
   on the TPM Testing ribbon; Ctrl+Shift+. / , to jump between samples).
3. **Upload All** validates, stages a trimmed copy (selected + populated
   sheets + Test SOP's), converts it to a **macro-free `.xlsx`**, writes that
   `.xlsx` to `DV_SynologyPath` and `DV_LocalPath` and ingests the same file
   into DataViewer, then **resets the live sheets for the next session** and
   re-hides everything. Only the source template keeps macros; the
   distributed copies never do.

Required named ranges: `DV_FileName`, `DV_SynologyPath`, `DV_LocalPath`,
`DV_Status`, `DV_Log`, `DV_TestSelection`. Optional: `DV_DataViewerExe`.

## The post-upload reset

After a successful upload, each uploaded sheet is **reverted to a blank snapshot
that lives inside the workbook** -- a very-hidden `_Template_NN` sheet, one per
canonical test type. The live sheet is deleted and replaced by a pristine copy
of its snapshot, so headers, per-sheet puff seed/interval, formula scaffolding,
formatting, and milestone notes are all restored exactly, at the snapshot's
sample-block count. Grow the sheet again with Add Sample for the next campaign.

The snapshots are **self-contained** -- no external template file, no DataViewer
install-path dependency, nothing to keep in sync. Build/refresh them with
`RebuildBlankTemplates` (Alt+F8) on a BLANK workbook; it refuses if any sheet
contains data, so you can't bake real samples into a template. The `_Template_`
prefix means snapshots are auto-very-hidden (ApplySheetVisibility) and
auto-excluded from the distributed `.xlsx` copies (the trim step).

**Fail-safe:** a sheet whose snapshot is missing -- or whose A1 title doesn't
match its snapshot -- is left intact and logged. Nothing is ever cleared without
a known-good source to restore from.

First-time setup: open a BLANK template, import the macros, run
`RebuildBlankTemplates` once -- the workbook is then self-contained. Design:
`../docs/superpowers/specs/2026-05-28-sidecar-reset-from-template-design.md`.
Cell map: `../docs/superpowers/specs/template-cell-map.md`.

## Install / update the workbook

The deployed file currently runs an **older** `DataViewerUpload` (with the
data-loss reset). Re-import to pick up the fix.

### Option A — one-shot script (recommended)

```bash
# dry run first (no changes, just reports the plan + checks prerequisites)
python excel-sidecar/install_sidecar.py --file "C:\path\to\Automated Testing Template.xlsm"
# apply (makes a timestamped .bak copy first)
python excel-sidecar/install_sidecar.py --file "C:\path\to\Automated Testing Template.xlsm" --apply
```

Requires `pywin32` and Excel Trust Center → "Trust access to the VBA project
object model" enabled. The script backs up the workbook, removes the old
modules (`DataViewerUpload`, `Module1`, `TestingTools`, `SampleNav`), imports
the canonical `.bas` files, and sets the `ThisWorkbook` code. The ribbon
(`customUI14.xml`) and icons must be applied with the Custom UI Editor if they
ever change (they rarely do).

### Option B — manual (always works)

1. Alt+F11 (VBE). Delete the old `DataViewerUpload`, `Module1`/`TestingTools`,
   `SampleNav` modules.
2. File → Import File → import `DataViewerUpload.bas`, `TestingTools.bas`,
   `SampleNav.bas`.
3. Double-click `ThisWorkbook` in the Project pane; paste the body of
   `ThisWorkbook.cls.txt` (below the existing `Option Explicit`).
4. Save. Close and reopen to fire `Workbook_Open`.

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

## Known issues (out of scope for the reset fix)

- **`MakeTempXlsx`** reopens a same-VBA-codename copy and `SaveAs`. It runs on
  a temp staging copy (not `ThisWorkbook`), so it is not the data-loss cause,
  but it's a latent COM concern. Left unchanged.
- **`TestingTools.ResetEquations`** ("Reset Formulas" button) restores col-A
  formulas from `_Template_Master`, which imposes puff interval 20 on any
  sheet. Fine for Lifetime; wrong interval for Intense (10) / Negative
  Pressure (1). Confirm before fixing — it's operator-invoked, with a prompt.

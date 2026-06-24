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
| `verify_sidecar.py` | **Drift detector** — diffs the deployed `.xlsm` against these files + structure checks (presence, visibility, anchors, banner, signature) | run from a shell |
| `build_clean_template.py` | **Clean rebuild** — copy source → open → drop non-kept sheets → re-save (clean package) → strip add-in + inject ribbon, via Excel COM + pywin32 | run on the work machine |
| `check_sources.py` | Headless invariant checker for the VBA/ribbon sources (no Excel needed) | run from a shell |
| `test_build_helpers.py` | Headless tests for the build helpers (run against a source `.xlsm`, no Excel needed) | run from a shell |
| `signing/` | **VBA code-signing kit** — owner cert script, tester one-time trust script, signing runbook (`signing/README.md`) | owner + tester machines |

## How the workbook works (operator's view)

1. On open, only **Lifetime Test**, **Test Selection**, and **Test SOP's**
   are visible. The **Test Selection** sheet is a single checkbox table:
   a title (B2), a one-line hint (B3), and 13 rows
   (`DV_TestSelection = 'Test Selection'!$B$4:$C$16`) — col B a native in-cell
   checkbox (boolean), col C the test name. The rows are in operator order:
   Custom Test Template, Lifetime Test, Long Puff Lifetime Test, Rapid Puff
   Lifetime Test, Intense Test, User Test Simulation, Big Headspace Serial Test,
   Viscosity Compatibility, Various Oil Compatibility, Temperature Cycling
   Test #1, Temperature Cycling Test #2, Negative Pressure Test, Test SOP's.
   Ticking a box unhides that test's sheet (defaults TRUE: Lifetime Test +
   Test SOP's); ticking a test whose sheet was **deleted** restores it fresh
   from its `_Template_NN` snapshot (the **lockbox** — tester mutations never
   break the flow). **Test SOP's** is a visibility toggle only — it is
   **always uploaded** whether or not its box is ticked. Below the table
   (rows 18–20) sits a three-line guidance **banner** ("Macros required…",
   "never email or upload this file…", "Upload All / Upload Checkpoint
   deliver…") — the only guidance channel that still works with macros
   disabled. Test Selection never enters distributed copies, so the banner
   can be loud.
2. The tech fills out the selected sheets (Add/Remove Sample + Reset Formulas
   on the TPM Testing ribbon; Ctrl+Shift+. / , to jump between samples).
   **Add/Remove Sample** detection is **structural**: it operates on every
   checkbox test sheet that uses the standard 12-column "puffs" block layout —
   i.e. **all of them except `Temperature Cycling Test #1`** (a step checklist)
   **and `Test SOP's`** (prose). There is no hardcoded sheet-name list, so it
   won't silently break on renamed or copied sheets.
   The **Clog** column is **automatic** — operators no longer type it. Each
   block's Clog is derived from that block's **Draw Pressure (kPa)**:
   non-numeric (e.g. `n/a`), blank or ≤ 5 → empty; > 5 and < 15 → **"Light
   Clog"** (yellow highlight); **≥ 15 → "Heavy Clog"** (red highlight, white
   text). **Reset Formulas** restores the calculated formulas (Clog included)
   in R1C1 form, so every block references its own columns — and it
   deliberately does **not** touch the puffs or before-weight columns, where
   testers type literals.
   The **puffs** column auto-fills: type **ANY positive number** into a puffs
   cell and every row below becomes "the row above + your number"; typing in
   the first data row (row 5) seeds the whole column. For a literal or
   irregular sequence, pick **`custom`** (clears the column) and then
   **paste** your numbers — typing them in one by one re-triggers the fill.
   Review copies and `_`-prefixed sheets are never touched by the picker.
3. **Upload All** prompts for a descriptive file name (InputBox, pre-filled
   with the last one; a name with no date gets today's date appended),
   validates, stages a trimmed copy (selected + populated sheets + Test
   SOP's), then materializes ONE clean copy by copying the staged sheets into
   a **fresh workbook** — so the distributed `.xlsx` is **macro-free AND
   ribbon-free** (no dead "TPM Testing" tab, no "Cannot run the macro" popups
   for receivers; tester-created **chart sheets** are also excluded by the
   trim). It writes that `.xlsx` to `DV_SynologyPath` and `DV_LocalPath`,
   launches DataViewer **on the Synology copy** (so later in-app edits land in
   the durable file), then **resets the live sheets for the next session**
   (selection resets to Lifetime Test + Test SOP's), restores any missing
   canonical sheets from their snapshots (lockbox — the workbook always ends
   canonical), and re-hides everything. Success shows a **delivery receipt**
   (the data is already delivered, where the shareable copy lives, never email
   the template, open-the-Synology-folder offer); failures distinguish
   "nothing was lost — run it again" (pre-delivery) from "your data WAS
   delivered — do NOT re-enter it" (post-delivery). Only the source template
   keeps macros; the distributed copies never do.
4. **Upload Checkpoint** is the same delivery **without the reset**: sheets
   are NOT reset and no Review copies are made — keep testing and upload again
   any time. Checkpoint stream: re-using the **same file name** silently
   overwrites that stream's own previous copies (`DV_LastUpload` remembers the
   last delivered name); a name that collides with a file this workbook did
   **not** just deliver gets a confirm first. Known limitation: `DV_LastUpload`
   is **single-slot** — alternating between two checkpoint streams in one
   session prompts the overwrite confirm on each switch (safe, just a confirm).
5. Closing the workbook while it still holds data that was not uploaded this
   session shows a **one-button reminder** (never a gate — closing is always
   allowed).

The paths, file name, and status/log no longer live on a visible sheet: they sit
on a **very-hidden `_Settings`** sheet (`DV_FileName`, `DV_SynologyPath`,
`DV_LocalPath`, `DV_DataViewerExe`, `DV_Status`, `DV_Log`, `DV_OrigFileName`,
`DV_LastUpload` in `B1:B8`), reachable only through the ribbon and popups.
Picked paths are saved to disk immediately (close + "Don't Save" no longer
discards them). Neither `_Settings` nor `Test Selection` is included in a
distributed `.xlsx`.

Required named ranges: `DV_FileName`, `DV_SynologyPath`, `DV_LocalPath`,
`DV_Status`, `DV_Log`, `DV_TestSelection`, `DV_DataViewerExe`,
`DV_OrigFileName`, `DV_LastUpload` (all but `DV_TestSelection` resolve onto
`_Settings`; `DV_TestSelection = 'Test Selection'!$B$4:$C$16`).

## Clean rebuild (the canonical way to update the workbook)

The deployed `.xlsm` is rebuilt by `build_clean_template.py`, which follows this
codebase's proven single-workbook pattern: copy the source file, open the copy,
delete the non-kept sheets, stamp canonical sheets blank from their snapshots,
import the canonical VBA, and re-save (Excel rewrites the whole package — fresh,
no calcChain rot), then strip the embedded web add-in and swap in the repo ribbon
at the zip level. Run on a machine with Excel + pywin32:

    python excel-sidecar/build_clean_template.py --source "C:\path\Automated Testing Template v1.1.xlsm" --out "C:\path\Automated Testing Template v1.2.xlsm"
    python excel-sidecar/verify_sidecar.py --file "C:\path\Automated Testing Template v1.2.xlsm"

Build guards: the script refuses to run when `--out` is the same file as
`--source`, refuses an existing `--out` unless you pass `--force` (which first
backs the old `--out` up), always takes a timestamped
`.bak-YYYYmmdd_HHMMSS` copy of the source, and aborts if the source reads as
MIP ciphertext under the wrong interpreter.

The second command should report all modules `MATCH`, `customUI14.xml == repo`,
`no web-extension/add-in parts`, and the structure checks OK (sheet presence +
visibility, named-range anchors including the literal
`'Test Selection'!$B$4:$C$16`, the featurePropertyBag checkbox carrier, the
banner). Requires Excel Trust Center → "Trust access to the VBA project
object model".

## VBA signing

Every shipped build gets its VBA project **signed** so testers see zero macro
prompts — including on copies that arrive with Mark of the Web. The kit (owner
cert script `make_cert.ps1`, tester one-time `tester-setup.ps1`, and the full
owner/tester flow) lives in **`signing/README.md`**. After signing, gate the
file with `verify_sidecar.py --file <out> --require-signature`; any VBA edit
inside the workbook strips the signature, so the check doubles as a drift
alarm.

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

- **Help** — three large buttons: **Sheet Guide** (the 12-column block layout,
  the dropdowns, the auto-calculated columns), **Sample Tools** (Add/Remove
  Sample, puffs auto-fill, navigation shortcuts), and **Instructions** (the
  upload workflow). Each opens a MsgBox; no instructions remain on any sheet.
  Placed immediately before DataViewer Upload.
- **DataViewer Upload** — three columns: a stacked text-only trio (**Upload
  All**, **Upload Checkpoint**, **Specify Test Name**); a stacked picker trio
  (**Pick Synology Folder**, **Pick Local Folder**, **Pick DataViewer File**);
  and **Delete All Review Sheets** as its own large column.
  - **Specify Test Name** prompts for a project/test name and renames the
    workbook on disk to `<project> - <original file name> (do not send).xlsm`
    (a clean rename — the previous file is removed; if removal fails, a loud
    warning says which stale file to delete). The "(do not send)" marker flags
    the working copy as a non-deliverable exactly where the
    email-the-template mistake happens: the Explorer / Outlook attach dialog.
    If a file already exists at the target name, a confirm is shown — never a
    silent overwrite. Entering a **blank** name restores the original file
    name (marker gone). This only affects the **on-disk workbook name**;
    **uploads automatically restore the original name** first, and the
    uploaded data file name is unchanged (still the descriptive name entered
    at Upload time).
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
Sheets that could not be reset are also listed honestly in the delivery
receipt — the data is safe, it just still sits on the live sheet.

**Lockbox reconcile:** after the per-sheet resets, Upload All restores any
canonical sheet the tester **deleted or renamed** during the session from its
`_Template_NN` snapshot, so the workbook always ends canonical. The same
restore runs when a tester re-ticks a deleted sheet's box on Test Selection.

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
v1.2 also checks structure: sheet presence + visibility states, the
`featurePropertyBag` part (the native-checkbox carrier), the literal
`DV_TestSelection` anchor (`'Test Selection'!$B$4:$C$16`), `DV_LastUpload`,
the Test Selection banner, and the ISNUMBER Clog formula. Pass
`--require-signature` after signing to additionally enforce a VBA project
signature (see `signing/README.md`).

## Known issues / limitations

- **`DV_LastUpload` is single-slot.** The checkpoint stream remembers only
  the LAST delivered name, so alternating between two checkpoint streams in
  one session prompts the overwrite confirm on each switch. Safe — just a
  confirm — and intentional for v1.2.
- *(resolved in v1.2)* **`MakeTempXlsx`** no longer `SaveAs`-es a
  same-VBA-codename copy: the staged sheets are copied into a **fresh
  workbook**, which also makes distributed copies ribbon-free (no customUI
  part survives; audit H9).
- *(resolved in v1.2)* **`TestingTools.ResetEquations`** no longer imposes the
  master's puff interval: the puffs and before-weight columns are deliberately
  **not restored** (testers type literals there — re-seed puffs by typing a
  value in row 5), and the remaining formulas restore via R1C1 so every block
  references its own columns.

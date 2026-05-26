# Handoff: integrate `DataViewerUpload.bas` into the Automated Testing Template

You are helping integrate a VBA module (`DataViewerUpload.bas`) into a
macro-enabled Excel workbook (`Automated Testing Template.xlsm`). The
module powers two buttons on a sheet named **"DataViewer Upload"**:

- **Button 1 → `Btn_UploadAll`** — runs a data-population checklist; if it
  passes, copies the workbook to two configured destination paths and
  shells `DataViewer.exe` so it ingests the data into Postgres.
- **Button 2 → `Btn_DryRunChecklist`** — runs only the checklist, no
  copies, no DataViewer launch.

Your job is to (1) verify the workbook has all the pieces the module
expects, (2) create/wire anything missing, and (3) walk the user
through a smoke test of both buttons. **Do not modify the `.bas` source
itself — adjust the workbook to match what the module reads.**

---

## 1. What the module expects

### 1a. Named ranges (must exist; case-sensitive)

These live on the **"DataViewer Upload"** sheet. The first three are
inputs the user fills in; the last two are outputs the macro writes to.

| Name              | Read/Write | Type             | Purpose                                                                   |
|-------------------|------------|------------------|---------------------------------------------------------------------------|
| `DV_FileName`     | Read       | single-cell text | Base filename for the destination copies. No extension. e.g. `Lot42_2026-05-21` |
| `DV_SynologyPath` | Read       | single-cell text | Folder path #1 to receive the `.xlsm` copy. e.g. `\\192.168.222.10\DataViewer\Incoming` |
| `DV_LocalPath`    | Read       | single-cell text | Folder path #2 to receive the `.xlsm` copy. e.g. `C:\DataViewer\Local`    |
| `DV_Status`       | Write      | single-cell text | Macro writes `"OK"` on success or `"Failed: <reason>"`                    |
| `DV_Log`          | Write      | single-cell text | Macro appends timestamped progress lines (multi-line; use `WRAP TEXT`)    |

**Optional:**

| Name              | Read | Purpose                                                                                       |
|-------------------|------|-----------------------------------------------------------------------------------------------|
| `DV_DataViewerExe`| Read | Override path to `DataViewer.exe`. Defaults to `C:\Program Files\DataViewer Enterprise\DataViewer.exe` |

If any of `DV_FileName`, `DV_SynologyPath`, or `DV_LocalPath` is empty,
the checklist fails up front.

### 1b. Sheets the checklist validates

The workbook must contain these 10 sheets (case-insensitive match):

1. `Lifetime Test`
2. `User Test Simulation`
3. `Long Puff Lifetime Test`
4. `Rapid Puff Lifetime Test`
5. `Intense Test`
6. `Big Headspace Serial Test`
7. `Negative Pressure Test`
8. `Temperature Cycling Test #2`
9. `Viscosity Compatibility`
10. `Various Oil Compatibility`

These sheets all share the same sample-block layout (next section). The
procedure sheet `Temperature Cycling Test #1` and the reference sheet
`Test SOP's` are **not validated** — they're allowed to exist with any
content.

### 1c. Sample-block layout (within each data sheet)

Each data sheet contains up to **8 sample blocks**, each **12 columns wide**.
Block N starts at column `N * 12 + 1` (1-based):

| Sample # | Start col | Excel column |
|----------|-----------|--------------|
| 0        | 1         | A            |
| 1        | 13        | M            |
| 2        | 25        | Y            |
| 3        | 37        | AK           |
| 4        | 49        | AW           |
| 5        | 61        | BI           |
| 6        | 73        | BU           |
| 7        | 85        | CG           |

**Within a block**, the checklist reads these cells:

| Row | Col offset (0-based) | Excel ref (sample 0) | Field                     |
|-----|----------------------|----------------------|---------------------------|
| 1   | +5                   | F1                   | **Sample ID** (must be unique within sheet) |
| 1   | +7                   | H1                   | **Heating technology** (must be non-empty if sample ID is non-empty) |
| 5+  | +0                   | A5 ...               | Puffs                     |
| 5+  | +1                   | B5 ...               | Mass before               |
| 5+  | +2                   | C5 ...               | Mass after                |
| 5+  | +8                   | I5 ...               | TPM (calculated)          |

A block is considered **populated** when its sample ID cell (row 1,
col offset +5) is non-empty after `Trim$`. Blocks with empty sample ID
are skipped (not an error).

A data row is considered **blank** (and ends the per-block scan) when
columns 0, 1, AND 2 are all empty. The scan starts at row 5 and runs
up to row 10000.

### 1d. Checklist rules

For each populated sample block:

- Sample ID is unique within the sheet.
- Heating technology cell is non-empty.
- For each populated data row:
  - If puffs > 0, the value is strictly greater than the previous
    populated row's puffs.
  - If mass-before > 0 and mass-after > 0, then mass-before > mass-after.
  - TPM is in the range `[0, 50]` mg/puff (permissive sanity bound).

---

## 2. What `Btn_UploadAll` does, step by step

1. Clears `DV_Log` and sets `DV_Status = "Starting upload..."`.
2. Runs the checklist; if there are failures, writes them line-by-line
   into `DV_Log`, sets `DV_Status = "Failed: N issue(s)"`, **aborts**.
3. `Application.DisplayAlerts = False`, `ThisWorkbook.Save`,
   `Application.DisplayAlerts = True`. This persists in-memory edits
   so the destination copies match what's on screen.
4. Checks that both destination folders exist (`FSO.FolderExists`).
   If either is missing → `DV_Status = "Failed: ... path not accessible"`.
5. Copies the `.xlsm` to:
   - `<DV_SynologyPath>\<DV_FileName>.xlsm`
   - `<DV_LocalPath>\<DV_FileName>.xlsm`
   Uses `FSO.CopyFile <src>, <dest>, True` (overwrite enabled).
6. Materializes a temporary `.xlsx` for DataViewer:
   - Copies the workbook to `%TEMP%\dvupload_YYYYMMDD_HHMMSS_<name>.xlsm`.
   - `Workbooks.Open` that staging `.xlsm` read-only.
   - `SaveAs FileFormat:=51` (`xlOpenXMLWorkbook`, i.e. `.xlsx`) to
     `%TEMP%\dvupload_YYYYMMDD_HHMMSS_<name>.xlsx`.
   - Closes the staged workbook (no save).
   - Deletes the staging `.xlsm`.
   - This strips macros — DataViewer doesn't need them and the COM
     reader is happier with plain `.xlsx`.
7. Resolves the DataViewer executable: `DV_DataViewerExe` if set,
   otherwise the default `C:\Program Files\DataViewer Enterprise\DataViewer.exe`.
   If the file doesn't exist → `DV_Status = "Failed: DataViewer.exe not found at ..."`.
8. `Shell "<dvExe>" "<tempXlsx>", vbNormalFocus`. Fire-and-forget.
   DataViewer's `SingleInstance` hands the path to a running instance
   (or starts a new one); `onFileLoadFinished` auto-saves the parsed
   data to Postgres.
9. Sets `DV_Status = "OK"` and a final timestamped log line.

The temp `.xlsx` is left in `%TEMP%` — Windows cleanup will reclaim it.

## 3. What `Btn_DryRunChecklist` does

1. Clears `DV_Log` and sets `DV_Status = "Running checklist..."`.
2. Runs the same `RunChecklist()` function.
3. If zero failures: `DV_Status = "OK"`.
4. If failures: writes each one to `DV_Log` with a timestamp, sets
   `DV_Status = "Failed: N checklist issue(s)"`.

No file copies, no DataViewer launch, no workbook save. Safe to click
anytime.

---

## 4. Your integration checklist

Work through these in order. Do not skip steps — the macro has no
fallback if a named range is missing; it will silently treat it as
empty and the checklist will report it as "DV_X is empty".

### 4a. Module import

1. Open the workbook in Excel (macros enabled).
2. `Alt+F11` to open the VBE.
3. `File → Import File...` → select `DataViewerUpload.bas`.
4. Confirm a new module named **`DataViewerUpload`** appears under
   `VBAProject (Automated Testing Template.xlsm) → Modules`.

### 4b. Named-range setup

For each name in the table in §1a, verify it exists on the
**"DataViewer Upload"** sheet. To check, use `Formulas → Name Manager`.

If a name is missing:

1. Click the cell where the value lives (or should live).
2. In the Name Box (left of the formula bar), type the name and press Enter.
3. Verify in Name Manager that **Scope** is "Workbook" (not a sheet
   scope) and **Refers to** is a single-cell absolute reference, e.g.
   `='DataViewer Upload'!$B$3`.

Place inputs (`DV_FileName`, `DV_SynologyPath`, `DV_LocalPath`) with a
label cell to their left so the user knows what to type. Place outputs
(`DV_Status`, `DV_Log`) with labels and apply **Wrap Text** + a tall
row to `DV_Log` so the multi-line output is readable.

### 4c. Sheet name verification

In Name Manager (or by clicking each sheet tab), confirm all 10
canonical sheet names from §1b exist with the exact spelling and
casing shown. The match is case-insensitive in the macro, but extra
spaces or different separators will fail. If a sheet is named (e.g.)
`Lifetime` instead of `Lifetime Test`, rename it.

### 4d. Sample-block layout verification

Open `Lifetime Test` (or any other populated data sheet). For sample 0:

- **F1** should contain the sample ID (e.g. `CCELL3.0-001`).
- **H1** should contain the heating technology (e.g. `CCELL3.0`).
- **A4** should be the "Puffs" header; **A5** the first puff count.
- **B5** should be the "Mass Before" value, **C5** the "Mass After"
  value, **I5** the calculated TPM.

For sample 1, the same fields live at **R1** (sample ID), **T1**
(heating tech), **M4** (Puffs header), etc. The pattern repeats every
12 columns.

If the user's template puts sample ID somewhere other than row 1
col offset +5, **stop and tell them** — the macro's constants
(`startCol + 5` for sample ID, `startCol + 7` for heating tech) would
need adjustment in the `.bas`. Don't silently rewrite the layout.

### 4e. Button wiring

1. Return to the "DataViewer Upload" sheet.
2. Right-click button 1 (the upload-all button) → `Assign Macro...` →
   select `Btn_UploadAll` → OK.
3. Right-click button 2 (the dry-run button) → `Assign Macro...` →
   select `Btn_DryRunChecklist` → OK.

If the buttons don't exist yet: `Developer → Insert → Form Controls →
Button (Form Control)`, draw a button, set its caption, then assign
the macro.

### 4f. Save the workbook

`Ctrl+S`. Confirm the macro persists across close/reopen.

### 4g. Smoke test

1. **Dry-run first.** With clean inputs (`DV_FileName` filled,
   both paths filled), click button 2. `DV_Status` should read `"OK"`
   within ~1 second and `DV_Log` should show a "Checklist passed" line.
2. **Negative test.** Clear `DV_FileName`, click button 2 again.
   `DV_Status` should now say `"Failed: 1 checklist issue(s)"` and
   `DV_Log` should contain `"  - DV_FileName is empty"`.
3. **Full upload.** Restore `DV_FileName`, click button 1.
   - `DV_Log` should show "Saving workbook" → "Copy → ..." (twice) →
     "Temp .xlsx: ..." → "Shell: ..." → "Upload dispatched...".
   - DataViewer Enterprise should pop to the front with the new file
     loaded in its file tree.
   - In DataViewer's bottom status bar, the DB sync indicator should
     go from "Saving..." to a synced state within a few seconds.
4. Verify the destination files exist:
   - `<DV_SynologyPath>\<DV_FileName>.xlsm`
   - `<DV_LocalPath>\<DV_FileName>.xlsm`

If DataViewer fails to launch, check that
`C:\Program Files\DataViewer Enterprise\DataViewer.exe` exists (or set
`DV_DataViewerExe` to the actual install path).

---

## 5. Constants the `.bas` declares

For reference; only change these if the user's template layout differs
from §1c:

```vba
Private Const DEFAULT_DATAVIEWER_EXE As String = _
    "C:\Program Files\DataViewer Enterprise\DataViewer.exe"
Private Const COLS_PER_SAMPLE As Long = 12
Private Const MAX_SAMPLES_PER_SHEET As Long = 8
Private Const FIRST_DATA_ROW As Long = 5
Private Const TPM_MAX_PLAUSIBLE As Double = 50#
```

`MAX_SAMPLES_PER_SHEET = 8` matches the A/M/Y/AK/AW/BI/BU/CG block
starts. If a sheet's layout supports more than 8 samples, bump this.

`FIRST_DATA_ROW = 5` assumes rows 1-3 are metadata and row 4 is the
column-header row. If the user's template has a different metadata
height, this needs updating.

---

## 6. Known caveats to surface to the user

- **The macro `Shell`'s fire-and-forget.** `DV_Status = "OK"` means
  "DataViewer was launched with the file" — not "Postgres write
  succeeded". To confirm the DB write, the user has to watch
  DataViewer's status bar. If DataViewer's NAS connection is offline,
  the file appears loaded but the save is queued / silently failed.
- **Temp `.xlsx` is left in `%TEMP%`.** Windows cleans this up
  eventually. No active cleanup logic.
- **`SaveAs FileFormat:=51` strips macros.** Intentional — DataViewer
  doesn't need them. The original `.xlsm` (with macros) is what
  gets copied to the destination paths.
- **Case-insensitive sheet matching.** `lifetime test` and `Lifetime
  Test` are treated the same. But trailing spaces or different
  punctuation (`Test#2` vs `Test #2`) are different sheets.

Report back when integration is complete, with the contents of
`DV_Log` after the smoke test in §4g step 3 (the full upload run).

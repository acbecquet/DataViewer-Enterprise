# Bug Report: `Btn_UploadAll` data-loss in `excel-sidecar/DataViewerUpload.bas`

## Symptom
After running `Btn_UploadAll` from the .xlsm template:
- Postgres receives correct test data.
- The source `.xlsm` ends up visually cleared of sample data, leaving only
  the template skeleton.

## Root cause

Two separate, compounding defects in `DataViewerUpload.bas`:

### 1. `MakeTempXlsx` (lines 348-377) re-opens a near-identical copy of `ThisWorkbook` inside the same Excel app

The function does:

1. `fso.CopyFile ThisWorkbook.FullName, stagedXlsm` (line 358) — file-system copy of the source.
2. `Application.EnableEvents = False` (line 361).
3. `Workbooks.Open(stagedXlsm, ReadOnly:=True)` (line 364) — opens the staged copy as a SECOND workbook in the same Excel instance.
4. `wb.SaveAs Filename:=outXlsx, FileFormat:=51` (line 366) — converts .xlsm → .xlsx, stripping macros.
5. `wb.Close SaveChanges:=False` (line 367).

The staged copy shares the source's VBA project codename and workbook structure. With the source still open, Excel's COM layer can route `Application.ActiveWorkbook` / sheet-resolution calls to the wrong workbook during the `SaveAs`. Combined with **Protected View** (Excel auto-flags files opened from `%TEMP%`) the SaveAs can produce a partial-but-valid output while corrupting the in-memory state of `ThisWorkbook`. When Excel later asks "save changes?" on close of the source, the user's last-clicked answer can persist the corrupted in-memory state back to the source file on disk — that's the data loss.

### 2. `EnableEvents = False` is never restored on the error path

`Btn_UploadAll`'s `Failed:` handler (lines 144-147) restores `DisplayAlerts` but NOT `EnableEvents`. If anything inside `MakeTempXlsx` throws (Protected View block, antivirus interception, AV-locked `%TEMP%`, etc.), events stay disabled for the rest of the Excel session. Any `Worksheet_Change` / `Workbook_BeforeSave` defenses the operator relies on to PRESERVE data (or even just calc-chain recompute) silently no-op until Excel is fully restarted.

### Why the DB still gets the right data

The `.xlsx` output of step 4 is read by DataViewer immediately (step 5 of `Btn_UploadAll`, lines 127-138) BEFORE the corruption window closes. DataViewer's openpyxl pipeline parses what's on disk at that instant — which is fine. The source-workbook corruption manifests later, when the user closes/reopens the .xlsm.

### Why `MakeTempXlsx` exists at all is now obsolete

The comment at lines 121-122 justifies the .xlsx materialization as "DataViewer's ingestion code is .xlsx-native and macro-free." That stopped being true in commit `81850bc` (v2.0.8): `MainWindow.cpp` now accepts `.xlsm` directly in `detectFileType` (line 1908), the file-open filter (line 2084), and the drag-drop check (line 4394). DataViewer can ingest the .xlsm directly; the entire open-saveAs-close dance is dead weight that introduces the data-loss risk for zero benefit.

## Fix (applied)

`MakeTempXlsx` is rewritten as `MakeTempXlsm`: a single `fso.CopyFile` to `%TEMP%\dvupload_<ts>_<basename>.xlsm`, returned to `Btn_UploadAll`. Excel is not touched. No second workbook is opened. No SaveAs. No macro stripping.

Additional hardening in `Btn_UploadAll`:

1. **Self-overwrite guard** (after line 104): validates that neither `synDest` nor `locDest` resolves to `ThisWorkbook.FullName` (case-insensitive). Aborts with a clear status message if it would. This prevents the rare-but-catastrophic case where `DV_LocalPath` + `DV_FileName` happen to point at the source file itself.
2. **Skip `ThisWorkbook.Save`** is too disruptive to change here, but the surrounding `DisplayAlerts` / `EnableEvents` toggles are wrapped so they always restore on the `Failed:` path (line 145).
3. **Output format**: temp file is now `.xlsm` (matches the user's "format consistency with the input template" requirement).

## Confidence

~70% on root cause #1 (Protected View / COM aliasing during in-place reopen — hard to reproduce deterministically without the operator's exact AV/Trust Center config). ~95% that ELIMINATING the in-place reopen removes the failure mode regardless of the precise mechanism, because there's now no second-workbook open and no SaveAs conversion to misbehave.

---

## Post-upload reset (second data-loss path) — and a correction to the above

**Correction first:** the root-cause #1 analysis above was performed against a
**stale copy** of the module. The repo's `excel-sidecar/DataViewerUpload.bas`
was a 394-line v2.0.8 fork that *predates the reset feature entirely*. The
macros that actually run live in the deployed
`C:\Users\S1134987\Documents\Tempates\Automated Testing Template.xlsm`, whose
`DataViewerUpload` module is 934 lines. The two were never the same file — that
divergence is what made the bug so confusing. The deployed VBA was recovered
read-only with `olevba` (the user's Python is on the MIP allowlist, so it reads
the encrypted `.xlsm` as plaintext).

### Real root cause

The deployed `Btn_UploadAll` ends with a **deterministic** reset of the live
workbook:

```
ResetLiveWorkbookAfterUpload keep      ' step 6
ThisWorkbook.Save                      ' persists whatever step 6 did
```

`ResetLiveWorkbookAfterUpload` → `ResetSheetToTemplate(ws)` builds the name
`"_Template_<SafeSheetName>"` (e.g. `_Template_Lifetime_Test`). **No such
per-sheet template sheets exist** in the workbook (only `_Template_Master`
does), so every sheet falls through to `HardClearSheetData`, which
`.ClearContents` the entire 12-col block rows 5-115 — **including the formula
columns A/B (puff & before-weight chains) and I/J/K/L (TPM etc.)**. That
deletes the formula scaffolding, and `ThisWorkbook.Save` writes the gutted
sheet back to disk. This happens on **every successful upload** — it is the
actual "source `.xlsm` ends up cleared" symptom, not the `MakeTempXlsx` COM
theory (which was a misdiagnosis from reading the wrong file).

Why `_Template_Master` can't just be used as the source: it is
Lifetime-specific (`A1="Lifetime Test"`, puff interval 20, Lifetime row-4
headers). Copying it onto Intense (interval 10) / Negative Pressure (interval
1, `"Smell (1-4)"` headers) etc. would corrupt their title / interval /
headers.

### Fix

`ResetSheetToTemplate` / `TemplateSheetName` / `HardClearSheetData` are replaced
by `ClearSheetEntries` + `BlockCount` — a **surgical clear** that wipes only the
operator-entered cells and leaves every formula, the per-sheet puff interval,
the `A1` title, and the row-4 headers intact. If a block's `A6` is not a
formula (already gutted by the old bug), it is skipped and logged rather than
"reset" — never make damage worse. Cell map:
`docs/superpowers/specs/template-cell-map.md`.

### Consolidation (so this can't recur)

`excel-sidecar/` is now the single source of truth, rebuilt verbatim from the
recovered deployed modules so that **repo == deployed + the reset fix only**
(proven by diff). Added: `TestingTools.bas` (deployed `Module1`, renamed),
`ThisWorkbook.cls.txt`, `customUI14.xml`, `README.md`, `verify_sidecar.py`
(drift detector), `install_sidecar.py` (optional one-shot importer). Removed:
`SampleNav-ThisWorkbook.txt`, `SampleNav-INSTALL.md`,
`SampleNav-ribbon-snippet.xml` (superseded).

### Files modified / added

- `excel-sidecar/DataViewerUpload.bas` — replaced stale fork with deployed
  module + surgical-clear reset.
- `excel-sidecar/TestingTools.bas`, `ThisWorkbook.cls.txt`, `customUI14.xml`,
  `README.md`, `verify_sidecar.py`, `install_sidecar.py` — new canonical
  artifacts / tooling.
- `excel-sidecar/SampleNav.bas` — refreshed to match deployed (identical).
- `docs/superpowers/specs/2026-05-27-sidecar-post-upload-reset.md` — rewritten
  to the corrected plan.

### Out of scope (flagged, not changed)

- The deployed `MakeTempXlsx` still reopens a same-codename copy + SaveAs, but
  it operates on a temp staging copy (not `ThisWorkbook`) and is not the
  data-loss cause.
- `TestingTools.ResetEquations` ("Reset Formulas") imposes interval 20 from
  `_Template_Master` (wrong for non-Lifetime sheets).

### Verification

VBA can't be unit-tested in-process. `verify_sidecar.py` proves repo↔deployed
parity (currently reports `DataViewerUpload` DIFFERS — the drift detector
correctly flagging that the deployed reset is out of date until re-imported).
Manual Excel test recipe is in `excel-sidecar/README.md`.

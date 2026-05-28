# Sidecar Post-Upload Reset — Corrected Plan & Solution

**Date:** 2026-05-27 (revised 2026-05-28)
**Branch:** `feature/xlsm-vba-sidecar` (sidecar worktree)
**Status:** SUPERSEDED 2026-05-28. The "surgical clear" described below
(`ClearSheetEntries`) wiped the `A5` puff seed -- the "row 5 keeps getting
deleted" bug -- and could not reset non-standard sheets like
`Temperature Cycling Test #1`. It is replaced by a whole-sheet revert from
the DataViewer-packaged template:
**`2026-05-28-sidecar-reset-from-template-design.md`**. Kept for history;
the consolidation / single-source-of-truth work below still stands.

> **This supersedes the original plan.** The original plan was written
> against the *deployed* `.xlsm` macros but the fix was attempted against a
> *stale repo copy* that didn't contain the reset code at all. See
> "What actually happened" below.

## What actually happened (the confusion)

Two Claude sessions edited two different things:

- The **deployed** `Automated Testing Template.xlsm` carries a 934-line
  `DataViewerUpload` module (selection-driven visibility, trim, post-upload
  reset, folder pickers) — this is what actually runs.
- The repo's `excel-sidecar/DataViewerUpload.bas` was a **stale 394-line
  fork** (the v2.0.8 baseline, before the reset feature). A prior session
  applied a `MakeTempXlsm` fix *to that stale fork*, while the plan/BUG_REPORT
  described the deployed module. They never referred to the same file.

The deployed VBA was recovered read-only with `olevba` (the user's Python is
on the MIP allowlist, so it reads the encrypted `.xlsm` as plaintext).

## The real bug

`Btn_UploadAll` (deployed) ends with:

```
ResetLiveWorkbookAfterUpload keep      ' step 6
ThisWorkbook.Save                      ' persists whatever step 6 did
```

`ResetLiveWorkbookAfterUpload` → `ResetSheetToTemplate(ws)` computes
`TemplateSheetName("Lifetime Test")` = `"_Template_Lifetime_Test"`. **No
per-sheet `_Template_<X>` sheets exist** (only `_Template_Master` does), so it
always falls through to `HardClearSheetData`, which `.ClearContents` the whole
12-col block rows 5-115 **including the formula columns (A/B chains and
I/J/K/L)** — deleting the formula scaffolding. `ThisWorkbook.Save` then writes
the gutted state back to the live source workbook. Deterministic, every upload.

(The earlier `MakeTempXlsx` COM-aliasing theory was a misdiagnosis caused by
reading the stale fork, which had no reset step.)

## Why not `_Template_Master`

`_Template_Master` is **Lifetime-specific**: `A1="Lifetime Test"`, puff
interval 20 (`A5=20`, `A6==A5+20`), Lifetime row-4 headers. Other sheets
differ (Intense interval 10, Negative Pressure interval 1 + `"Smell (1-4)"` /
`"Clog (Y/N)"` headers). Copying the master onto them would corrupt their
title, interval, and headers. (`Module1.ResetEquations` has the same latent
flaw — see Known issues.)

## The fix — surgical clear

`ResetSheetToTemplate` / `TemplateSheetName` / `HardClearSheetData` are
replaced by `ClearSheetEntries` + `BlockCount`. Per 12-col block, clear **only
the operator-entered cells** and leave every formula, the per-sheet puff
interval, the `A1` title, and the row-4 headers intact:

- Header values: r1 +3 date, +5 sampleID, +7 heatTech, +10 burn; r2 +1 media,
  +3 resistance, +7 puffRegime, +10 clog; r3 +1 viscosity, +3 tester,
  +5 voltage, +7 oilMass, +10 leak.
- Data values: `A5`/`B5` seeds + cols +2..+7 (after-wt / draw / resist / smell
  / clog / notes) rows 5-115.
- **Never touched:** `A6:A115` / `B6:B115` chains, `I/J/K/L` formulas,
  `F2`/`I3`/`L2`/`L3`, all labels.
- **Poka-yoke:** if a block's `A6` isn't a formula (already gutted by the old
  bug), it's skipped and logged rather than "reset" — never make it worse.

Cell map source: `docs/superpowers/specs/template-cell-map.md` (verified
against `_Template_Master` and 3 live sheets).

## Consolidation (single source of truth)

`excel-sidecar/` now holds the canonical, deployable macros, rebuilt verbatim
from the recovered deployed modules (so repo == deployed + the reset fix only,
proven by diff):

- `DataViewerUpload.bas` — deployed 934-line module + surgical-clear reset.
- `TestingTools.bas` — deployed `Module1` (renamed; "Module: TestingTools").
- `SampleNav.bas` — navigation (identical to deployed).
- `ThisWorkbook.cls.txt` — the 4 event handlers to paste into ThisWorkbook.
- `customUI14.xml` — the "TPM Testing" ribbon.
- `README.md` — orientation + install + cell map + known issues.
- `verify_sidecar.py` — extracts the deployed `.xlsm` VBA and diffs it against
  these files (drift detector — the anti-divergence poka-yoke).
- `install_sidecar.py` — optional one-shot importer (dry-run default; backs up
  first; needs "Trust access to the VBA project object model").

Stale files removed: `SampleNav-ThisWorkbook.txt`, `SampleNav-INSTALL.md`,
`SampleNav-ribbon-snippet.xml` (superseded by the canonical copies + README).

## Out of scope (flagged, not changed)

- The deployed `MakeTempXlsx` still reopens a same-codename copy + SaveAs (the
  old COM-aliasing concern). It operates on a temp staging copy, not
  `ThisWorkbook`, and is not the data-loss cause. Left as-is per "the rest is
  good"; documented in README Known issues.
- `Module1.ResetEquations` imposes interval 20 from `_Template_Master`.

## Verification

VBA can't be unit-tested in-process. `verify_sidecar.py` proves repo↔deployed
parity (will report `DataViewerUpload` DIFFERS until the user re-imports —
that's the drift detector working). Manual Excel test recipe is in README.

## Constraints honored

NO commits on this branch. Live operator `.xlsm` never modified (read-only
recovery only). `%USERPROFILE%\SynologyDrive\` untouched. No `src/` / `.pro` /
Qt changes. Export-button WIP isolated on `feature/xlsm-export`.

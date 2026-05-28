# Sample Navigation Sidecar — Design

**Date:** 2026-05-27
**Branch:** `feature/xlsm-vba-sidecar`
**Status:** Approved verbally; deliverables in `excel-sidecar/`.

## Goal

Add per-row sample-block navigation to the operator's
`Automated Testing Template.xlsm` (located at
`C:\Users\S1134987\Documents\Tempates\Automated Testing Template.xlsm`).
Two hotkeys + four ribbon buttons. **No** changes to the DataViewer-Enterprise
Qt code, the `.pro`, or `src/`. Files are produced as artifacts under
`excel-sidecar/` and installed manually — the live `.xlsm` is never touched
by tooling.

## Constants (must match `DataViewerUpload.bas`)

| Constant | Value | Meaning |
|---|---|---|
| `COLS_PER_SAMPLE` | 12 | sample block width |
| `FIRST_SAMPLE_COL` | 3 | column C |
| `MAX_SAMPLES` | 8 | per sheet |
| `HEADER_ROW` | 4 | header row |
| `LAST_SAMPLE_COL` | 87 (= 3 + 7·12) | column CI |

Valid sample columns: `{C, O, AA, AM, AY, BK, BW, CI}`.

## Macros (`excel-sidecar/SampleNav.bas`)

| Macro | Behavior | Clamp |
|---|---|---|
| `JumpRight12` | `ActiveCell.Column + 12`, same row | At `LAST_SAMPLE_COL` |
| `JumpLeft12`  | `ActiveCell.Column − 12`, same row | At `FIRST_SAMPLE_COL` |
| `JumpFirstSample` | `Cells(ActiveCell.Row, 3).Select` | n/a |
| `JumpLastSample`  | Scan header row at sample cols right-to-left; jump to first non-empty header on current row | Do nothing if none |

Each main macro is no-arg (so `Application.OnKey` can call it). A 1-line
`_Ribbon(control As Object)` wrapper per macro satisfies the customUI
`onAction` signature requirement and delegates to the main macro.

`SafeActiveCell()` helper returns `Nothing` on chart sheets / protected
view / empty workbook; macros early-exit so no error dialogs fire.

## Hotkeys (`excel-sidecar/SampleNav-ThisWorkbook.txt`)

| Combo | Macro | OnKey code |
|---|---|---|
| Ctrl+Shift+. | `JumpRight12` | `"+^."` |
| Ctrl+Shift+, | `JumpLeft12`  | `"+^,"` |

Registered in `Workbook_Open`, restored to Excel defaults in
`Workbook_BeforeClose` (so the bindings don't leak into other workbooks
in the same Excel session).

First/Last get buttons only — no hotkey, since Ctrl+Shift+Home/End are
critical Excel shortcuts we won't clobber.

## Ribbon group (`excel-sidecar/SampleNav-ribbon-snippet.xml`)

New `<group id="grpSampleNav" label="Sample Navigation">` inserted into
the existing `<tab>` for "TPM Testing" in the workbook's `customUI14.xml`.
Four large buttons, plain text labels (no unicode glyphs — render
inconsistently across Excel versions):

| Label | imageMso | onAction |
|---|---|---|
| First | `GoToFirstPage` | `JumpFirstSample_Ribbon` |
| Prev Sample | `GoBack` | `JumpLeft12_Ribbon` |
| Next Sample | `GoForward` | `JumpRight12_Ribbon` |
| Last | `GoToLastPage` | `JumpLastSample_Ribbon` |

## Install (`excel-sidecar/SampleNav-INSTALL.md`)

Step-by-step, with troubleshooting and uninstall.

Primary path: Office RibbonX Editor for the XML, VBE for `.bas` import +
ThisWorkbook paste. Unzip-method fallback documented for users without
the Custom UI Editor.

## Scope guarantees

- Tooling does not open, read, or write the live `.xlsm`.
- No changes to `src/`, `.pro`, Qt code, or any DataViewer software.
- No writes to `%USERPROFILE%\SynologyDrive\`.
- Files land only in `excel-sidecar/` (artifacts) and
  `docs/superpowers/specs/` (this doc). No git commit until requested.

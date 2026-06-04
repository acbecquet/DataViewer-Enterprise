# Excel Sidecar Clean Rebuild — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Author the corrected `excel-sidecar/` sources (crash/save-loop fixes, non-destructive reset that keeps a persistent `- Review` copy, ribbon consolidation incl. Delete All Review Sheets) plus an Excel-driven `build_clean_template.py`, all gated by headless checks; then hand off the Excel build + acceptance to the work machine.

**Architecture:** The repo `excel-sidecar/` is the single source of truth. We edit the canonical `.bas`/`.cls.txt`/`customUI14.xml` in this worktree and verify them headlessly with a new `check_sources.py` (invariant checker) + the extended `verify_sidecar.py`. A new `build_clean_template.py` (pywin32) builds a brand-new workbook from those sources on the work machine — so the rebuilt `.xlsm` matches the repo by construction (drift resolved). VBA can't run here; behavioral acceptance is a documented operator checklist.

**Tech Stack:** VBA (Excel macros), Office customUI (ribbon XML), Python 3.13 (MIP-allowlisted, with `oletools`/`olevba`), pywin32 (work machine only), `zipfile` for OOXML surgery.

**Spec:** `docs/superpowers/specs/2026-06-04-excel-sidecar-clean-rebuild.md`

**Run Python with:** `"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe"` (alias `PY` below).
**Worktree root (cwd for all commands):** `C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/excel-sidecar-rebuild`
**Live source workbook (`SRC`):** `C:/Users/S1134987/Documents/Templates/Automated Testing Template v1.xlsm`

---

## File Structure

| File | Responsibility |
|---|---|
| `excel-sidecar/check_sources.py` | **NEW.** Headless invariant checker over the VBA/ribbon sources (crash fixes, Review/Delete-All procedures, onAction↔handler cross-check, Review-name table). The executable spec we drive to green. |
| `excel-sidecar/DataViewerUpload.bas` | **MODIFY.** Reset→Review (`ResetSheetToBlankWithReview`), `DeleteAllReviewSheets`, naming helpers (`ReviewBaseName`/`UniqueReviewName`/`IsReviewSheet`), ribbon wrappers. Remove `RestoreSheetFromTemplate`. |
| `excel-sidecar/TestingTools.bas` | **MODIFY.** Add crash-safe `TryPuffStepPicker`. |
| `excel-sidecar/ThisWorkbook.cls.txt` | **MODIFY.** Thin dispatch; drop dead double-click handler; unconditional `Saved=True` at end of Open. |
| `excel-sidecar/SampleNav.bas` | **UNCHANGED.** |
| `excel-sidecar/customUI14.xml` | **MODIFY.** Add 3rd group "DataViewer Upload" (5 buttons). |
| `excel-sidecar/verify_sidecar.py` | **MODIFY.** Add ribbon-match + no-web-add-in checks. |
| `excel-sidecar/build_clean_template.py` | **NEW.** pywin32 Excel-driven clean rebuild + `inject_customui`/`check_preconditions` (headless-testable helpers). |
| `excel-sidecar/test_build_helpers.py` | **NEW.** Headless tests for `inject_customui` + `check_preconditions`. |
| `excel-sidecar/README.md`, `excel-sidecar/RUNBOOK-migrate-existing-template.md` | **MODIFY.** Rebuild flow + Review feature + operator acceptance checklist. |

> Note: all `excel-sidecar/*.bas`/`.cls.txt`/`.xml` are plain text in this worktree (outside `Documents`, so no MIP ciphertext). Edit them directly with Edit/Write.

---

## Task 1: Headless invariant checker (`check_sources.py`)

Write the executable spec first. It will FAIL against the current sources (the new procedures/group don't exist yet); Tasks 2–5 drive it to green.

**Files:**
- Create: `excel-sidecar/check_sources.py`

- [ ] **Step 1: Write the checker**

Create `excel-sidecar/check_sources.py` with exactly:

```python
#!/usr/bin/env python3
"""Headless invariant checks for the excel-sidecar VBA/ribbon SOURCES.

Validates design invariants checkable without Excel: the crash-fix structure,
the Review/Delete-All procedures, the Review-name table, and that every ribbon
onAction has a matching VBA handler. Run with any Python.

    python excel-sidecar/check_sources.py
Exit 0 if all invariants hold, 1 otherwise.
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))


def rd(name):
    return open(os.path.join(HERE, name), encoding="utf-8").read()


results = []


def check(name, cond, detail=""):
    results.append((bool(cond), name, detail))


dvu = rd("DataViewerUpload.bas")
tt = rd("TestingTools.bas")
twb = rd("ThisWorkbook.cls.txt")
sn = rd("SampleNav.bas")
ui = rd("customUI14.xml")

# --- ThisWorkbook: crash / save-loop fixes ---
check("ThisWorkbook drops dead double-click handler",
      "Workbook_SheetBeforeDoubleClick" not in twb)
check("Workbook_Open sets ThisWorkbook.Saved=True",
      re.search(r"Workbook_Open\(\).*?ThisWorkbook\.Saved\s*=\s*True.*?End Sub",
                twb, re.S) is not None)
check("SheetChange dispatches picker then DataViewer handler",
      re.search(r"TestingTools\.TryPuffStepPicker\(Sh,\s*Target\)", twb) is not None
      and re.search(r"DataViewerUpload\.OnWorkbookSheetChange\s+Sh,\s*Target", twb)
      is not None)

# --- TestingTools: crash-safe picker ---
check("TestingTools defines TryPuffStepPicker",
      re.search(r"Public\s+Function\s+TryPuffStepPicker\b", tt) is not None)
m = re.search(r"Public\s+Function\s+TryPuffStepPicker\b.*?End Function", tt, re.S)
pbody = m.group(0) if m else ""
check("Picker guarantees EnableEvents restore (On Error GoTo PuffDone + label)",
      "On Error GoTo PuffDone" in pbody
      and re.search(r"PuffDone:\s*\r?\n\s*Application\.EnableEvents\s*=\s*savedEvents",
                    pbody) is not None)
check("Picker disables events exactly once and restores via saved state",
      pbody.count("Application.EnableEvents = False") == 1
      and "Application.EnableEvents = savedEvents" in pbody)

# --- DataViewerUpload: Review + reset + Delete-All + ribbon wrappers ---
for token in ["Public Sub DeleteAllReviewSheets",
              "Private Function ReviewBaseName",
              "Private Function UniqueReviewName",
              "Private Function IsReviewSheet",
              "Private Sub ResetSheetToBlankWithReview"]:
    check("DataViewerUpload defines `%s`" % token.split()[-1], token in dvu)
check("ResetLiveWorkbookAfterUpload calls ResetSheetToBlankWithReview",
      "ResetSheetToBlankWithReview ThisWorkbook" in dvu)
check("Old RestoreSheetFromTemplate removed",
      "Sub RestoreSheetFromTemplate" not in dvu)
for w in ["Ribbon_UploadAll", "Ribbon_DryRun", "Ribbon_PickSynology",
          "Ribbon_PickLocal", "Ribbon_DeleteReviewSheets"]:
    check("DataViewerUpload defines wrapper `%s`" % w,
          re.search(r"Public\s+Sub\s+%s\s*\(\s*control\s+As\s+IRibbonControl" % w,
                    dvu) is not None)

# --- Review base-name table: length + width vs CanonicalDataSheets ---
def vba_array(text, func):
    mm = re.search(r"Function\s+%s\b.*?Array\((.*?)\)" % func, text, re.S)
    return re.findall(r'"([^"]*)"', mm.group(1)) if mm else []

canon = vba_array(dvu, "CanonicalDataSheets")
rbase = vba_array(dvu, "ReviewBaseName")
check("ReviewBaseName has one entry per canonical sheet (12)",
      len(rbase) == len(canon) == 12, "canon=%d rbase=%d" % (len(canon), len(rbase)))
toolong = [b for b in rbase if len(b) + len(" - Review") > 31]
check("Every Review base fits '<base> - Review' in 31 chars", not toolong, str(toolong))

# --- Ribbon onAction <-> VBA handler cross-check ---
handlers = set(re.findall(r"Public\s+Sub\s+(\w+)\s*\(\s*control\b",
                          "\n".join([dvu, tt, sn])))
actions = set(re.findall(r'onAction="([^"]+)"', ui))
missing = sorted(actions - handlers)
check("Every ribbon onAction has a matching VBA handler", not missing,
      "missing=%s" % missing)
check("customUI has the 3rd 'DataViewer Upload' group",
      'label="DataViewer Upload"' in ui)

# --- report ---
ok = True
for passed, name, detail in results:
    line = ("PASS " if passed else "FAIL ") + name
    if detail and not passed:
        line += "  [%s]" % detail
    print(line)
    ok = ok and passed
print("-" * 60)
print("RESULT:", "ALL PASS" if ok else "FAILURES")
sys.exit(0 if ok else 1)
```

- [ ] **Step 2: Run it — expect FAILURES (baseline)**

Run: `cd <worktree> && PY excel-sidecar/check_sources.py`
Expected: several `FAIL` lines (no `DeleteAllReviewSheets`, no `TryPuffStepPicker`, no 3rd group, etc.) and `RESULT: FAILURES` (exit 1). This confirms the checker detects the missing work.

- [ ] **Step 3: Commit**

```bash
git add excel-sidecar/check_sources.py
git commit -m "test(sidecar): headless invariant checker for the rebuild sources"
```

---

## Task 2: `DataViewerUpload.bas` — Review reset, Delete-All, naming, ribbon wrappers

**Files:**
- Modify: `excel-sidecar/DataViewerUpload.bas`

- [ ] **Step 1: Replace the reset with the non-destructive Review reset**

In `ResetLiveWorkbookAfterUpload`, change the per-sheet call from
`RestoreSheetFromTemplate ThisWorkbook, sheetName, i`
to
`ResetSheetToBlankWithReview ThisWorkbook, sheetName, i`.

Then **delete the entire `Private Sub RestoreSheetFromTemplate(...)` procedure** and insert this in its place:

```vba
Private Sub ResetSheetToBlankWithReview(liveWb As Workbook, sheetName As String, idx As Long)
    ' Non-destructive reset: the live data sheet BECOMES a "<base> - Review" copy
    ' (a rename, not a delete), and a pristine blank from the internal snapshot
    ' (_Template_NN) takes its place + tab position. Transactional: on any failure
    ' the sheet is renamed back to its canonical name and left fully intact.
    Dim tplName As String
    tplName = TemplateSheetName(idx)
    If Not WorkbookHasSheetIn(liveWb, tplName) Then
        StampLog "  '" & sheetName & "': no internal snapshot (" & tplName & _
                 ") - not reset (left intact). Build snapshots first (RebuildBlankTemplates)."
        Exit Sub
    End If
    If Not SheetExists(sheetName) Then Exit Sub

    Dim orig As Worksheet
    Set orig = liveWb.Worksheets(sheetName)
    Dim reviewName As String
    reviewName = UniqueReviewName(idx)

    Dim savedAlerts As Boolean
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Fail

    ' 1) Rename the live sheet into its Review copy (data/format/formulas intact).
    orig.Name = reviewName

    ' 2) Copy the snapshot in at the Review sheet's (former canonical) position.
    Dim tpl As Worksheet
    Set tpl = liveWb.Worksheets(tplName)
    Dim savedVis As XlSheetVisibility
    savedVis = tpl.Visible
    tpl.Visible = xlSheetVisible

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim w As Worksheet
    For Each w In liveWb.Worksheets
        seen(w.Name) = True
    Next
    tpl.Copy Before:=orig                 ' lands just before the (renamed) Review sheet
    Dim fresh As Worksheet
    For Each w In liveWb.Worksheets
        If Not seen.Exists(w.Name) Then
            Set fresh = w
            Exit For
        End If
    Next
    tpl.Visible = savedVis

    ' POKA-YOKE: if the copy added nothing, undo the rename and bail (intact).
    If fresh Is Nothing Then
        orig.Name = sheetName
        StampLog "  '" & sheetName & "': snapshot copy added no sheet - reset skipped, left intact."
        Application.DisplayAlerts = savedAlerts
        Exit Sub
    End If

    fresh.Name = sheetName
    fresh.Visible = xlSheetVisible        ' ApplySheetVisibility sets the real state

    ' 3) Move the Review sheet to the end to keep the working area uncluttered.
    orig.Move After:=liveWb.Worksheets(liveWb.Worksheets.Count)

    Application.DisplayAlerts = savedAlerts
    StampLog "  '" & sheetName & "': reset; data preserved as '" & reviewName & "'."
    Exit Sub

Fail:
    On Error Resume Next
    If SheetExists(reviewName) And Not SheetExists(sheetName) Then
        liveWb.Worksheets(reviewName).Name = sheetName
    End If
    Application.DisplayAlerts = savedAlerts
    On Error GoTo 0
    StampLog "  '" & sheetName & "': reset FAILED (" & Err.Description & ") - left as-is."
End Sub
```

- [ ] **Step 2: Add the Review-naming + detection helpers**

Insert near the other utilities (e.g. just below `TemplateSheetName`):

```vba
Private Function ReviewBaseName(idx As Long) As String
    ' Curated <=20-char labels so "<base> - Review [N]" stays within Excel's
    ' 31-char sheet-name limit. MUST stay parallel to CanonicalDataSheets().
    Dim names As Variant
    names = Array( _
        "Lifetime Test", _
        "User Test Simulation", _
        "Long Puff Lifetime", _
        "Rapid Puff Lifetime", _
        "Intense Test", _
        "Big Headspace Serial", _
        "Negative Pressure", _
        "Temp Cycling Test #2", _
        "Viscosity Compat", _
        "Various Oil Compat", _
        "Custom Test Template", _
        "Temp Cycling Test #1")
    If idx >= LBound(names) And idx <= UBound(names) Then
        ReviewBaseName = CStr(names(idx))
    Else
        ReviewBaseName = "Sheet " & (idx + 1)
    End If
End Function

Private Function UniqueReviewName(idx As Long) As String
    Dim base As String
    base = ReviewBaseName(idx)
    Dim candidate As String
    candidate = base & " - Review"
    If Len(candidate) <= 31 And Not SheetExists(candidate) Then
        UniqueReviewName = candidate
        Exit Function
    End If
    Dim n As Long, suffix As String, b As String
    For n = 2 To 99
        suffix = " - Review " & n
        b = base
        If Len(b) + Len(suffix) > 31 Then b = Left$(b, 31 - Len(suffix))
        candidate = b & suffix
        If Not SheetExists(candidate) Then
            UniqueReviewName = candidate
            Exit Function
        End If
    Next n
    UniqueReviewName = Left$(base, 18) & " - Rev " & idx   ' near-impossible fallback
End Function

Private Function IsReviewSheet(ws As Worksheet) As Boolean
    ' A sheet is a Review copy iff its name contains " - Review" AND it is not a
    ' canonical/utility/template sheet (hard guard against false positives).
    Dim nm As String
    nm = ws.Name
    If InStr(1, nm, " - Review", vbTextCompare) = 0 Then Exit Function
    If StrComp(nm, UPLOAD_SHEET_NAME, vbTextCompare) = 0 Then Exit Function
    If StrComp(nm, SOPS_SHEET_NAME, vbTextCompare) = 0 Then Exit Function
    If Left$(nm, 10) = "_Template_" Then Exit Function
    If Left$(nm, 6) = "_Macro" Then Exit Function
    Dim s As Variant
    For Each s In CanonicalDataSheets()
        If StrComp(NormalizeSheetName(nm), NormalizeSheetName(CStr(s)), vbTextCompare) = 0 Then
            Exit Function
        End If
    Next
    IsReviewSheet = True
End Function
```

- [ ] **Step 3: Add `DeleteAllReviewSheets`**

Insert in the public-button-handlers area (e.g. just below `Btn_UploadAll`):

```vba
Public Sub DeleteAllReviewSheets()
    Dim victims As Collection
    Set victims = New Collection
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsReviewSheet(ws) Then victims.Add ws.Name
    Next
    If victims.Count = 0 Then
        MsgBox "There are no Review sheets to delete.", vbInformation, "Delete All Review Sheets"
        Exit Sub
    End If
    If MsgBox("Delete " & victims.Count & " Review sheet(s)? This cannot be undone.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Delete All Review Sheets") <> vbYes Then
        Exit Sub
    End If

    Dim savedEvents As Boolean, savedAlerts As Boolean, savedSU As Boolean
    savedEvents = Application.EnableEvents
    savedAlerts = Application.DisplayAlerts
    savedSU = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    On Error Resume Next
    ThisWorkbook.Worksheets(UPLOAD_SHEET_NAME).Activate   ' never delete the active view's last sheet
    On Error GoTo Cleanup

    Dim nm As Variant, deleted As Long
    deleted = 0
    For Each nm In victims
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(nm)).Delete
        If Err.Number = 0 Then deleted = deleted + 1
        Err.Clear
        On Error GoTo Cleanup
    Next

Cleanup:
    Application.EnableEvents = savedEvents
    Application.DisplayAlerts = savedAlerts
    Application.ScreenUpdating = savedSU
    On Error GoTo 0
    MsgBox "Deleted " & deleted & " Review sheet(s).", vbInformation, "Delete All Review Sheets"
End Sub
```

- [ ] **Step 4: Add the ribbon wrappers**

Append at the end of the module (after the path-picker handlers):

```vba
' ============================================================================
' Ribbon callbacks (TPM Testing tab -> "DataViewer Upload" group)
' ============================================================================
Public Sub Ribbon_UploadAll(control As IRibbonControl)
    Btn_UploadAll
End Sub
Public Sub Ribbon_DryRun(control As IRibbonControl)
    Btn_DryRunChecklist
End Sub
Public Sub Ribbon_PickSynology(control As IRibbonControl)
    Btn_PickSynologyFolder
End Sub
Public Sub Ribbon_PickLocal(control As IRibbonControl)
    Btn_PickLocalFolder
End Sub
Public Sub Ribbon_DeleteReviewSheets(control As IRibbonControl)
    DeleteAllReviewSheets
End Sub
```

- [ ] **Step 5: Run the checker — DataViewerUpload checks now PASS**

Run: `cd <worktree> && PY excel-sidecar/check_sources.py`
Expected: all `DataViewerUpload …`, `ReviewBaseName …`, and `Review base fits …` lines now `PASS`. Picker/ThisWorkbook/ribbon-group checks still `FAIL` (Tasks 3–5). The onAction↔handler check may still `FAIL` until the ribbon group exists (Task 5) — fine.

- [ ] **Step 6: Commit**

```bash
git add excel-sidecar/DataViewerUpload.bas
git commit -m "feat(sidecar): non-destructive reset (Review copy), Delete-All, ribbon wrappers"
```

---

## Task 3: `TestingTools.bas` — crash-safe puff-step-picker

**Files:**
- Modify: `excel-sidecar/TestingTools.bas`

- [ ] **Step 1: Add `TryPuffStepPicker`**

Insert just above the `' Ribbon callbacks` line near the bottom of the module:

```vba
Public Function TryPuffStepPicker(Sh As Object, Target As Range) As Boolean
    ' Auto-fills a sample block's puff column when the user types a step value
    ' (1/2/5/10/20/50) or "custom" in the puffs column (each block's 1st column,
    ' header row 4 = "puffs"), rows 5..BLOCK_ROWS. Returns True if it handled the
    ' edit. CRASH-SAFE: Application.EnableEvents is ALWAYS restored at PuffDone,
    ' so a mid-edit error can never leave events disabled for the session.
    TryPuffStepPicker = False
    If Sh Is Nothing Then Exit Function
    If TypeName(Sh) <> "Worksheet" Then Exit Function

    Dim pick As Range
    On Error Resume Next
    Set pick = Application.Intersect(Target, Sh.UsedRange)
    On Error GoTo 0
    If pick Is Nothing Then Exit Function
    If pick.Count <> 1 Then Exit Function

    Dim c As Long
    c = pick.Column
    If ((c - 1) Mod BLOCK_COLS) <> 0 Then Exit Function           ' not a block's first column
    If pick.row < 5 Or pick.row > BLOCK_ROWS Then Exit Function
    If LCase$(Trim$(CStr(Sh.Cells(4, c).value))) <> "puffs" Then Exit Function

    Dim v As String
    v = LCase$(Trim$(CStr(pick.value)))
    Dim stp As Variant
    stp = Empty
    Select Case v
        Case "1": stp = 1
        Case "2": stp = 2
        Case "5": stp = 5
        Case "10": stp = 10
        Case "20": stp = 20
        Case "50": stp = 50
        Case "custom": stp = "custom"
        Case Else: Exit Function          ' nothing to do - let other handlers run
    End Select

    Dim savedEvents As Boolean
    savedEvents = Application.EnableEvents
    On Error GoTo PuffDone
    Application.EnableEvents = False

    If VarType(stp) = vbString Then        ' "custom"
        pick.ClearContents
        pick.Select
    ElseIf pick.row = 5 Then
        Sh.Cells(5, c).value = stp         ' seed row; rows below chain off it
    Else
        Dim colL As String
        colL = ColLetter(c)
        Dim rr As Long
        For rr = pick.row To BLOCK_ROWS
            Sh.Cells(rr, c).Formula = "=" & colL & (rr - 1) & "+" & stp
        Next rr
    End If
    TryPuffStepPicker = True

PuffDone:
    Application.EnableEvents = savedEvents
    On Error GoTo 0
End Function
```

- [ ] **Step 2: Run the checker — picker checks PASS**

Run: `cd <worktree> && PY excel-sidecar/check_sources.py`
Expected: `TestingTools defines TryPuffStepPicker`, `Picker guarantees EnableEvents restore …`, and `Picker disables events exactly once …` now `PASS`.

- [ ] **Step 3: Commit**

```bash
git add excel-sidecar/TestingTools.bas
git commit -m "feat(sidecar): crash-safe puff-step-picker (guaranteed EnableEvents restore)"
```

---

## Task 4: `ThisWorkbook.cls.txt` — thin dispatch, drop dead handler, Saved guard

**Files:**
- Modify: `excel-sidecar/ThisWorkbook.cls.txt`

- [ ] **Step 1: Replace the body** (keep the leading comment header)

Replace everything from `Option Explicit` to end of file with:

```vba
Option Explicit

Private Sub Workbook_Open()
    DataViewerUpload.ApplySheetVisibility
    Application.OnKey "+^.", "JumpRight12"
    Application.OnKey "+^,", "JumpLeft12"
    ' Open does only presentational work (visibility, hotkeys), all re-derived
    ' on the next open - so don't nag to save a no-edit session on close.
    ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If TestingTools.TryPuffStepPicker(Sh, Target) Then Exit Sub
    DataViewerUpload.OnWorkbookSheetChange Sh, Target
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Restore Excel default behavior so the bindings don't leak into other
    ' workbooks opened in the same Excel session.
    On Error Resume Next
    Application.OnKey "+^."
    Application.OnKey "+^,"
End Sub
```

(The old `Workbook_SheetBeforeDoubleClick` is intentionally gone — the selection now uses TRUE/FALSE data-validation dropdowns at rows 3–16, not a double-click toggle at rows 20–33.)

- [ ] **Step 2: Run the checker — ThisWorkbook checks PASS**

Run: `cd <worktree> && PY excel-sidecar/check_sources.py`
Expected: `ThisWorkbook drops dead double-click handler`, `Workbook_Open sets ThisWorkbook.Saved=True`, and `SheetChange dispatches …` now `PASS`.

- [ ] **Step 3: Commit**

```bash
git add excel-sidecar/ThisWorkbook.cls.txt
git commit -m "fix(sidecar): thin ThisWorkbook - drop dead handler, no-nag Saved, picker dispatch"
```

---

## Task 5: `customUI14.xml` — add the "DataViewer Upload" ribbon group

**Files:**
- Modify: `excel-sidecar/customUI14.xml`

- [ ] **Step 1: Insert the 3rd group** before `</tab>` (immediately after the closing `</group>` of `grpSampleNav`):

```xml
        <group id="grpDvUpload" label="DataViewer Upload">
          <button id="btnUploadAll" label="Upload All" size="large" imageMso="PublishToWebSite"
                  onAction="Ribbon_UploadAll"
                  screentip="Validate, upload, and reset"
                  supertip="Runs the checklist, copies the data to Synology + Local + DataViewer, then resets each uploaded sheet (a '- Review' copy is kept)."/>
          <button id="btnDryRun" label="Dry-Run Checklist" size="large" imageMso="FileValidation"
                  onAction="Ribbon_DryRun"
                  screentip="Validate without uploading"
                  supertip="Runs the upload checklist and reports issues without copying or resetting anything."/>
          <button id="btnPickSyn" label="Pick Synology Folder" size="normal" imageMso="FolderOpen"
                  onAction="Ribbon_PickSynology"
                  screentip="Choose the Synology data folder"/>
          <button id="btnPickLoc" label="Pick Local Folder" size="normal" imageMso="FolderOpen"
                  onAction="Ribbon_PickLocal"
                  screentip="Choose the local data folder"/>
          <button id="btnDelReviews" label="Delete All Review Sheets" size="large" imageMso="Delete"
                  onAction="Ribbon_DeleteReviewSheets"
                  screentip="Delete every '- Review' sheet"
                  supertip="One click removes all '- Review' sheets. Canonical, template, and utility sheets are never touched."/>
        </group>
```

- [ ] **Step 2: Run the checker — EVERYTHING PASS**

Run: `cd <worktree> && PY excel-sidecar/check_sources.py`
Expected: every line `PASS`, including `Every ribbon onAction has a matching VBA handler` and `customUI has the 3rd 'DataViewer Upload' group`; final `RESULT: ALL PASS` (exit 0).

- [ ] **Step 3: Commit**

```bash
git add excel-sidecar/customUI14.xml
git commit -m "feat(sidecar): ribbon 'DataViewer Upload' group (Upload/Dry-Run/pickers/Delete-Reviews)"
```

---

## Task 6: Extend `verify_sidecar.py` — ribbon match + no-web-add-in

**Files:**
- Modify: `excel-sidecar/verify_sidecar.py`

- [ ] **Step 1: Add a ribbon + add-in check, wired into `main`**

Add these helpers above `main()`:

```python
def _norm_xml(s):
    import re as _re
    return _re.sub(r"\s+", " ", s.replace("\r\n", "\n")).strip()


def check_ribbon_and_addin(xlsm_path, repo_dir):
    """Returns (any_diff, lines). Compares the workbook's customUI14.xml to the
    repo copy (normalized) and asserts no web-extension (add-in) parts remain."""
    import zipfile
    lines, any_diff = [], False
    try:
        with zipfile.ZipFile(xlsm_path) as z:
            names = z.namelist()
            wb_ui = z.read("customUI/customUI14.xml").decode("utf-8") \
                if "customUI/customUI14.xml" in names else ""
            webext = [n for n in names if n.startswith("xl/webextensions/")]
    except Exception as e:  # pragma: no cover
        return True, ["[!] could not read workbook zip: %r" % e]
    repo_ui = open(os.path.join(repo_dir, "customUI14.xml"), encoding="utf-8").read()
    if not wb_ui:
        any_diff = True
        lines.append("[!] customUI14.xml: MISSING from workbook")
    elif _norm_xml(wb_ui) == _norm_xml(repo_ui):
        lines.append("[OK]      customUI14.xml == repo")
    else:
        any_diff = True
        lines.append("[DIFFERS] customUI14.xml != repo")
    if webext:
        any_diff = True
        lines.append("[!] web add-in parts present: %s" % ", ".join(webext))
    else:
        lines.append("[OK]      no web-extension/add-in parts")
    return any_diff, lines
```

Then, inside `main()`, just before the final `print("RESULT:", ...)`, add:

```python
    rib_diff, rib_lines = check_ribbon_and_addin(args.file, args.repo)
    print("-" * 60)
    for l in rib_lines:
        print(l)
    any_diff = any_diff or rib_diff
```

- [ ] **Step 2: Run against the live source — expect DIFFERS (it has the OLD 2-group ribbon)**

Run: `cd <worktree> && PY excel-sidecar/verify_sidecar.py --file "C:/Users/S1134987/Documents/Templates/Automated Testing Template v1.xlsm"`
Expected: module lines as before, then `[DIFFERS] customUI14.xml != repo` (the live ribbon lacks the new group) and `[!] web add-in parts present: xl/webextensions/...`, ending `RESULT: DRIFT DETECTED`. This proves the new checks fire correctly. (After the work-machine rebuild, the same command against the rebuilt file should report all `[OK]` / `all modules match`.)

- [ ] **Step 3: Commit**

```bash
git add excel-sidecar/verify_sidecar.py
git commit -m "test(sidecar): verify_sidecar also checks ribbon match + absence of web add-in"
```

---

## Task 7: `build_clean_template.py` + headless helper tests

**Files:**
- Create: `excel-sidecar/build_clean_template.py`
- Create: `excel-sidecar/test_build_helpers.py`

- [ ] **Step 1: Write `build_clean_template.py`**

```python
#!/usr/bin/env python3
"""Build a CLEAN Automated Testing Template .xlsm from a (possibly messy)
source, on the WORK MACHINE via Excel COM (pywin32).

Creates a brand-new workbook (fresh package: no inherited corruption, no web
add-in, no calcChain rot), copies the source's sheets in with full fidelity,
stamps canonical sheets blank from their snapshots, imports the canonical VBA
from this folder, removes the on-sheet upload buttons, saves, then injects the
ribbon at the zip level.

    python excel-sidecar/build_clean_template.py \
        --source "C:\\...\\Automated Testing Template v1.xlsm" \
        --out    "C:\\...\\Automated Testing Template v1 (clean).xlsm"

Requires Windows + Excel + pywin32, and Excel Trust Center ->
"Trust access to the VBA project object model" enabled.
"""
import argparse
import os
import re
import shutil
import sys
import zipfile

REPO = os.path.dirname(os.path.abspath(__file__))
CANON = [
    "Lifetime Test", "User Test Simulation", "Long Puff Lifetime Test",
    "Rapid Puff Lifetime Test", "Intense Test", "Big Headspace Serial Test",
    "Negative Pressure Test", "Temperature Cycling Test #2",
    "Viscosity Compatibility", "Various Oil Compatibility",
    "Custom Test Template", "Temperature Cycling Test #1",
]
SNAPSHOTS = ["_Template_%02d" % i for i in range(12)]
KEEP = ["Test SOP's", "DataViewer Upload"] + CANON + ["_Template_Master"] + SNAPSHOTS
NAMED = {  # name -> cell ref on the DataViewer Upload sheet
    "DV_FileName": "$I$6", "DV_SynologyPath": "$I$8", "DV_LocalPath": "$I$10",
    "DV_Status": "$H$16", "DV_Log": "$H$17", "DV_DataViewerExe": "$I$12",
    "DV_TestSelection": "$A$3:$B$16",
}
UPLOAD_SHEET = "DataViewer Upload"
XL_OPENXML_MACRO = 52   # xlOpenXMLWorkbookMacroEnabled (.xlsm)
XL_VERYHIDDEN = 2
XL_HIDDEN = 0
XL_VISIBLE = -1


def check_preconditions(source):
    problems = []
    if not os.path.isfile(source):
        problems.append("source not found: %s" % source)
    try:
        import win32com.client  # noqa: F401
    except Exception:
        problems.append("pywin32 not installed (run: pip install pywin32)")
    return problems


def _grab(pattern, text, label):
    m = re.search(pattern, text)
    if not m:
        raise RuntimeError("could not find %s in source package" % label)
    return m.group(0)


def inject_customui(target_xlsm, repo_dir, scaffold_src):
    """Add the ribbon to a saved .xlsm by lifting the proven customUI wiring
    from `scaffold_src` (a workbook that already has a working customUI14) and
    swapping in the repo customUI14.xml + icons. Pure zip surgery (no Excel).
    Idempotent. Raises if a web add-in survives in the output."""
    new_ui = open(os.path.join(repo_dir, "customUI14.xml"), "rb").read()
    with zipfile.ZipFile(scaffold_src) as zs:
        sn = set(zs.namelist())
        src_root_rels = zs.read("_rels/.rels").decode("utf-8")
        src_ctypes = zs.read("[Content_Types].xml").decode("utf-8")
        ui_rels = zs.read("customUI/_rels/customUI14.xml.rels") \
            if "customUI/_rels/customUI14.xml.rels" in sn else None
        images = {n: zs.read(n) for n in sn if n.startswith("customUI/images/")}
    # The exact <Relationship .../> for customUI14 and the customUI Override +
    # png Default, copied verbatim from a workbook that already works.
    rel = _grab(r'<Relationship[^>]*customUI/customUI14\.xml[^>]*/>',
                src_root_rels, "customUI relationship")
    override = _grab(r'<Override[^>]*customUI/customUI14\.xml[^>]*/>',
                     src_ctypes, "customUI content-type override")
    png_default = _grab(r'<Default[^>]*Extension="png"[^>]*/>',
                        src_ctypes, "png content-type default")

    tmp = target_xlsm + ".tmp"
    with zipfile.ZipFile(target_xlsm) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            n = item.filename
            if n in ("customUI/customUI14.xml",):
                continue  # re-added below
            data = zin.read(n)
            if n == "_rels/.rels":
                s = data.decode("utf-8")
                if "customUI/customUI14.xml" not in s:
                    s = s.replace("</Relationships>", rel + "</Relationships>")
                data = s.encode("utf-8")
            elif n == "[Content_Types].xml":
                s = data.decode("utf-8")
                if "customUI/customUI14.xml" not in s:
                    s = s.replace("</Types>", override + "</Types>")
                if 'Extension="png"' not in s:
                    s = s.replace("</Types>", png_default + "</Types>")
                data = s.encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("customUI/customUI14.xml", new_ui)
        if ui_rels is not None:
            zout.writestr("customUI/_rels/customUI14.xml.rels", ui_rels)
        for n, b in images.items():
            zout.writestr(n, b)
    os.replace(tmp, target_xlsm)

    with zipfile.ZipFile(target_xlsm) as z:
        out = z.namelist()
    assert "customUI/customUI14.xml" in out, "customUI not injected"
    assert not any(n.startswith("xl/webextensions/") for n in out), \
        "web add-in parts present in output"


def build(source, out):
    import win32com.client
    backup = source + ".bak"
    shutil.copyfile(source, backup)
    print("Backed up source ->", backup)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable (no macro prompts)
    try:
        # VBA object-model trust is required to import modules.
        try:
            _ = xl.VBE
        except Exception:
            raise RuntimeError(
                "Excel Trust Center -> Macro Settings -> 'Trust access to the "
                "VBA project object model' must be enabled.")

        target = xl.Workbooks.Add()
        src = xl.Workbooks.Open(source, ReadOnly=True, UpdateLinks=0)

        # 1) Copy each kept sheet (full fidelity) in canonical order.
        src_names = [s.Name for s in src.Worksheets]
        keep_present = [n for n in KEEP if n in src_names]
        for n in keep_present:
            src.Worksheets(n).Copy(After=target.Worksheets(target.Worksheets.Count))
        # 2) Delete the workbook's original default sheet(s).
        for s in list(target.Worksheets):
            if s.Name not in keep_present:
                s.Delete()
        src.Close(SaveChanges=False)

        # 3) Break any cross-workbook links introduced by the copies.
        try:
            links = target.LinkSources(1)   # xlExcelLinks
            if links:
                for lk in links:
                    target.BreakLink(lk, 1)
        except Exception:
            pass

        # 4) Recreate workbook-scoped named ranges on the upload sheet.
        up = target.Worksheets(UPLOAD_SHEET)
        for nm, ref in NAMED.items():
            try:
                target.Names(nm).Delete()
            except Exception:
                pass
            target.Names.Add(nm, "='%s'!%s" % (UPLOAD_SHEET, ref))

        # 5) Stamp each canonical sheet blank from its snapshot (pristine template).
        for i, name in enumerate(CANON):
            tpl = "_Template_%02d" % i
            if name in [s.Name for s in target.Worksheets] and \
                    tpl in [s.Name for s in target.Worksheets]:
                t = target.Worksheets(tpl)
                d = target.Worksheets(name)
                d.Cells.Clear()
                t.UsedRange.Copy(d.Range("A1"))

        # 6) Remove the on-sheet upload buttons (the ribbon hosts them now).
        for shp in list(up.Shapes):
            try:
                if "Btn_" in (shp.OnAction or ""):
                    shp.Delete()
            except Exception:
                pass

        # 7) Visibility default + import the canonical VBA.
        _set_default_visibility(target)
        _import_vba(target)

        # 8) Save as .xlsm.
        if os.path.exists(out):
            os.remove(out)
        target.SaveAs(out, FileFormat=XL_OPENXML_MACRO)
        target.Close(SaveChanges=False)
    finally:
        xl.Quit()

    # 9) Inject the ribbon at the zip level using the source as scaffold.
    inject_customui(out, REPO, source)
    print("Built clean workbook ->", out)


def _set_default_visibility(wb):
    visible = {"DataViewer Upload", "Test SOP's", "Lifetime Test"}
    for s in wb.Worksheets:
        n = s.Name
        if n.startswith("_Template_") or n.startswith("_Macro"):
            s.Visible = XL_VERYHIDDEN
        elif n in visible:
            s.Visible = XL_VISIBLE
        else:
            s.Visible = XL_HIDDEN


def _import_vba(wb):
    proj = wb.VBProject
    # Remove any standard modules with our names, then import fresh.
    for comp in list(proj.VBComponents):
        if comp.Type == 1 and comp.Name in ("DataViewerUpload", "TestingTools",
                                             "SampleNav", "Module1"):
            proj.VBComponents.Remove(comp)
    for basfile in ("DataViewerUpload.bas", "TestingTools.bas", "SampleNav.bas"):
        proj.VBComponents.Import(os.path.join(REPO, basfile))
    # ThisWorkbook is a document module: replace its code (don't add a component).
    twb = proj.VBComponents("ThisWorkbook").CodeModule
    twb.DeleteLines(1, twb.CountOfLines)
    body = open(os.path.join(REPO, "ThisWorkbook.cls.txt"), encoding="utf-8").read()
    twb.AddFromString(body)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--source", required=True, help="messy source .xlsm")
    ap.add_argument("--out", required=True, help="clean output .xlsm to write")
    args = ap.parse_args()
    problems = check_preconditions(args.source)
    if problems:
        print("Preconditions failed:")
        for p in problems:
            print("  -", p)
        return 1
    build(args.source, args.out)
    print("Done. Now run: python excel-sidecar/verify_sidecar.py --file \"%s\"" % args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Write the headless helper test**

```python
#!/usr/bin/env python3
"""Headless tests for build_clean_template helpers (no Excel needed).

    python excel-sidecar/test_build_helpers.py --source "C:\\...\\...v1.xlsm"
Exit 0 on PASS.
"""
import argparse
import os
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_clean_template as B   # noqa: E402


def make_stripped(src, dst):
    """Copy `src` minus its customUI/* and xl/webextensions/* PARTS -> a stand-in
    for a freshly-built workbook. inject_customui is idempotent, so any leftover
    rels/content-type entries are harmless for this test."""
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            n = it.filename
            if n.startswith("customUI/") or n.startswith("xl/webextensions/"):
                continue
            zout.writestr(it, zin.read(n))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    args = ap.parse_args()
    fails = []

    # check_preconditions detects a missing source.
    if not B.check_preconditions("C:/does/not/exist.xlsm"):
        fails.append("check_preconditions should flag a missing source")

    # inject_customui produces a wired ribbon and strips no-add-in.
    tmpdir = os.environ.get("TEMP", ".")
    stripped = os.path.join(tmpdir, "_dv_stripped.xlsm")
    make_stripped(args.source, stripped)
    B.inject_customui(stripped, HERE, args.source)
    with zipfile.ZipFile(stripped) as z:
        names = z.namelist()
        out_ui = z.read("customUI/customUI14.xml").decode("utf-8")
        root_rels = z.read("_rels/.rels").decode("utf-8")
        ctypes = z.read("[Content_Types].xml").decode("utf-8")
    repo_ui = open(os.path.join(HERE, "customUI14.xml"), encoding="utf-8").read()
    if out_ui.strip() != repo_ui.strip():
        fails.append("injected customUI14.xml != repo copy")
    if "customUI/_rels/customUI14.xml.rels" not in names:
        fails.append("customUI rels missing")
    if "customUI/customUI14.xml" not in root_rels:
        fails.append("root _rels/.rels missing customUI relationship")
    if "customUI/customUI14.xml" not in ctypes:
        fails.append("[Content_Types].xml missing customUI override")
    if any(n.startswith("xl/webextensions/") for n in names):
        fails.append("web add-in parts survived injection")
    try:
        os.remove(stripped)
    except OSError:
        pass

    for f in fails:
        print("FAIL", f)
    print("RESULT:", "ALL PASS" if not fails else "FAILURES")
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 3: Run the helper test against the live source**

Run: `cd <worktree> && PY excel-sidecar/test_build_helpers.py --source "C:/Users/S1134987/Documents/Templates/Automated Testing Template v1.xlsm"`
Expected: `RESULT: ALL PASS` (exit 0). This validates the zip-level ribbon injection + precondition logic headlessly. (The Excel-COM `build()` is exercised on the work machine in Task 9.)

- [ ] **Step 4: Commit**

```bash
git add excel-sidecar/build_clean_template.py excel-sidecar/test_build_helpers.py
git commit -m "feat(sidecar): Excel-driven clean rebuild script + headless helper tests"
```

---

## Task 8: Docs — README + RUNBOOK

**Files:**
- Modify: `excel-sidecar/README.md`
- Modify: `excel-sidecar/RUNBOOK-migrate-existing-template.md`

- [ ] **Step 1: Update `README.md`**

Add a section documenting the new behavior + build flow. Insert after the "How the workbook works" section:

```markdown
## Clean rebuild (the canonical way to update the workbook)

The deployed `.xlsm` is rebuilt from these sources with
`build_clean_template.py` — a brand-new workbook (fresh package: no inherited
corruption, no web add-in) into which the source's sheets are copied, the VBA
imported, and the ribbon injected. Run on a machine with Excel + pywin32:

    python excel-sidecar/build_clean_template.py \
        --source "C:\path\Automated Testing Template v1.xlsm" \
        --out    "C:\path\Automated Testing Template v1 (clean).xlsm"
    python excel-sidecar/verify_sidecar.py --file "C:\path\...(clean).xlsm"   # expect all MATCH / OK

Requires Excel Trust Center -> "Trust access to the VBA project object model".
Headless checks that gate the sources (run anywhere):

    python excel-sidecar/check_sources.py        # invariant checks -> ALL PASS
    python excel-sidecar/test_build_helpers.py --source "C:\path\...v1.xlsm"

## Upload All keeps a Review copy (non-destructive reset)

On Upload All, after the data is distributed, each uploaded sheet is **renamed
into a `<name> - Review` copy** (data/formatting/formulas intact) and a fresh
blank sheet from the internal `_Template_NN` snapshot takes its place. Review
sheets persist (moved to the end of the workbook) until you click **Delete All
Review Sheets** on the ribbon. They never enter the distributed `.xlsx` copies.

## Ribbon

Every action now lives on the **TPM Testing** ribbon tab:
Sample Blocks · Sample Navigation · **DataViewer Upload** (Upload All, Dry-Run
Checklist, Pick Synology Folder, Pick Local Folder, Delete All Review Sheets).
```

Also update the "Install / update" note that says the deployed file runs an
older `DataViewerUpload` — replace it with: "Rebuild with
`build_clean_template.py` (see above); the result matches this folder by
construction."

- [ ] **Step 2: Update `RUNBOOK-migrate-existing-template.md`** with the operator acceptance checklist (the §9 list from the spec) and the manual fallback (new workbook → Move/Copy sheets → import VBA in the VBE → apply ribbon via the Custom UI Editor → save). Use the exact checklist from Task 9 Step 2 below.

- [ ] **Step 3: Commit**

```bash
git add excel-sidecar/README.md excel-sidecar/RUNBOOK-migrate-existing-template.md
git commit -m "docs(sidecar): rebuild flow, Review feature, ribbon, operator acceptance"
```

---

## Task 9: Work-machine build + acceptance (HANDOFF — run by the user on the work machine)

> This task cannot run in the dev environment (no Excel). It is the operator's verification round-trip. Everything above is authored + headless-verified; this proves it in Excel.

**Files:** none (produces the rebuilt `.xlsm`, not committed).

- [ ] **Step 1: Build + drift-verify**

On the work machine (Excel installed, `pip install pywin32`, VBA-object-model trust on):

```bat
PY excel-sidecar\check_sources.py
PY excel-sidecar\test_build_helpers.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm"
PY excel-sidecar\build_clean_template.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm" --out "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1 (clean).xlsm"
PY excel-sidecar\verify_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1 (clean).xlsm"
```
Expected: `check_sources` + `test_build_helpers` PASS; `build` reports "Built clean workbook"; `verify_sidecar` reports all modules MATCH, `customUI14.xml == repo`, `no web-extension/add-in parts`, `RESULT: all modules match`.

- [ ] **Step 2: Operator acceptance checklist (Excel, on the rebuilt file)**

1. Open the rebuilt file, change nothing, close → **no save prompt**.
2. Toggle a test TRUE/FALSE → matching sheet shows/hides; Add/Remove Sample, Reset Formulas, First/Prev/Next/Last + Ctrl+Shift+. / , all work.
3. Type `20` in a puffs seed cell → the column fills; type into a puffs cell that errors (e.g. on a protected area) → after the error, **events still alive** (toggles still respond).
4. Dry-Run Checklist passes on a populated sheet.
5. Upload All on ≥1 populated test → `.xlsx` lands in Synology + Local + opens in DataViewer; each uploaded sheet now has a `… - Review` copy at the end; the canonical sheet is blank/fresh; selection reset to Lifetime-only.
6. Upload a second test without deleting reviews → a second Review sheet appears (a counter if it's the same test); open a distributed `.xlsx` → **no** Review sheets in it.
7. Delete All Review Sheets → all `… - Review` gone in one click; canonical/utility sheets untouched.
8. (Optional) Re-run the forensic dump → no `xl/webextensions/*`, no `Workbook_SheetBeforeDoubleClick`.

- [ ] **Step 3: Ship** — per the branch-to-main workflow: only after the user approves the rebuilt file do they transfer it / merge the branch. **Never auto-drop on Synology.**

---

## Notes for the executor

- Run every Python step with the MIP-allowlisted interpreter
  `C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe`.
- `excel-sidecar/*` sources in this worktree are plain text (outside `Documents`) — edit directly.
- Do not modify `SampleNav.bas`.
- If an `imageMso` renders blank on the work machine, swap it for a verified Office imageMso ID — non-blocking (the label still shows).
- Keep `excel-sidecar/` as the single source of truth: any later macro change happens here, then rebuild.
```

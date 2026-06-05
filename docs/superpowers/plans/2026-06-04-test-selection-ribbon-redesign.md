# Test Selection + Ribbon Redesign — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> (recommended) or `superpowers:executing-plans` to implement this plan task-by-task.
> Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce the operator's upload sheet to a single, well-formatted checkbox
table ("Test Selection") and move every other affordance — paths, file name,
status/log, instructions — onto the TPM Testing ribbon or into popups.

**Architecture:** The deployed `.xlsm` is rebuilt by `build_clean_template.py` from
the canonical sources in `excel-sidecar/`. VBA + ribbon are authored and
headless-verified in the dev env (`check_sources.py`, `test_build_helpers.py`); the
actual build (pywin32 + Excel) + behavioral acceptance run on the work machine. The
native in-cell checkboxes **cannot** be created via COM, so the build **preserves**
the source's existing `A3:A15` checkbox cells (reorders by value-write, never clears
them). See spec: `docs/superpowers/specs/2026-06-04-test-selection-ribbon-redesign.md`.

**Tech Stack:** Excel VBA, Office customUI14 ribbon XML, Python 3.13 (allowlisted MIP
reader) + pywin32, openpyxl/oletools for headless checks.

**Conventions:**
- Worktree root: `C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\excel-sidecar-rebuild`
  (branch `feature/excel-sidecar-rebuild`). Files here are plain text (non-MIP) — edit directly.
- Python: `C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe`.
- Live workbook (build `--source`): `C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm`.
- Run all `python excel-sidecar/...` commands from the worktree root.

---

## File Structure

| File | Change |
|------|--------|
| `excel-sidecar/DataViewerUpload.bas` | Sheet rename, Test SOP's toggle, `_Settings` very-hide, file-name prompt, MsgBox results, dynamic-ribbon callbacks, Pick-DataViewer + Instructions handlers |
| `excel-sidecar/customUI14.xml` | Ribbon redesign (Help before Upload, 2-col layout, Delete column, Pick DataViewer File, Active Folders editBoxes, `onLoad`) |
| `excel-sidecar/check_sources.py` | New invariants for the above |
| `excel-sidecar/build_clean_template.py` | Rename, `_Settings`, name relocation, in-place Test Selection re-lay (preserve checkboxes), visibility |
| `excel-sidecar/test_build_helpers.py` | Assert new `NAMED`/`KEEP`/`UPLOAD_SHEET` |
| `excel-sidecar/verify_sidecar.py` | New workbook-structure check |
| `excel-sidecar/README.md`, `RUNBOOK-migrate-existing-template.md` | Document new design + Option-B fallback + acceptance |
| `ThisWorkbook.cls.txt`, `TestingTools.bas`, `SampleNav.bas` | **No change** |

---

# PHASE A — authored + headless-verified in the dev env

## Task 1: `DataViewerUpload.bas` — behavior changes

**Files:** Modify `excel-sidecar/DataViewerUpload.bas`

- [ ] **Step 1 — Rename the sheet constant.** Replace:

```vba
Private Const UPLOAD_SHEET_NAME As String = "DataViewer Upload"
```

with:

```vba
Private Const UPLOAD_SHEET_NAME As String = "Test Selection"
```

(Comments that mention the *ribbon group* "DataViewer Upload" stay — only the sheet
is renamed. Optionally update the workflow header comment's "DataViewer Upload" sheet
references to "Test Selection" for accuracy.)

- [ ] **Step 2 — Test SOP's visibility toggle.** In `ApplySheetVisibility`, replace:

```vba
    ' Always-visible utility sheets.
    EnsureVisible UPLOAD_SHEET_NAME, xlSheetVisible
    EnsureVisible SOPS_SHEET_NAME, xlSheetVisible
```

with:

```vba
    ' Always-visible: the Test Selection sheet itself.
    EnsureVisible UPLOAD_SHEET_NAME, xlSheetVisible
    ' Test SOP's: visibility now follows its checkbox (default visible).
    Dim sopVisible As Boolean
    sopVisible = True
    If selection.Exists(NormalizeSheetName(SOPS_SHEET_NAME)) Then
        sopVisible = selection(NormalizeSheetName(SOPS_SHEET_NAME))
    End If
    If sopVisible Then
        EnsureVisible SOPS_SHEET_NAME, xlSheetVisible
    Else
        EnsureVisible SOPS_SHEET_NAME, xlSheetHidden
    End If
```

- [ ] **Step 3 — Very-hide `_Settings`.** In `ApplySheetVisibility`, after
  `EnsureVeryHidden "_Template_Master"`, add:

```vba
    EnsureVeryHidden "_Settings"
```

- [ ] **Step 4 — Default Test SOP's visible on reset.** In `ResetSelectionToDefault`,
  replace the inner `If StrComp(... DEFAULT_SELECTED_SHEET ...)` block with:

```vba
        If IsDefaultVisibleSheet(sheetName) Then
            selRange.Cells(row, 1).value = True
        Else
            selRange.Cells(row, 1).value = False
        End If
```

  and add this helper just below `ResetSelectionToDefault`:

```vba
Private Function IsDefaultVisibleSheet(sheetName As String) As Boolean
    ' Sheets that should be TRUE after a post-upload reset: Lifetime + Test SOP's.
    Dim n As String
    n = NormalizeSheetName(sheetName)
    IsDefaultVisibleSheet = _
        (StrComp(n, NormalizeSheetName(DEFAULT_SELECTED_SHEET), vbTextCompare) = 0) Or _
        (StrComp(n, NormalizeSheetName(SOPS_SHEET_NAME), vbTextCompare) = 0)
End Function
```

- [ ] **Step 5 — Drop the file-name check from the checklist.** In `RunChecklist`,
  delete this line (file name is an upload-time prompt now):

```vba
    If Len(Trim$(GetNamed("DV_FileName"))) = 0 Then failures.Add "DV_FileName is empty"
```

- [ ] **Step 6 — Prompt for the file name at upload.** Add this helper (e.g. above
  `Btn_DryRunChecklist`):

```vba
Private Function PromptForFileName() As String
    Dim cur As String
    cur = Trim$(GetNamed("DV_FileName"))
    Dim r As Variant
    r = Application.InputBox( _
        Prompt:="Enter a descriptive file name (Product + Test + Date):", _
        Title:="Upload file name", Default:=cur, Type:=2)
    If VarType(r) = vbBoolean Then
        PromptForFileName = ""        ' Cancel
    Else
        PromptForFileName = Trim$(CStr(r))
    End If
End Function
```

  In `Btn_UploadAll`, immediately after `StampLog "Upload started"`, insert:

```vba
    ' Ask for the descriptive file name up front (pre-filled with the last one).
    Dim fName As String
    fName = PromptForFileName()
    If Len(fName) = 0 Then
        SetNamed "DV_Status", "Cancelled (no file name)"
        StampLog "Upload cancelled - no file name entered"
        Exit Sub
    End If
    SetNamed "DV_FileName", fName
```

- [ ] **Step 7 — Results via MsgBox.** Add the helper (near `WriteFailures`):

```vba
Private Sub ShowFailures(title As String, failures As Collection)
    Dim msg As String, n As Long, shown As Long
    shown = 0
    For n = 1 To failures.Count
        If shown >= 20 Then
            msg = msg & vbLf & "   ...and " & (failures.Count - shown) & _
                  " more (see the log on the hidden _Settings sheet)."
            Exit For
        End If
        msg = msg & vbLf & "  - " & CStr(failures(n))
        shown = shown + 1
    Next
    MsgBox "Found " & failures.Count & " issue(s):" & vbLf & msg, vbExclamation, title
End Sub
```

  In `Btn_DryRunChecklist`, replace the `If failures.Count = 0 ... Else ... End If`
  tail with:

```vba
    If failures.Count = 0 Then
        SetNamed "DV_Status", "OK"
        StampLog "Checklist passed"
        MsgBox "Checklist passed - ready to upload.", vbInformation, "Dry-Run Checklist"
    Else
        WriteFailures failures
        SetNamed "DV_Status", "Failed: " & failures.Count & " checklist issue(s)"
        ShowFailures "Dry-Run Checklist", failures
    End If
```

  In `Btn_UploadAll`, add MsgBoxes at these existing branches:
  - after `WriteFailures failures` (checklist-failed branch), before `Exit Sub`:
    `ShowFailures "Upload All", failures`
  - the Synology-not-accessible branch: after its `SetNamed`, add
    `MsgBox "Synology path not accessible:" & vbLf & synPath, vbExclamation, "Upload All"`
  - the Local-not-accessible branch: add
    `MsgBox "Local path not accessible:" & vbLf & locPath, vbExclamation, "Upload All"`
  - the `keep.Count <= 1` branch: add
    `MsgBox "No selected sheets contain data.", vbExclamation, "Upload All"`
  - the DataViewer-not-found branch: add
    `MsgBox "DataViewer.exe not found at:" & vbLf & dvExe, vbExclamation, "Upload All"`
  - just before the success `Exit Sub` (after `StampLog "Done"`):
    ```vba
    MsgBox "Upload complete." & vbLf & vbLf & baseName & ".xlsx was sent to the " & _
           "Synology and Local folders and opened in DataViewer." & vbLf & _
           "Each uploaded sheet was reset (a '- Review' copy was kept).", _
           vbInformation, "Upload All"
    ```
  - in the `Failed:` handler, after its `SetNamed`, add
    `MsgBox "Upload failed:" & vbLf & Err.Description, vbCritical, "Upload All"`

- [ ] **Step 8 — Commit.**

```bash
git add excel-sidecar/DataViewerUpload.bas
git commit -m "feat(sidecar): Test Selection rename, SOP toggle, file-name prompt, MsgBox results"
```

---

## Task 2: `DataViewerUpload.bas` — dynamic-ribbon + new callbacks

**Files:** Modify `excel-sidecar/DataViewerUpload.bas`

- [ ] **Step 1 — Module-level ribbon handle.** Add near the top of the module (after
  the `Private Const` block):

```vba
' Captured at ribbon load so the path rows can be refreshed after a Pick.
Public gRibbon As IRibbonUI
```

- [ ] **Step 2 — Add the callback block** (append to the "Ribbon callbacks" section at
  the end of the module, alongside the existing `Ribbon_*` wrappers):

```vba
Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub RefreshPathLabels()
    On Error Resume Next
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "ebSynPath"
        gRibbon.InvalidateControl "ebLocPath"
        gRibbon.InvalidateControl "ebExePath"
    End If
    On Error GoTo 0
End Sub

' --- Active Folders read-only path rows (editBox getText/getSupertip/onChange) ---
Public Sub GetSynPathText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetNamed("DV_SynologyPath")
End Sub
Public Sub GetLocPathText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetNamed("DV_LocalPath")
End Sub
Public Sub GetExePathText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetNamed("DV_DataViewerExe")
End Sub
Public Sub GetSynPathTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Synology folder:" & vbLf & GetNamed("DV_SynologyPath")
End Sub
Public Sub GetLocPathTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Local folder:" & vbLf & GetNamed("DV_LocalPath")
End Sub
Public Sub GetExePathTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "DataViewer.exe:" & vbLf & GetNamed("DV_DataViewerExe")
End Sub
' Read-only: discard any keystroke by re-reading the stored value.
Public Sub PathReadOnly(control As IRibbonControl, text As String)
    RefreshPathLabels
End Sub

' --- Pick DataViewer File ---
Public Sub Btn_PickDataViewerExe()
    Dim chosen As String
    chosen = PickFileForNamed("Choose DataViewer.exe", "DataViewer (*.exe),*.exe")
    If Len(chosen) > 0 Then
        SetNamed "DV_DataViewerExe", chosen
        RefreshPathLabels
    End If
End Sub

Private Function PickFileForNamed(title As String, filterStr As String) As String
    Dim r As Variant
    r = Application.GetOpenFilename(filterStr, , title)
    If VarType(r) = vbBoolean Then
        PickFileForNamed = ""
    Else
        PickFileForNamed = CStr(r)
    End If
End Function

' --- Instructions ---
Public Sub Btn_ShowInstructions()
    MsgBox InstructionsText(), vbInformation, "DataViewer Upload - Instructions"
End Sub

Private Function InstructionsText() As String
    Dim s As String
    s = "How to upload test data to DataViewer" & vbLf & vbLf
    s = s & "1. Test Selection sheet: tick the tests you're running. Each ticked" & vbLf
    s = s & "   test's sheet appears - fill it in. (Untick to hide a sheet again.)" & vbLf & vbLf
    s = s & "2. Set your destinations once, from the TPM Testing ribbon:" & vbLf
    s = s & "   Pick Synology Folder  -  Pick Local Folder  -  Pick DataViewer File" & vbLf
    s = s & "   The current paths show in the 'Active Folders' box on the ribbon." & vbLf & vbLf
    s = s & "3. (Optional) Dry-Run Checklist validates your data without uploading." & vbLf & vbLf
    s = s & "4. Upload All: enter a descriptive file name when prompted" & vbLf
    s = s & "   (Product + Test + Date). The data is copied to the Synology and Local" & vbLf
    s = s & "   folders, opened in DataViewer, and each uploaded sheet is reset -" & vbLf
    s = s & "   a '<name> - Review' copy is kept so you can see what was sent." & vbLf & vbLf
    s = s & "5. A summary pops up when it finishes. Use 'Delete All Review Sheets'" & vbLf
    s = s & "   to clear the review copies once you're done with them." & vbLf & vbLf
    s = s & "Tip: 'Test SOP's' is always included in every upload, whether or not" & vbLf
    s = s & "its box is ticked."
    InstructionsText = s
End Function

Public Sub Ribbon_PickDataViewer(control As IRibbonControl)
    Btn_PickDataViewerExe
End Sub
Public Sub Ribbon_Instructions(control As IRibbonControl)
    Btn_ShowInstructions
End Sub
```

- [ ] **Step 3 — Refresh the ribbon after the existing folder picks.** Add
  `RefreshPathLabels` to both existing handlers:

```vba
Public Sub Btn_PickSynologyFolder()
    PickFolderInto "DV_SynologyPath", "Choose the Synology data folder"
    RefreshPathLabels
End Sub

Public Sub Btn_PickLocalFolder()
    PickFolderInto "DV_LocalPath", "Choose the local data folder"
    RefreshPathLabels
End Sub
```

- [ ] **Step 4 — Commit.**

```bash
git add excel-sidecar/DataViewerUpload.bas
git commit -m "feat(sidecar): dynamic ribbon path rows, Pick DataViewer File, Instructions popup"
```

---

## Task 3: `customUI14.xml` — ribbon redesign

**Files:** Replace `excel-sidecar/customUI14.xml` entirely. **Preserve the existing
`grpSampleBlocks` and `grpSampleNav` group contents verbatim** (copy them from the
current file into the new structure below).

- [ ] **Step 1 — Write the new ribbon.** Full file:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_OnLoad">
  <ribbon>
    <tabs>
      <tab id="tabTPMTesting" label="TPM Testing" insertAfterMso="TabHome">

        <group id="grpSampleBlocks" label="Sample Blocks">
          <!-- UNCHANGED: copy the three buttons (Add/Remove/Reset) from the current file -->
        </group>

        <group id="grpSampleNav" label="Sample Navigation">
          <!-- UNCHANGED: copy the four buttons (First/Prev/Next/Last) from the current file -->
        </group>

        <group id="grpHelp" label="Help">
          <button id="btnInstructions" label="Instructions" size="large" imageMso="Info"
                  onAction="Ribbon_Instructions"
                  screentip="How to use the upload workflow"
                  supertip="Step-by-step instructions for selecting tests and uploading to DataViewer."/>
        </group>

        <group id="grpDvUpload" label="DataViewer Upload">
          <box id="boxDvActions" boxStyle="vertical">
            <button id="btnUploadAll" label="Upload All" size="normal"
                    onAction="Ribbon_UploadAll"
                    screentip="Validate, upload, and reset"
                    supertip="Runs the checklist, copies the data to Synology + Local + DataViewer, then resets each uploaded sheet (a '- Review' copy is kept)."/>
            <button id="btnDryRun" label="Dry-Run Checklist" size="normal"
                    onAction="Ribbon_DryRun"
                    screentip="Validate without uploading"
                    supertip="Runs the upload checklist and reports issues without copying or resetting anything."/>
          </box>
          <box id="boxDvPickers" boxStyle="vertical">
            <button id="btnPickSyn" label="Pick Synology Folder" size="normal" imageMso="FolderOpen"
                    onAction="Ribbon_PickSynology" screentip="Choose the Synology data folder"/>
            <button id="btnPickLoc" label="Pick Local Folder" size="normal" imageMso="FolderOpen"
                    onAction="Ribbon_PickLocal" screentip="Choose the local data folder"/>
            <button id="btnPickExe" label="Pick DataViewer File" size="normal" imageMso="FileOpen"
                    onAction="Ribbon_PickDataViewer" screentip="Choose the DataViewer.exe to launch"/>
          </box>
          <button id="btnDelReviews" label="Delete All Review Sheets" size="large" imageMso="Delete"
                  onAction="Ribbon_DeleteReviewSheets"
                  screentip="Delete every '- Review' sheet"
                  supertip="One click removes all '- Review' sheets. Canonical, template, and utility sheets are never touched."/>
        </group>

        <group id="grpDvPaths" label="Active Folders">
          <editBox id="ebSynPath" label="Synology:" sizeString="WWWWWWWWWWWWWWWWWWWWWWWWW"
                   getText="GetSynPathText" onChange="PathReadOnly" getSupertip="GetSynPathTip"
                   screentip="Current Synology folder (read-only)"/>
          <editBox id="ebLocPath" label="Local:" sizeString="WWWWWWWWWWWWWWWWWWWWWWWWW"
                   getText="GetLocPathText" onChange="PathReadOnly" getSupertip="GetLocPathTip"
                   screentip="Current local folder (read-only)"/>
          <editBox id="ebExePath" label="DataViewer:" sizeString="WWWWWWWWWWWWWWWWWWWWWWWWW"
                   getText="GetExePathText" onChange="PathReadOnly" getSupertip="GetExePathTip"
                   screentip="Current DataViewer.exe (read-only)"/>
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>
```

- [ ] **Step 2 — Verify well-formedness.**

```bash
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" -c "import xml.dom.minidom,sys; xml.dom.minidom.parse(r'excel-sidecar/customUI14.xml'); print('XML OK')"
```
Expected: `XML OK`.

- [ ] **Step 3 — Commit.**

```bash
git add excel-sidecar/customUI14.xml
git commit -m "feat(sidecar): ribbon redesign - Help, 2-col upload/pick, Active Folders path rows"
```

---

## Task 4: `check_sources.py` — extend invariants

**Files:** Modify `excel-sidecar/check_sources.py`. Read the file first; it already has
a check list + an onAction↔handler cross-check. Add the following invariants in the
same style (each prints `[PASS]`/`[FAIL]` and feeds the final ALL PASS/`FAIL` tally).

- [ ] **Step 1 — Ribbon callback cross-check (extend existing).** Extend the ribbon
  callback extraction to also collect `getText`, `getSupertip`, `onChange`, and the
  root `onLoad` attribute values (not just `onAction`). For every collected callback
  name, assert a matching `Public Sub <name>` exists in `DataViewerUpload.bas`. New
  callbacks that must resolve: `Ribbon_OnLoad`, `GetSynPathText`, `GetLocPathText`,
  `GetExePathText`, `GetSynPathTip`, `GetLocPathTip`, `GetExePathTip`, `PathReadOnly`,
  `Ribbon_PickDataViewer`, `Ribbon_Instructions`.

- [ ] **Step 2 — Ribbon structure checks.** Assert in `customUI14.xml`:
  - `onLoad="Ribbon_OnLoad"` on `<customUI>`.
  - `grpHelp` appears **before** `grpDvUpload` (use string index comparison).
  - `btnUploadAll` and `btnDryRun` have **no** `image`/`imageMso` attribute.
  - `btnDelReviews` has `size="large"`.
  - editBox ids `ebSynPath`, `ebLocPath`, `ebExePath` all present.
  - Each `<box ...>` contains at most 3 child `<button`s (≤3-rows rule).

- [ ] **Step 3 — VBA invariants.** Assert in `DataViewerUpload.bas`:
  - `UPLOAD_SHEET_NAME As String = "Test Selection"` present.
  - `PromptForFileName` defined; `Btn_UploadAll` calls `PromptForFileName`.
  - `RunChecklist` does **not** contain `DV_FileName is empty`.
  - `MsgBox` appears in both `Btn_DryRunChecklist` and `Btn_UploadAll`; `ShowFailures`
    defined.
  - `IsDefaultVisibleSheet` defined and references `SOPS_SHEET_NAME`.
  - `BuildKeepList` still contains `keep.Add SOPS_SHEET_NAME, True`.
  - `EnsureVeryHidden "_Settings"` present.
  - `Public gRibbon As IRibbonUI` and `RefreshPathLabels` defined; `RefreshPathLabels`
    referenced inside `Btn_PickSynologyFolder` and `Btn_PickLocalFolder`.
  - `InstructionsText` defined and contains the phrase `always included in every upload`.

- [ ] **Step 4 — Run green.**

```bash
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/check_sources.py
```
Expected: `RESULT: ALL PASS` (count increased by the new checks).

- [ ] **Step 5 — Commit.**

```bash
git add excel-sidecar/check_sources.py
git commit -m "test(sidecar): check_sources invariants for the Test Selection redesign"
```

---

## Task 5: `build_clean_template.py` — rename, `_Settings`, in-place Test Selection re-lay

**Files:** Modify `excel-sidecar/build_clean_template.py`. The native checkboxes
cannot be created via COM, so the build **preserves** the source's `A3:A15` cells.

- [ ] **Step 1 — Module constants.** Replace the `CANON`/`SNAPSHOTS`/`KEEP`/`NAMED`/
  `UPLOAD_SHEET` block near the top with:

```python
UPLOAD_SHEET = "Test Selection"
OLD_UPLOAD_SHEET = "DataViewer Upload"
SETTINGS_SHEET = "_Settings"
CANON = [
    "Lifetime Test", "User Test Simulation", "Long Puff Lifetime Test",
    "Rapid Puff Lifetime Test", "Intense Test", "Big Headspace Serial Test",
    "Negative Pressure Test", "Temperature Cycling Test #2",
    "Viscosity Compatibility", "Various Oil Compatibility",
    "Custom Test Template", "Temperature Cycling Test #1",
]
SNAPSHOTS = ["_Template_%02d" % i for i in range(12)]
KEEP = ["Test SOP's", UPLOAD_SHEET, SETTINGS_SHEET] + CANON + ["_Template_Master"] + SNAPSHOTS

# name -> full (sheet-qualified) RefersTo
NAMED = {
    "DV_TestSelection": "'Test Selection'!$A$3:$B$15",
    "DV_FileName":      "'_Settings'!$B$1",
    "DV_SynologyPath":  "'_Settings'!$B$2",
    "DV_LocalPath":     "'_Settings'!$B$3",
    "DV_DataViewerExe": "'_Settings'!$B$4",
    "DV_Status":        "'_Settings'!$B$5",
    "DV_Log":           "'_Settings'!$B$6",
}

# Display order on the Test Selection sheet: (sheet name, default checked)
SELECTION_ROWS = [
    ("Custom Test Template", False), ("Lifetime Test", True),
    ("Long Puff Lifetime Test", False), ("Rapid Puff Lifetime Test", False),
    ("Intense Test", False), ("User Test Simulation", False),
    ("Big Headspace Serial Test", False), ("Viscosity Compatibility", False),
    ("Various Oil Compatibility", False), ("Temperature Cycling Test #1", False),
    ("Temperature Cycling Test #2", False), ("Negative Pressure Test", False),
    ("Test SOP's", True),
]
SETTINGS_LABELS = ["File name (last used)", "Synology folder", "Local folder",
                   "DataViewer.exe", "Status", "Log"]
```

- [ ] **Step 2 — New helpers.** Add these functions (near `_set_default_visibility`):

```python
def _carry_over_values(wb):
    """Read current path/file values before restructuring, so the operator does not
    re-pick folders after a rebuild."""
    carried = {}
    for nm in ("DV_FileName", "DV_SynologyPath", "DV_LocalPath", "DV_DataViewerExe"):
        try:
            carried[nm] = wb.Names(nm).RefersToRange.Value
        except Exception:
            carried[nm] = ""
    return carried


def _rename_upload_sheet(wb):
    """Rename 'DataViewer Upload' -> 'Test Selection' (idempotent)."""
    for cand in (OLD_UPLOAD_SHEET, UPLOAD_SHEET):
        try:
            wb.Worksheets(cand).Name = UPLOAD_SHEET
            return
        except Exception:
            continue


def _build_settings_sheet(wb, carried):
    try:
        st = wb.Worksheets(SETTINGS_SHEET)
    except Exception:
        st = wb.Worksheets.Add()
        st.Name = SETTINGS_SHEET
    st.Visible = XL_VISIBLE        # very-hidden later by _set_default_visibility
    st.Cells.Clear()
    for i, label in enumerate(SETTINGS_LABELS, start=1):
        st.Cells(i, 1).Value = label
    st.Cells(1, 2).Value = carried.get("DV_FileName", "") or ""
    st.Cells(2, 2).Value = carried.get("DV_SynologyPath", "") or ""
    st.Cells(3, 2).Value = carried.get("DV_LocalPath", "") or ""
    st.Cells(4, 2).Value = carried.get("DV_DataViewerExe", "") or ""
    st.Cells(5, 2).Value = ""     # Status
    st.Cells(6, 2).Value = ""     # Log
    st.Columns("A:B").AutoFit()


def _relay_test_selection(wb):
    """Re-lay the Test Selection sheet IN PLACE. Preserve the native checkboxes on
    A3:A15 (value writes only - never Clear that column)."""
    ws = wb.Worksheets(UPLOAD_SHEET)
    ws.Cells(1, 1).Value = "TEST SELECTION"
    ws.Cells(2, 1).Value = "Check the tests you're running."
    ws.Cells(1, 2).Value = ""
    ws.Cells(2, 2).Value = ""
    for i, (name, checked) in enumerate(SELECTION_ROWS):
        r = 3 + i
        ws.Cells(r, 1).Value = bool(checked)   # keeps the native checkbox
        ws.Cells(r, 2).Value = name
    # Clear only the clutter: old stray row + everything from column C on.
    ws.Range("A16:B200").ClearContents()
    ws.Range("C1:AZ200").Clear()
    # Cosmetic table formatting.
    ws.Range("A1:B1").Merge()
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2:B2").Merge()
    ws.Range("A2").Font.Italic = True
    ws.Columns(1).ColumnWidth = 9
    ws.Columns(2).ColumnWidth = 34
    ws.Range("A3:B15").Borders.LineStyle = 1   # xlContinuous
```

- [ ] **Step 3 — Wire into `build()`.** Immediately after `wb = xl.Workbooks.Open(...)`
  and its print, add:

```python
        carried = _carry_over_values(wb)
        _rename_upload_sheet(wb)
```

  After the canonical-stamp loop (step 3, the `for i, name in enumerate(CANON)` block)
  and before the name-recreate loop, add:

```python
        _build_settings_sheet(wb, carried)
        _relay_test_selection(wb)
```

  Replace the name-recreate loop body so it uses the full RefersTo (names now span two
  sheets):

```python
        for nm, ref in NAMED.items():
            try:
                wb.Names(nm).Delete()
            except Exception:
                pass
            wb.Names.Add(nm, "=" + ref)
```

  (Delete the old `up = wb.Worksheets(UPLOAD_SHEET)` line that prefixed the loop only
  if it's unused afterward; the `Btn_*` shape-removal step still needs
  `up = wb.Worksheets(UPLOAD_SHEET)` — keep that line with the shape-removal step.)

- [ ] **Step 4 — Update `_set_default_visibility`.** Replace its body with:

```python
def _set_default_visibility(wb):
    visible = {"Test Selection", "Test SOP's", "Lifetime Test"}
    for s in wb.Worksheets:
        n = s.Name
        if n.startswith("_Template_") or n.startswith("_Macro") or n == "_Settings":
            s.Visible = XL_VERYHIDDEN
        elif n in visible:
            s.Visible = XL_VISIBLE
        else:
            s.Visible = XL_HIDDEN
```

- [ ] **Step 5 — Byte-compile check** (no Excel needed):

```bash
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" -m py_compile excel-sidecar/build_clean_template.py && echo "compile OK"
```
Expected: `compile OK`.

- [ ] **Step 6 — Commit.**

```bash
git add excel-sidecar/build_clean_template.py
git commit -m "feat(sidecar): build renames sheet, builds _Settings, re-lays Test Selection (preserve checkboxes)"
```

---

## Task 6: `test_build_helpers.py` + `verify_sidecar.py` — extend checks

**Files:** Modify `excel-sidecar/test_build_helpers.py`, `excel-sidecar/verify_sidecar.py`

- [ ] **Step 1 — `test_build_helpers.py`.** Import the build module's constants and
  assert: `UPLOAD_SHEET == "Test Selection"`; `SETTINGS_SHEET in KEEP`;
  `"Test Selection" in KEEP`; `NAMED["DV_FileName"]` contains `_Settings`;
  `NAMED["DV_TestSelection"]` contains `Test Selection`; `len(SELECTION_ROWS) == 13`
  and `SELECTION_ROWS[0][0] == "Custom Test Template"` and
  `SELECTION_ROWS[-1] == ("Test SOP's", True)`.

- [ ] **Step 2 — `verify_sidecar.py` workbook-structure check.** Add this function and
  call it from `main()` (print its lines like `check_ribbon_and_addin`, OR-in its
  `any_diff`):

```python
def check_workbook_structure(xlsm_path):
    """openpyxl: Test Selection present, _Settings very hidden, no 'DataViewer Upload',
    named ranges resolve to the right sheets, and the 13 selection names are in order."""
    lines, any_diff = [], False
    EXPECTED = ["Custom Test Template", "Lifetime Test", "Long Puff Lifetime Test",
                "Rapid Puff Lifetime Test", "Intense Test", "User Test Simulation",
                "Big Headspace Serial Test", "Viscosity Compatibility",
                "Various Oil Compatibility", "Temperature Cycling Test #1",
                "Temperature Cycling Test #2", "Negative Pressure Test", "Test SOP's"]
    DEST = {"DV_TestSelection": "Test Selection", "DV_FileName": "_Settings",
            "DV_SynologyPath": "_Settings", "DV_LocalPath": "_Settings",
            "DV_DataViewerExe": "_Settings", "DV_Status": "_Settings", "DV_Log": "_Settings"}
    try:
        import openpyxl
    except Exception as e:
        return False, ["[i] openpyxl unavailable; structure check skipped (%r)" % e]
    try:
        wb = openpyxl.load_workbook(xlsm_path, read_only=False, keep_vba=True)
    except Exception as e:
        return True, ["[!] could not open workbook: %r" % e]

    def ok(cond, msg):
        nonlocal any_diff
        lines.append(("[OK]      " if cond else "[DIFFERS] ") + msg)
        if not cond:
            any_diff = True

    sn = wb.sheetnames
    ok("Test Selection" in sn, "'Test Selection' sheet present")
    ok("DataViewer Upload" not in sn, "old 'DataViewer Upload' sheet removed")
    ok("_Settings" in sn, "_Settings sheet present")
    try:
        ok(wb["_Settings"].sheet_state == "veryHidden", "_Settings is very hidden")
    except Exception:
        ok(False, "_Settings readable")

    # defined-name destinations (version-tolerant access)
    dn = wb.defined_names
    def dest_of(name):
        try:
            return dn[name].value
        except Exception:
            try:
                return dict(dn.items())[name].value
            except Exception:
                return None
    for nm, sheet in DEST.items():
        d = dest_of(nm)
        ok(d is not None and sheet in d, "%s -> %s (%r)" % (nm, sheet, d))

    if "Test Selection" in sn:
        ws = wb["Test Selection"]
        got = [ws.cell(3 + i, 2).value for i in range(13)]
        ok(got == EXPECTED, "Test Selection B3:B15 names in canonical order")
    return any_diff, lines
```

  In `main()`, after the `check_ribbon_and_addin` block:

```python
    st_diff, st_lines = check_workbook_structure(args.file)
    for l in st_lines:
        print(l)
    any_diff = any_diff or st_diff
```

- [ ] **Step 3 — Run the headless-capable checks.**

```bash
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/test_build_helpers.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm"
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" -m py_compile excel-sidecar/verify_sidecar.py && echo "verify compiles"
```
Expected: `test_build_helpers` ALL PASS; `verify compiles`. (Full `verify_sidecar`
runs in Phase B against the built file.)

- [ ] **Step 4 — Commit.**

```bash
git add excel-sidecar/test_build_helpers.py excel-sidecar/verify_sidecar.py
git commit -m "test(sidecar): build-helper constants + verify_sidecar workbook-structure check"
```

---

## Task 7: Docs — README + RUNBOOK + memory

**Files:** Modify `excel-sidecar/README.md`, `excel-sidecar/RUNBOOK-migrate-existing-template.md`

- [ ] **Step 1 — README.** Update "How the workbook works" + "Ribbon" sections: sheet
  is now **Test Selection** (checkbox table only, 13 rows in the documented order);
  paths/file-name/status moved off-sheet (`_Settings` very-hidden sheet, ribbon path
  rows, InputBox file name, MsgBox results); ribbon groups Help · DataViewer Upload ·
  Active Folders; Test SOP's is a visibility toggle that's still always uploaded.
  Note that native checkboxes are **preserved** by the build (not creatable via COM).

- [ ] **Step 2 — RUNBOOK.** Update the acceptance checklist with the spec's §"Operator
  acceptance criteria" items. Add an **Option-B fallback** subsection: *if the rebuilt
  Test Selection shows `TRUE`/`FALSE` text instead of checkboxes, a COM value-write
  stripped the native control — recover by (a) opening the rebuilt file, applying
  native checkboxes to `A3:A15` via Excel-on-web Office Script
  `range.control={type:"Checkbox"}` (or the ribbon Insert ▸ Checkbox), (b) running a
  one-time `SeedSelectionTemplate` to snapshot it as `_Template_Sel`, and (c) switching
  `_relay_test_selection` to stamp that snapshot.* (Describe; do not implement unless
  acceptance fails.)

- [ ] **Step 3 — Memory.** Update
  `C:\Users\S1134987\.claude\projects\C--Users-S1134987-Documents-Python-DataViewer-Dev-DataViewer-Enterprise\memory\excel-sidecar-rebuild.md`
  with a short "Test Selection redesign (2026-06-04)" status paragraph + the new spec/
  plan paths.

- [ ] **Step 4 — Commit.**

```bash
git add excel-sidecar/README.md excel-sidecar/RUNBOOK-migrate-existing-template.md
git commit -m "docs(sidecar): document Test Selection redesign + Option-B checkbox fallback"
```

- [ ] **Step 5 — Final headless gate (all of Phase A).**

```bash
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/check_sources.py
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/test_build_helpers.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm"
```
Expected: both `RESULT: ALL PASS`.

---

# PHASE B — work machine (build + behavioral acceptance)

## Task 8: Build + verify the rebuilt workbook

- [ ] **Step 1 — (Optional, 2-min de-risk) checkbox-preservation probe.** On a scratch
  copy, write a Boolean to a checkbox cell via COM, save, reopen, and eyeball that it's
  still a checkbox. If it became `TRUE`/`FALSE` text, use the Option-B fallback before
  building. (Skip if confident; Step 4 acceptance catches it regardless.)

- [ ] **Step 2 — Headless gates** (must be ALL PASS before building):

```bash
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/check_sources.py
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/test_build_helpers.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm"
```

- [ ] **Step 3 — Build** (backs up the source automatically):

```bash
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/build_clean_template.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm" --out "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1 (clean).xlsm"
```
Expected: `Built clean workbook -> ...`.

- [ ] **Step 4 — Verify:**

```bash
"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/verify_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1 (clean).xlsm"
```
Expected: all modules match, `customUI14.xml == repo`, no web add-in parts, and the new
structure block all `[OK]` → `RESULT: all modules match`.

## Task 9: Operator acceptance (in Excel) — then merge

- [ ] **Step 1 — Run the acceptance checklist** in
  `excel-sidecar/RUNBOOK-migrate-existing-template.md` against
  `…v1 (clean).xlsm`. The new-feature items:
  1. Test Selection shows **only** title + hint + 13-row table in the documented order.
  2. **A3:A15 render as checkboxes** (if text instead → Option-B fallback).
  3. Ticking/unticking a test shows/hides its sheet; **Test SOP's** toggles its sheet's
     visibility but is **still uploaded when unticked**.
  4. Ribbon: Help before DataViewer Upload; Upload All/Dry-Run text-only & stacked;
     three Pick buttons stacked; Delete its own column; **no group > 3 rows**.
  5. **Active Folders** shows the three paths (full on hover) and updates immediately on
     each Pick, incl. **Pick DataViewer File**.
  6. **Instructions** popup opens; no instructions remain on any sheet.
  7. **Upload All** prompts for a file name (pre-filled); success popup; `.xlsx` to
     Synology + Local + DataViewer; `- Review` copy kept; selection resets to
     Lifetime + SOP's.
  8. **Dry-Run** shows pass/fail popup, never prompts for a name.
  9. `_Settings` invisible; a distributed `.xlsx` has **no** `_Settings`/`Test Selection`/
     Review sheets.
  10. All prior acceptance items still pass (no save-prompt on close, nav/add/remove/
      reset, puff-picker crash-safety, Review accumulation, Delete-All).

- [ ] **Step 2 — On acceptance pass:** rename `…v1 (clean).xlsm` over the live template,
  then finish the branch via `superpowers:finishing-a-development-branch` (merge
  `feature/excel-sidecar-rebuild` → `main`). **Never** auto-drop on Synology.

---

## Self-Review (run before dispatching execution)

- **Spec coverage:** every spec section maps to a task — Mod 1 sheet → T5; Mod 1 ribbon
  → T2/T3; Mod 2 Instructions → T2/T3; `_Settings` → T1/T5; SOP toggle → T1; file-name
  → T1; MsgBox → T1; gates → T4/T6; docs → T7; build/accept → T8/T9. ✓
- **Type/name consistency:** callback names match between `customUI14.xml` (T3) and the
  `Public Sub`s (T2); `NAMED`/`SELECTION_ROWS` (T5) match `check_workbook_structure`'s
  `DEST`/`EXPECTED` (T6) and the spec tables. ✓
- **No placeholders:** every code step shows complete code; deferred mechanics resolved
  in the spec; Option-B fallback is intentionally describe-only (only built if
  acceptance fails). ✓

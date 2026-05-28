# Sidecar Reset-From-Template Implementation Plan

> **SUPERSEDED (2026-05-28):** Tasks below implemented an *external packaged
> template* as the reset source. That was then replaced by **internal blank
> snapshots** (`_Template_NN` sheets inside the workbook + `RebuildBlankTemplates`)
> because the external copy drifted out of date. The shipped code and the design
> doc `../specs/2026-05-28-sidecar-reset-from-template-design.md` are the source
> of truth; the task steps here are kept only as implementation history.


> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the buggy surgical post-upload clear with a whole-sheet revert from the DataViewer-packaged template, so every uploaded sheet returns to a known-good blank template state (correct seeds, intervals, formulas, formatting) with no per-cell logic.

**Architecture:** All changes are in one VBA module, `excel-sidecar/DataViewerUpload.bas`, plus a README prose update. `ResetLiveWorkbookAfterUpload` opens the packaged template read-only once and, for each uploaded canonical sheet, deletes the live sheet and replaces it with a pristine copy from the template (preserving tab position). Fail-safe: if the template can't be found/opened, nothing is cleared.

**Tech Stack:** Excel VBA (imported into `Automated Testing Template.xlsm`). Edits are applied to the `.bas` text file via Python read-replace-rewrite (the MIP-safe convention on this machine: plaintext, CRLF for VBE import). VBA cannot be unit-tested in-process, so each task verifies structurally (Sub/Function balance, CRLF, symbol presence) and the user runs the final Excel test.

**Design doc:** `docs/superpowers/specs/2026-05-28-sidecar-reset-from-template-design.md`

---

## File Structure

- **Modify:** `excel-sidecar/DataViewerUpload.bas`
  - Add `ResolveTemplatePath()` and `RestoreSheetFromTemplate()`.
  - Rewrite `ResetLiveWorkbookAfterUpload()`.
  - Remove `ClearSheetEntries`, `BlockCount`, and the `LAST_DATA_ROW` constant.
  - Update the module header comment (step "e").
- **Modify:** `excel-sidecar/README.md` — rewrite the "post-upload reset" section.

Each task ends with a structural-verification step and a commit. The module stays import-valid at every commit.

**Path constants used by every patch script:**
- BAS = `C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\DataViewerUpload.bas`
- README = `C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\README.md`
- WORKTREE = `C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar`

---

## Task 0: Create the reusable verifier

**Files:**
- Create: `C:\Users\S1134987\AppData\Local\Temp\verify_dvu.py` (scratch tool, not committed)

- [ ] **Step 1: Write the verifier** (used by every later task)

```python
# write to C:\Users\S1134987\AppData\Local\Temp\verify_dvu.py
import os
p = r"C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\DataViewerUpload.bas"
b = open(p, "rb").read()
t = b.decode("utf-8")
crlf = b.count(b"\r\n"); bare = b.count(b"\n") - crlf
subs = t.count("\nPrivate Sub ") + t.count("\nPublic Sub ") + t.count("\nSub ")
esub = t.count("\nEnd Sub")
fns = t.count("\nPrivate Function ") + t.count("\nPublic Function ") + t.count("\nFunction ")
efn = t.count("\nEnd Function")
print("CRLF=%d bareLF=%d | Sub=%d EndSub=%d (%s) | Func=%d EndFunc=%d (%s)" %
      (crlf, bare, subs, esub, "OK" if subs == esub else "MISMATCH",
       fns, efn, "OK" if fns == efn else "MISMATCH"))
for m in ["ResolveTemplatePath", "RestoreSheetFromTemplate", "ClearSheetEntries",
          "BlockCount", "LAST_DATA_ROW", "FIRST_DATA_ROW", "DV_TemplatePath"]:
    print("  %-26s count=%d" % (m, t.count(m)))
```

- [ ] **Step 2: Run it on the current file** (baseline)

Run: `py -3.13 C:\Users\S1134987\AppData\Local\Temp\verify_dvu.py`
Expected (current state): `Sub=21 EndSub=21 (OK) | Func=20 EndFunc=20 (OK)`, `ClearSheetEntries count=2`, `BlockCount count=3`, `LAST_DATA_ROW count=2`, `ResolveTemplatePath count=0`, `RestoreSheetFromTemplate count=0`, `DV_TemplatePath count=0`, `bareLF=0`.

---

## Task 1: Add `ResolveTemplatePath` + `RestoreSheetFromTemplate`

These two helpers are added but not yet called (the file stays import-valid).

**Files:**
- Modify: `excel-sidecar/DataViewerUpload.bas` (insert before `Private Sub ClearSheetEntries`)

- [ ] **Step 1: Apply the patch** (save to a temp .py and run it)

```python
import os
p = r"C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\DataViewerUpload.bas"
s = open(p, "r", encoding="utf-8").read().replace("\r\n", "\n").replace("\r", "\n")

helpers = r'''' ============================================================================
' Post-upload reset: revert each uploaded sheet to the packaged template
' ============================================================================

Private Function ResolveTemplatePath() As String
    ' Locate the DataViewer-packaged template (the "Standardized Test Template").
    ' Priority:
    '   1. DV_TemplatePath named range (explicit override) if the file exists.
    '   2. <folder of DataViewer.exe>\resources\templates\Standardized Test
    '      Template*.xlsx  - newest by modified date if several (survives the
    '      annual filename roll). DataViewer.exe is resolved exactly as the
    '      launch step resolves it (DV_DataViewerExe override or default).
    ' Returns "" if nothing suitable is found; the caller then skips the reset.
    ResolveTemplatePath = ""

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim override As String
    override = Trim$(GetNamed("DV_TemplatePath"))
    If Len(override) > 0 Then
        If fso.FileExists(override) Then
            ResolveTemplatePath = override
            Exit Function
        End If
    End If

    Dim dvExe As String
    dvExe = ResolveDataViewerExe()
    If Len(dvExe) = 0 Then Exit Function
    If Not fso.FileExists(dvExe) Then Exit Function

    Dim tdir As String
    tdir = fso.BuildPath(fso.GetParentFolderName(dvExe), "resources\templates")
    If Not fso.FolderExists(tdir) Then Exit Function

    Dim f As Object, best As String, bestDate As Date
    best = ""
    For Each f In fso.GetFolder(tdir).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "xlsx" Then
            If LCase$(f.Name) Like LCase$("Standardized Test Template*") Then
                If best = "" Or f.DateLastModified > bestDate Then
                    best = f.Path
                    bestDate = f.DateLastModified
                End If
            End If
        End If
    Next
    ResolveTemplatePath = best
End Function

Private Sub RestoreSheetFromTemplate(liveWb As Workbook, tplWb As Workbook, _
                                     sheetName As String)
    ' Replace one live data sheet with a pristine copy from the template,
    ' preserving its tab position. Transactional: the original is removed only
    ' AFTER its replacement is in place, so a mid-failure never loses the sheet
    ' (worst case a harmless duplicate remains and is logged).
    If Not WorkbookHasSheetIn(tplWb, sheetName) Then
        StampLog "  '" & sheetName & "': no packaged template - not auto-reset (left intact)."
        Exit Sub
    End If

    Dim orig As Worksheet
    Set orig = liveWb.Worksheets(sheetName)

    Dim savedAlerts As Boolean
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Fail

    ' Copy the template sheet in just before the original; the new copy becomes
    ' the active sheet. Capture it by reference, then delete the original and
    ' take over its canonical name.
    tplWb.Worksheets(sheetName).Copy Before:=orig
    Dim fresh As Worksheet
    Set fresh = liveWb.ActiveSheet

    orig.Delete
    fresh.Name = sheetName

    Application.DisplayAlerts = savedAlerts
    Exit Sub

Fail:
    Application.DisplayAlerts = savedAlerts
    StampLog "  '" & sheetName & "': reset FAILED (" & Err.Description & ") - left as-is."
End Sub

'''

anchor = "Private Sub ClearSheetEntries(ws As Worksheet)"
assert s.count(anchor) == 1, "ClearSheetEntries anchor missing/duplicate"
assert "Private Function ResolveTemplatePath" not in s, "ResolveTemplatePath already present"
assert "Private Sub RestoreSheetFromTemplate" not in s, "RestoreSheetFromTemplate already present"
s = s.replace(anchor, helpers + anchor, 1)

if os.path.exists(p):
    os.remove(p)
with open(p, "w", encoding="utf-8", newline="\r\n") as f:
    f.write(s)
print("Task 1 applied")
```

- [ ] **Step 2: Verify structurally**

Run: `py -3.13 C:\Users\S1134987\AppData\Local\Temp\verify_dvu.py`
Expected: `Sub=22 EndSub=22 (OK) | Func=21 EndFunc=21 (OK)`, `ResolveTemplatePath count=1`, `RestoreSheetFromTemplate count=1`, `DV_TemplatePath count=1`, `bareLF=0`, `ClearSheetEntries count=2` (still present).

- [ ] **Step 3: Commit**

```bash
git -C "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar" add excel-sidecar/DataViewerUpload.bas
git -C "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar" commit -m "feat(sidecar): add ResolveTemplatePath + RestoreSheetFromTemplate helpers"
```

---

## Task 2: Rewrite `ResetLiveWorkbookAfterUpload` to revert from template

Now wire the reset to the new helpers. `ClearSheetEntries` becomes unused (still defined; removed in Task 3) so the module stays import-valid.

**Files:**
- Modify: `excel-sidecar/DataViewerUpload.bas`

- [ ] **Step 1: Apply the patch**

```python
import os
p = r"C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\DataViewerUpload.bas"
s = open(p, "r", encoding="utf-8").read().replace("\r\n", "\n").replace("\r", "\n")

old = r'''Private Sub ResetLiveWorkbookAfterUpload(keep As Object)
    ' For each canonical data sheet that ended up in the keep-list, restore
    ' it from its hidden _Template_<SafeName> (if present) or hard-clear it.
    ' Then reset DV_TestSelection to default and re-hide everything.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Cleanup

    Dim sheetName As Variant
    For Each sheetName In CanonicalDataSheets()
        If keep.Exists(CStr(sheetName)) And SheetExists(CStr(sheetName)) Then
            ClearSheetEntries ThisWorkbook.Worksheets(CStr(sheetName))
        End If
    Next

    ResetSelectionToDefault
    ApplySheetVisibility

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub'''

new = r'''Private Sub ResetLiveWorkbookAfterUpload(keep As Object)
    ' Revert each uploaded canonical data sheet wholesale to the DataViewer-
    ' packaged template, then reset DV_TestSelection to default and re-hide
    ' everything so the workbook is fresh for the next session.
    '
    ' Fail-safe: if the template can't be located or opened, NOTHING is cleared.
    ' The live data is left exactly as-is and the reason is logged - we never
    ' destroy the operator's data without a known-good source to restore from.
    Dim tplWb As Workbook
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Cleanup

    Dim tplPath As String
    tplPath = ResolveTemplatePath()
    If Len(tplPath) = 0 Then
        StampLog "  WARN: template not found - sheets NOT reset (data left intact)." & _
                 " Set DV_TemplatePath or check the DataViewer install."
        GoTo Cleanup
    End If

    On Error Resume Next
    Set tplWb = Application.Workbooks.Open(fileName:=tplPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo Cleanup
    If tplWb Is Nothing Then
        StampLog "  WARN: could not open template - sheets NOT reset (data intact): " & tplPath
        GoTo Cleanup
    End If

    Dim sheetName As Variant
    For Each sheetName In CanonicalDataSheets()
        If keep.Exists(CStr(sheetName)) And SheetExists(CStr(sheetName)) Then
            RestoreSheetFromTemplate ThisWorkbook, tplWb, CStr(sheetName)
        End If
    Next

    tplWb.Close SaveChanges:=False
    Set tplWb = Nothing

    ResetSelectionToDefault
    ApplySheetVisibility

Cleanup:
    On Error Resume Next
    If Not tplWb Is Nothing Then tplWb.Close SaveChanges:=False
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub'''

assert s.count(old) == 1, "old ResetLiveWorkbookAfterUpload not found/unique"
s = s.replace(old, new)
assert s.count("RestoreSheetFromTemplate ThisWorkbook, tplWb") == 1
assert "ClearSheetEntries ThisWorkbook.Worksheets" not in s, "old ClearSheetEntries call still present"

if os.path.exists(p):
    os.remove(p)
with open(p, "w", encoding="utf-8", newline="\r\n") as f:
    f.write(s)
print("Task 2 applied")
```

- [ ] **Step 2: Verify structurally**

Run: `py -3.13 C:\Users\S1134987\AppData\Local\Temp\verify_dvu.py`
Expected: `Sub=22 EndSub=22 (OK) | Func=21 EndFunc=21 (OK)`, `ClearSheetEntries count=1` (definition only - the call is gone), `RestoreSheetFromTemplate count=2` (def + call), `bareLF=0`.

- [ ] **Step 3: Commit**

```bash
git -C "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar" add excel-sidecar/DataViewerUpload.bas
git -C "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar" commit -m "feat(sidecar): reset reverts uploaded sheets from the packaged template"
```

---

## Task 3: Remove the dead surgical-clear code

Remove `ClearSheetEntries`, `BlockCount`, and the `LAST_DATA_ROW` constant. `FIRST_DATA_ROW` stays (used by validation + trim).

**Files:**
- Modify: `excel-sidecar/DataViewerUpload.bas`

- [ ] **Step 1: Apply the patch**

```python
import os
p = r"C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\DataViewerUpload.bas"
s = open(p, "r", encoding="utf-8").read().replace("\r\n", "\n").replace("\r", "\n")

# Remove the ClearSheetEntries + BlockCount procedures (they are adjacent).
start_marker = "\n\nPrivate Sub ClearSheetEntries(ws As Worksheet)"
start = s.index(start_marker)                         # points at the blank line before the Sub
k = s.index("Private Function BlockCount", start)
end = s.index("End Function", k) + len("End Function")
removed = s[start:end]
assert "ClearSheetEntries" in removed and "BlockCount" in removed
s = s[:start] + s[end:]

# Remove the LAST_DATA_ROW constant line.
old_consts = ("Private Const FIRST_DATA_ROW As Long = 5   ' rows 1-3 metadata, row 4 headers\n"
              "Private Const LAST_DATA_ROW As Long = 115  ' 12x115 sample block (matches _Template_Master)\n")
new_consts = "Private Const FIRST_DATA_ROW As Long = 5   ' rows 1-3 metadata, row 4 headers\n"
assert s.count(old_consts) == 1, "FIRST/LAST_DATA_ROW const block not found/unique"
s = s.replace(old_consts, new_consts)

# Sanity: the surgical-clear symbols are fully gone.
assert "ClearSheetEntries" not in s
assert "BlockCount" not in s
assert "LAST_DATA_ROW" not in s
assert "FIRST_DATA_ROW" in s

if os.path.exists(p):
    os.remove(p)
with open(p, "w", encoding="utf-8", newline="\r\n") as f:
    f.write(s)
print("Task 3 applied")
```

- [ ] **Step 2: Verify structurally**

Run: `py -3.13 C:\Users\S1134987\AppData\Local\Temp\verify_dvu.py`
Expected: `Sub=21 EndSub=21 (OK) | Func=20 EndFunc=20 (OK)`, `ClearSheetEntries count=0`, `BlockCount count=0`, `LAST_DATA_ROW count=0`, `FIRST_DATA_ROW count>=1`, `ResolveTemplatePath count=1`, `RestoreSheetFromTemplate count=2`, `bareLF=0`.

- [ ] **Step 3: Commit**

```bash
git -C "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar" add excel-sidecar/DataViewerUpload.bas
git -C "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar" commit -m "refactor(sidecar): remove dead surgical-clear code (ClearSheetEntries/BlockCount/LAST_DATA_ROW)"
```

---

## Task 4: Update the module header comment + README

**Files:**
- Modify: `excel-sidecar/DataViewerUpload.bas` (header comment, step "e")
- Modify: `excel-sidecar/README.md` (the "post-upload reset" section)

- [ ] **Step 1: Patch the module header comment**

```python
import os
p = r"C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\DataViewerUpload.bas"
s = open(p, "r", encoding="utf-8").read().replace("\r\n", "\n").replace("\r", "\n")

old_e = ("'        e. resets each selected sheet to its hidden _Template_<Name>\n"
         "'           (or hard-clears if no template exists),\n")
new_e = ("'        e. reverts each selected sheet to the DataViewer-packaged template\n"
         "'           (Standardized Test Template); if the template can't be found,\n"
         "'           the reset is skipped and the data is left intact (logged),\n")
assert s.count(old_e) == 1, "header step e not found/unique"
s = s.replace(old_e, new_e)

if os.path.exists(p):
    os.remove(p)
with open(p, "w", encoding="utf-8", newline="\r\n") as f:
    f.write(s)
print("Task 4 (header) applied")
```

- [ ] **Step 2: Patch the README reset section**

```python
import os
p = r"C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\README.md"
s = open(p, "r", encoding="utf-8").read().replace("\r\n", "\n").replace("\r", "\n")

a = s.index("## The post-upload reset")
b = s.index("## Install / update the workbook")
NEW_SECTION = (
"## The post-upload reset\n"
"\n"
"After a successful upload, each uploaded sheet is **reverted wholesale to the\n"
"DataViewer-packaged template** (`Standardized Test Template ...xlsx`, found\n"
"under `resources/templates/` next to `DataViewer.exe` via the `DV_DataViewerExe`\n"
"path; override with a `DV_TemplatePath` named range). The live sheet is deleted\n"
"and replaced by a pristine copy from the template, so headers, per-sheet puff\n"
"seed/interval, formula scaffolding, formatting, and milestone notes are all\n"
"restored exactly, at the template's default sample-block count. Grow the sheet\n"
"again with Add Sample for the next campaign.\n"
"\n"
"**Fail-safe:** if the template can't be located or opened, the reset is skipped\n"
"and the live data is left intact (logged). Sheets with no packaged-template\n"
"counterpart (e.g. `Custom Test Template`) are likewise skipped + logged. Nothing\n"
"is ever cleared without a known-good source to restore from.\n"
"\n"
"The earlier surgical `ClearSheetEntries` (and the per-sheet `_Template_<X>` /\n"
"hard-clear before it) are gone -- see\n"
"`../docs/superpowers/specs/2026-05-28-sidecar-reset-from-template-design.md`.\n"
"The template cell map is documented at\n"
"`../docs/superpowers/specs/template-cell-map.md`.\n"
"\n"
)
assert a < b, "section markers out of order"
s = s[:a] + NEW_SECTION + s[b:]

# Repoint the stale top-of-file reference to the superseding design doc.
s = s.replace("2026-05-27-sidecar-post-upload-reset.md",
              "2026-05-28-sidecar-reset-from-template-design.md")

if os.path.exists(p):
    os.remove(p)
with open(p, "w", encoding="utf-8", newline="\n") as f:
    f.write(s)
print("Task 4 (README) applied")
```

- [ ] **Step 3: Verify**

Run: `py -3.13 C:\Users\S1134987\AppData\Local\Temp\verify_dvu.py`
Expected: unchanged from Task 3 (`Sub=21 EndSub=21 (OK) | Func=20 EndFunc=20 (OK)`, `bareLF=0`).

Run: `grep -n "reverted wholesale" "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar/excel-sidecar/README.md"`
Expected: one match. Also confirm `grep -c "ClearSheetEntries" .../README.md` returns 0.

- [ ] **Step 4: Commit**

```bash
git -C "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar" add excel-sidecar/DataViewerUpload.bas excel-sidecar/README.md
git -C "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/sidecar" commit -m "docs(sidecar): describe reset-from-template in module header + README"
```

---

## Task 5: Final verification + manual test recipe

**Files:** none (verification only)

- [ ] **Step 1: Full structural check**

Run: `py -3.13 C:\Users\S1134987\AppData\Local\Temp\verify_dvu.py`
Expected: `Sub=21 EndSub=21 (OK) | Func=20 EndFunc=20 (OK)`, `bareLF=0`, `ClearSheetEntries=0`, `BlockCount=0`, `LAST_DATA_ROW=0`, `ResolveTemplatePath=1`, `RestoreSheetFromTemplate=2`, `DV_TemplatePath=1`, `FIRST_DATA_ROW>=1`.

- [ ] **Step 2: Confirm the file is plaintext (not MIP-encrypted)**

Run: `py -3.13 -c "import sys; b=open(r'C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\sidecar\excel-sidecar\DataViewerUpload.bas','rb').read(40); print(b[:20]); print('PLAINTEXT' if b.lstrip().startswith(b'Attribute VB_Name') else 'CHECK')"`
Expected: starts with `Attribute VB_Name` -> `PLAINTEXT`. (If it shows `%TSD-Header`, run the repo's MIP decrypt before re-importing.)

- [ ] **Step 3: Hand off to the user for the Excel test**

VBA can't be run here. Give the user this recipe:
1. Re-import `DataViewerUpload.bas` into `Automated Testing Template.xlsm` (VBE -> remove old `DataViewerUpload` -> Import File), or run `python excel-sidecar/install_sidecar.py --file "<path>" --apply`.
2. Make sure `DV_DataViewerExe` (cell I12) points at the real `DataViewer.exe` (the template is found beside it under `resources\templates\`).
3. Select a few tests including an **expanded** sheet (e.g. Lifetime Test grown to 12 sample blocks) and **Temperature Cycling Test #1**. Enter a little data. Click **Upload All**.
4. After the run, confirm on each uploaded sheet:
   - Block count is back to the template default (e.g. Lifetime = 2 blocks).
   - Row 5 is intact: `A5` shows the puff seed (5 / 20 / 1 per sheet); the `A6 = A5 + interval` chain is present.
   - Milestone notes / `B5` seeds present where the template has them (e.g. Negative Pressure `H5` = "INITIAL 10 PUFFS").
   - `Temperature Cycling Test #1` was reset too (it was skipped before).
5. Confirm selection reset to Lifetime Test only, and everything re-hid except Lifetime Test / Test SOP's / DataViewer Upload.
6. **Fail-safe check:** set `DV_TemplatePath` to a path that doesn't exist, upload again -> the log shows "template not found - sheets NOT reset (data left intact)" and the data is untouched. Clear `DV_TemplatePath` afterward.
7. If `Custom Test Template` was uploaded, confirm it was left intact with a "no packaged template" log line.

- [ ] **Step 4: (optional) run the drift detector**

`python excel-sidecar/verify_sidecar.py --file "<path to Automated Testing Template.xlsm>"`
Before re-importing, `DataViewerUpload` will report `DIFFERS` (deployed still has the old reset) - that's expected. After re-importing, it should report `MATCH`.

---

## Self-Review (completed by plan author)

- **Spec coverage:** template source + runtime resolution (Task 1 `ResolveTemplatePath`); whole-sheet transactional replace (Task 1 `RestoreSheetFromTemplate`); reset flow + fail-safe on missing template (Task 2); removals (Task 3); header + README prose (Task 4); structural + manual verification incl. fail-safe and Custom Test Template edge (Task 5). All design-doc sections map to a task.
- **Placeholders:** none - every step has exact code or exact commands + expected output.
- **Type/name consistency:** `ResolveTemplatePath()` / `RestoreSheetFromTemplate(liveWb, tplWb, sheetName)` defined in Task 1 are called with matching signatures in Task 2. Reused existing helpers `GetNamed`, `ResolveDataViewerExe`, `WorkbookHasSheetIn`, `StampLog`, `SheetExists`, `CanonicalDataSheets`, `ResetSelectionToDefault`, `ApplySheetVisibility` (all verified present in the module).

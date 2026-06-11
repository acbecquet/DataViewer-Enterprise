# Excel Sidecar Template v1.2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Resolve every finding from the v1.1 production audit (`docs/superpowers/specs/2026-06-10-excel-sidecar-v11-production-audit.md`) and ship template v1.2 with the owner's new features: Upload Checkpoint, row-5 puff seeding, lockbox sheet restore, ribbon-free distributed copies, and VBA signing.

**Architecture:** All tester-facing behavior lives in three VBA modules + ThisWorkbook class (source of truth in `excel-sidecar/*.bas`, imported into the .xlsm by `build_clean_template.py` via Excel COM). Headless gates (`check_sources.py`, `test_build_helpers.py`) run anywhere; the build + `verify_sidecar.py` run on this machine (Excel + pywin32 present). VBA cannot be compiled here — correctness is enforced by the headless gates + the final build/verify + the operator acceptance checklist.

**Tech Stack:** VBA (Excel), Python 3 (zipfile/openpyxl/pywin32), Ribbon XML (customUI14), PowerShell (signing kit).

**Owner's design rules (binding, from memory `excel-sidecar-product-vision`):**
1. **Poka-yokes, never gates** — warn + confirm; never hard-refuse a tester action unless data loss would be silent.
2. **Checkpoint stream** — same-name destination overwrite is INTENTIONAL for a workbook's own checkpoint series; only cross-stream collisions get a confirm.
3. **Lockbox** — `_Template_NN` snapshots restore tester-mangled canonical sheets; tester mutations never break upload.
4. Distributed copies must be macro-free AND ribbon-free.

**Conventions for executors:**
- Worktree: `C:\Users\S1134987\.config\superpowers\worktrees\DataViewer-Enterprise\excel-sidecar-v1.2`, branch `feature/excel-sidecar-v1.2` (created in Task 0). All paths below are relative to the worktree root.
- NEW files must be written via Python delete-and-rewrite (MIP labels — see CLAUDE.md). Editing existing plaintext files with the Edit tool is fine.
- After every task: run `"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/check_sources.py` → must end `RESULT: ALL PASS` (Task 6 updates the checker itself; Tasks 1–5 must not break existing checks).
- VBA style: match the modules' existing comment density and the `Saved/restore` state patterns. `Option Explicit` everywhere; no new globals beyond those specified.
- Commit after each task with the message given.

---

### Task 0: Worktree + branch (orchestrator does this inline)

- [ ] `git worktree add "C:/Users/S1134987/.config/superpowers/worktrees/DataViewer-Enterprise/excel-sidecar-v1.2" -b feature/excel-sidecar-v1.2` from `feature/v2.4.0-bugfix-batch` (f2bea50 or later).

---

### Task 1: TestingTools.bas — Reset/Add/Remove/puff-picker fixes (audit C1, C4, H1, H2, H3, M-h, M-i + row-5 seeding)

**Files:**
- Modify: `excel-sidecar/TestingTools.bas`

- [ ] **Step 1: Fix C1 + H3 in `ResetEquations`** (currently lines 123–160). Replace the specs array and the copy loop. Drop relCols 1 and 2 (puffs + before-weight: testers legitimately type literals there; restoring them imposed master's interval 20 — audit H3) and use `FormulaR1C1` so block ≥ 2 formulas stay position-relative instead of cross-wiring to block 1's columns (audit C1):

```vba
    ' Formula cell ranges within a block: (relCol, rowStart, rowEnd).
    ' Cols 1 (puffs) and 2 (before weight) are deliberately NOT restored: testers
    ' type literals there (custom puff sequences, re-weighs) and the master would
    ' impose its own interval. Re-seed puffs by typing a value in row 5 instead.
    Dim specs As Variant
    specs = Array(Array(6, 2, 2), Array(9, 3, 3), Array(9, 5, 115), _
                  Array(10, 5, 115), Array(11, 6, 115), Array(12, 2, 2), Array(12, 3, 3), Array(12, 5, 115), _
                  Array(7, 5, 115))

    Dim spec As Variant, relCol As Long, r1 As Long, r2 As Long
    For Each spec In specs
        relCol = spec(0): r1 = spec(1): r2 = spec(2)
        ' FormulaR1C1 keeps relative references position-independent: restoring
        ' block N must reference block N's own columns, not block 1's (audit C1).
        ws.Range(ws.Cells(r1, startCol + relCol - 1), ws.Cells(r2, startCol + relCol - 1)).FormulaR1C1 = _
            master.Range(master.Cells(r1, relCol), master.Cells(r2, relCol)).FormulaR1C1
    Next spec
```

- [ ] **Step 2: Add an error handler to `ResetEquations` so manual-calc can't leak** (audit Low). Wrap the mutation section: after the confirm `MsgBox`, replace `Application.ScreenUpdating = False` / `Application.Calculation = xlCalculationManual` with:

```vba
    On Error GoTo ResetFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
```

and after the success `MsgBox ... "Your data is intact."` add before `End Sub`:

```vba
    Exit Sub
ResetFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Reset Formulas hit an error and stopped: " & Err.Description, vbExclamation, "TPM Testing"
```

- [ ] **Step 3: Fix H2 in `AddSample`** — destination-emptiness confirm (poka-yoke, not a gate) + the same calc-leak handler. After `destStartCol` is computed (line 78) and BEFORE the manual-calc/copy section, insert (and move the `ScreenUpdating/Calculation` lines AFTER this check so a "No" leaves state untouched):

```vba
    ' POKA-YOKE (audit H2): if a row-4 header was edited, CountBlocks undercounts
    ' and the "new" block would paste over live data. Confirm instead of wiping.
    Dim destUsed As Long
    destUsed = Application.WorksheetFunction.CountA( _
        ws.Range(ws.Cells(1, destStartCol), ws.Cells(BLOCK_ROWS, destStartCol + BLOCK_COLS - 1)))
    If destUsed > 0 Then
        If MsgBox("The destination columns " & ColLetter(destStartCol) & ":" & _
                  ColLetter(destStartCol + BLOCK_COLS - 1) & " already contain " & destUsed & _
                  " filled cell(s)." & vbCrLf & "(A sample header in row 4 may have been " & _
                  "edited, making the block count wrong.)" & vbCrLf & vbCrLf & _
                  "Overwrite them with a blank sample block? This cannot be undone.", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "TPM Testing - Add Sample") <> vbYes Then Exit Sub
    End If
    On Error GoTo AddFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
```

and before `End Sub` of AddSample:

```vba
    Exit Sub
AddFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Add Sample hit an error and stopped: " & Err.Description, vbExclamation, "TPM Testing"
```

- [ ] **Step 4: Fix M-i in `RemoveSample`** — show whether the doomed block holds data. Count constants only (a pristine block carries ~700 template formulas, so a plain CountA would always warn and the "empty" branch would be dead). After `titleVal` is read (line 109), insert:

```vba
    ' Constants only: a pristine block carries ~700 template formulas, so a
    ' plain CountA would always warn and the "empty" message could never show.
    Dim filled As Long
    On Error Resume Next
    filled = ws.Range(ws.Cells(5, startCol), ws.Cells(BLOCK_ROWS, endCol)) _
                .SpecialCells(xlCellTypeConstants).Count
    On Error GoTo 0          ' no constants -> 1004 -> filled stays 0
```

(AddSample's `destUsed` CountA stays a plain CountA on purpose — there, formulas legitimately mean "destination occupied".)

and change the confirm body's `"Title: " & titleVal & vbCrLf & vbCrLf` line to:

```vba
              "Title: " & titleVal & vbCrLf & _
              IIf(filled > 0, "THIS BLOCK CONTAINS " & filled & " FILLED DATA CELL(S).", _
                  "This block is empty.") & vbCrLf & vbCrLf & _
```

- [ ] **Step 5: Fix M-h in `IsAllowedSheet`** — also reject `_`-prefixed sheets (snapshots/_Settings pass the structural check if ever unhidden). After the `" - Review"` check (line 20), insert:

```vba
    If Left$(ws.Name, 1) = "_" Then GoTo NotAllowed
```

- [ ] **Step 6: Port the deployed drift (C4) + row-5 seeding + H1 guards into `TryPuffStepPicker`.** Replace the function body from the `Dim v As String` line (185) through the `End If` before `TryPuffStepPicker = True` (217) with:

```vba
    Dim v As String
    ' Cell error values (=1/0 etc.) would raise type mismatch in CStr below;
    ' an error value can never be a step or "custom".
    If IsError(pick.value) Then Exit Function
    v = LCase$(Trim$(CStr(pick.value)))
    Dim stp As Variant
    stp = Empty
    Select Case v
        Case "custom": stp = "custom"
        Case Else
            ' v1.2 (ports the owner's in-workbook v1.1 fix): ANY positive number
            ' is a step. Poka-yokes, never gates -- testers may use any interval.
            ' Nested (VBA And does not short-circuit; CDbl on text would raise 13).
            If Not IsNumeric(v) Then Exit Function
            If CDbl(v) <= 0 Then Exit Function
            stp = CDbl(v)
    End Select

    Dim savedEvents As Boolean
    savedEvents = Application.EnableEvents
    On Error GoTo PuffDone
    Application.EnableEvents = False

    ' The legacy step-list dropdown would reject free-form values; the puffs
    ' column is free-form by design now (owner decision, v1.2) -- both for
    ' numeric steps and for 'custom' hand-entered sequences.
    On Error Resume Next
    Sh.Range(Sh.Cells(5, c), Sh.Cells(BLOCK_ROWS, c)).Validation.Delete
    On Error GoTo PuffDone

    If VarType(stp) = vbString Then        ' "custom"
        pick.ClearContents
        pick.Select
    Else
        Dim startFill As Long
        If pick.row = 5 Then
            ' Row-5 seeding (v1.2): the first puff value also becomes the step --
            ' rows below fill with prev + value, same as typing on a later row.
            Sh.Cells(5, c).value = stp
            startFill = 6
        Else
            startFill = pick.row
        End If
        ' Single range write: each cell = the cell above + step. R1C1 keeps the
        ' reference relative per-row; Str$ guarantees a period decimal separator.
        Sh.Range(Sh.Cells(startFill, c), Sh.Cells(BLOCK_ROWS, c)).FormulaR1C1 = _
            "=R[-1]C+" & Trim$(Str$(stp))
    End If
    TryPuffStepPicker = True
```

Also add the H1 guards right after the `TypeName(Sh) <> "Worksheet"` check (line 170):

```vba
    ' Never run on archived Review copies or hidden plumbing (audit H1): they
    ' share the test-sheet layout but must not be mutated by the step picker.
    If InStr(1, Sh.Name, " - Review", vbTextCompare) > 0 Then Exit Function
    If Left$(Sh.Name, 1) = "_" Then Exit Function
```

- [ ] **Step 7: Run the gate.** `"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/check_sources.py` → `RESULT: ALL PASS`.

- [ ] **Step 8: Commit.** `git add excel-sidecar/TestingTools.bas && git commit -m "fix(sidecar): R1C1 Reset Formulas, any-number puff step + row-5 seed, Review/_ guards, Add/Remove poka-yokes (v1.2 T1)"`

---

### Task 2: DataViewerUpload.bas — rename safety + template marker (audit C2, M-e, P2)

**Files:**
- Modify: `excel-sidecar/DataViewerUpload.bas` (the `RenameWorkbookTo` / `Btn_SpecifyName` region, lines 1480–1529)

- [ ] **Step 1: Replace `RenameWorkbookTo` entirely with:**

```vba
Private Sub RenameWorkbookTo(ByVal newFullPath As String)
    Dim oldPath As String: oldPath = ThisWorkbook.FullName
    If StrComp(oldPath, newFullPath, vbTextCompare) = 0 Then Exit Sub
    On Error GoTo Fail
    ' POKA-YOKE (audit C2): never silently SaveAs over an existing workbook.
    ' Parallel-project copies share DV_OrigFileName, so the auto-revert path
    ' could otherwise destroy a sibling copy with alerts suppressed.
    If Len(Dir$(newFullPath)) > 0 Then
        If MsgBox("A file already exists at:" & vbCrLf & newFullPath & vbCrLf & vbCrLf & _
                  "Overwrite it? If another project copy uses this name, choose No " & _
                  "and pick a different test name.", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Rename workbook") <> vbYes Then Exit Sub
    End If
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=newFullPath, FileFormat:=52   ' 52 = xlOpenXMLWorkbookMacroEnabled
    If Len(Dir$(oldPath)) > 0 Then
        On Error Resume Next
        Kill oldPath                                            ' clean rename
        If Err.Number <> 0 Then
            ' Audit M-e: a swallowed Kill leaves two near-identical .xlsm files;
            ' a tester can enter data into the stale one. Say so out loud.
            StampLog "WARN: rename left old file behind: " & oldPath & " (" & Err.Description & ")"
            MsgBox "The workbook was renamed, but the previous file could not be removed:" & _
                   vbCrLf & oldPath & vbCrLf & vbCrLf & _
                   "Delete it manually so only one copy exists.", vbExclamation, "Rename workbook"
            Err.Clear
        End If
        On Error GoTo Fail
    End If
    Application.DisplayAlerts = True
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    MsgBox "Could not rename the file:" & vbCrLf & Err.Description, vbExclamation, "Specify Test Name"
End Sub
```

- [ ] **Step 2: Template marker (P2).** In `Btn_SpecifyName`, change the target construction (lines 1523–1527) to:

```vba
    If Len(proj) = 0 Then
        target = orig & ".xlsm"
    Else
        ' "(do not send)" marks the renamed working copy as a non-deliverable at
        ' the exact place the send-the-template mistake happens: the Explorer /
        ' Outlook attach dialog (audit P2). Upload All still reverts to orig.
        target = proj & " - " & orig & " (do not send).xlsm"
    End If
```

(`RevertToOriginalName` already targets `orig & ".xlsm"`, so the marker disappears on revert — no further change.)

- [ ] **Step 3: Run gate + commit.** `check_sources.py` → ALL PASS. `git add excel-sidecar/DataViewerUpload.bas && git commit -m "fix(sidecar): rename exists-confirm + loud Kill failure + '(do not send)' marker (v1.2 T2, audit C2/M-e/P2)"`

---

### Task 3: DataViewerUpload.bas — upload engine v1.2 (audit C3, H4, H5, H9, M-a, M-b, M-c, M-l, P3, P4, P7 + Upload Checkpoint + lockbox)

This is the big one: one coherent rewrite of the upload path. One agent, complete code below.

**Files:**
- Modify: `excel-sidecar/DataViewerUpload.bas`

- [ ] **Step 1: Module-level state.** Below `Public gRibbon As IRibbonUI` (line 67) add:

```vba
' True once this session has successfully delivered data (Upload All or
' Checkpoint, or copies-landed-but-exe-missing). Drives the BeforeClose nudge.
Private gUploadedThisSession As Boolean
```

- [ ] **Step 2: Replace `Btn_UploadAll` (lines 290–443) with the shared engine + two thin entry points.** Complete replacement code:

```vba
Public Sub Btn_UploadAll()
    RunUpload False
End Sub

Public Sub Btn_UploadCheckpoint()
    RunUpload True
End Sub

Private Sub RunUpload(ByVal asCheckpoint As Boolean)
    ' Shared engine for Upload All (resets sheets afterwards, keeps Review
    ' copies) and Upload Checkpoint (identical delivery, NO reset -- the owner's
    ' checkpoint-stream model: re-using the same file name overwrites the
    ' previous checkpoint; a new name creates a new file).
    Dim actionName As String
    actionName = IIf(asCheckpoint, "Upload Checkpoint", "Upload All")
    On Error GoTo Failed
    ClearLog
    SetNamed "DV_Status", "Starting " & actionName & "..."
    StampLog actionName & " started"

    ' v1.1: restore the original on-disk file name before uploading.
    RevertToOriginalName

AskName:
    Dim fName As String
    fName = PromptForFileName()
    If Len(fName) = 0 Then
        SetNamed "DV_Status", "Cancelled (no file name)"
        StampLog actionName & " cancelled - no file name entered"
        Exit Sub
    End If
    If fName Like "*[" & "\/:*?<>|" & Chr$(34) & "]*" Then
        SetNamed "DV_Status", "Cancelled (invalid file name)"
        MsgBox "The file name can't contain any of these characters:" & vbLf & _
               "   \ / : * ? " & Chr$(34) & " < > |", vbExclamation, "Upload file name"
        GoTo AskName
    End If
    ' P3 (findability + collision safety): a name with no date gets today's date.
    ' The prompt pre-fills the LAST name (date included), so a checkpoint stream
    ' keeps overwriting its own file across days, by design.
    If Not HasDateToken(fName) Then
        fName = fName & " - " & Format$(Date, "yyyy-mm-dd")
        StampLog "Date appended to file name: " & fName
    End If
    SetNamed "DV_FileName", fName

    ' --- 1. Checklist (abort on failure), then warnings (confirm to continue) ---
    Dim failures As Collection
    Set failures = RunChecklist()
    If failures.Count > 0 Then
        WriteFailures failures
        SetNamed "DV_Status", "Failed: checklist has " & failures.Count & " issue(s)"
        ShowFailures actionName, failures
        Exit Sub
    End If
    StampLog "Checklist passed"

    Dim warns As Collection
    Set warns = CollectUploadWarnings()
    If warns.Count > 0 Then
        Dim wmsg As String, wv As Variant
        For Each wv In warns
            wmsg = wmsg & vbLf & "  - " & CStr(wv)
            StampLog "  WARN: " & CStr(wv)
        Next
        If MsgBox("Heads up:" & vbLf & wmsg & vbLf & vbLf & "Continue with " & actionName & "?", _
                  vbYesNo + vbExclamation, actionName) <> vbYes Then
            SetNamed "DV_Status", "Cancelled at warnings"
            Exit Sub
        End If
    End If

    Dim baseName As String, synPath As String, locPath As String
    baseName = Trim$(GetNamed("DV_FileName"))
    synPath = Trim$(GetNamed("DV_SynologyPath"))
    locPath = Trim$(GetNamed("DV_LocalPath"))

    ' --- 2. Persist in-memory edits so the copies pick them up ---
    StampLog "Saving workbook"
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim synDest As String, locDest As String
    synDest = AppendBackslash(synPath) & baseName & ".xlsx"
    locDest = AppendBackslash(locPath) & baseName & ".xlsx"

    If Not fso.FolderExists(synPath) Then
        SetNamed "DV_Status", "Failed: Synology path not accessible: " & synPath
        MsgBox "Synology path not accessible:" & vbLf & synPath & vbLf & vbLf & _
               "Nothing was uploaded and nothing is lost. Fix the connection (or " & _
               "re-pick the folder via 'Pick Synology Folder') and run " & actionName & _
               " again." & vbLf & "Do NOT email this workbook as a workaround.", _
               vbExclamation, actionName
        Exit Sub
    End If
    If Not fso.FolderExists(locPath) Then
        SetNamed "DV_Status", "Failed: Local path not accessible: " & locPath
        MsgBox "Local path not accessible:" & vbLf & locPath & vbLf & vbLf & _
               "Nothing was uploaded and nothing is lost. Re-pick the folder via " & _
               "'Pick Local Folder' and run " & actionName & " again." & vbLf & _
               "Do NOT email this workbook as a workaround.", vbExclamation, actionName
        Exit Sub
    End If

    ' --- 2b. Cross-stream collision check (audit C3). Overwriting OUR OWN last
    '         upload is the checkpoint stream working as designed; overwriting a
    '         file this workbook did not produce gets a confirm. ---
    Dim lastUp As String
    lastUp = Trim$(GetNamed("DV_LastUpload"))
    If (fso.FileExists(synDest) Or fso.FileExists(locDest)) And _
       StrComp(baseName, lastUp, vbTextCompare) <> 0 Then
        If MsgBox("A file named '" & baseName & ".xlsx' already exists in the " & _
                  "destination folder(s), and it was not the last file uploaded from " & _
                  "this workbook (that was: " & IIf(Len(lastUp) > 0, lastUp, "none") & ")." & _
                  vbLf & vbLf & "Overwrite it? Choose No to enter a different name.", _
                  vbYesNo + vbExclamation + vbDefaultButton2, actionName) <> vbYes Then
            GoTo AskName
        End If
        StampLog "Operator confirmed overwrite of existing '" & baseName & ".xlsx'"
    End If

    ' --- 3a. Build keep-list (selected + populated + Test SOP's) ---
    Dim keep As Object
    Set keep = BuildKeepList()
    StampLog "Trim keep-list size: " & keep.Count
    If keep.Count <= 1 Then
        SetNamed "DV_Status", "Failed: no selected sheets contain data"
        StampLog "Aborting: keep-list contains only Test SOP's"
        MsgBox "No selected sheets contain data.", vbExclamation, actionName
        Exit Sub
    End If

    ' --- 3b. Stage a trimmed .xlsm in TEMP (intermediate only) ---
    Dim trimmedXlsm As String
    trimmedXlsm = Environ$("TEMP") & "\dvupload_staging_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & baseName & ".xlsm"
    fso.CopyFile ThisWorkbook.FullName, trimmedXlsm, True
    StampLog "Staging trimmed copy: " & trimmedXlsm
    TrimSheetsInWorkbook trimmedXlsm, keep

    ' --- 3c. Materialize ONE clean .xlsx: macro-free AND ribbon-free (audit H9;
    '         the sheets are copied into a fresh workbook, so no customUI part
    '         survives to throw "Cannot run the macro" popups for receivers). ---
    Dim cleanXlsx As String
    cleanXlsx = MakeTempXlsx(fso, trimmedXlsm, baseName)
    StampLog "Clean .xlsx: " & cleanXlsx

    On Error Resume Next
    fso.DeleteFile trimmedXlsm, True
    On Error GoTo Failed

    ' --- 3d. Distribute to both destinations ---
    StampLog "Copy -> " & synDest
    fso.CopyFile cleanXlsx, synDest, True
    StampLog "Copy -> " & locDest
    fso.CopyFile cleanXlsx, locDest, True
    SetNamed "DV_LastUpload", baseName
    gUploadedThisSession = True

    ' Staging artifacts are no longer needed: DataViewer ingests the Synology
    ' copy (audit H5), so the per-run TEMP dir can go immediately (audit M-l/11).
    On Error Resume Next
    fso.DeleteFolder Left$(cleanXlsx, InStrRev(cleanXlsx, "\") - 1), True
    On Error GoTo Failed

    ' --- 4. Launch DataViewer ON THE SYNOLOGY COPY (audit H5): the DB file_path
    '        must point at a durable location, and later in-app edits must land
    '        in the distributed file, not an orphaned TEMP copy. ---
    Dim dvExe As String
    dvExe = ResolveDataViewerExe()
    If Not fso.FileExists(dvExe) Then
        SetNamed "DV_Status", "Partial: copies delivered; DataViewer not launched"
        MsgBox "Your data WAS copied to both folders:" & vbLf & _
               "  " & synDest & vbLf & "  " & locDest & vbLf & vbLf & _
               "But DataViewer.exe was not found at:" & vbLf & "  " & dvExe & vbLf & vbLf & _
               "Use 'Pick DataViewer File' on the ribbon, then run " & actionName & _
               " again with the SAME file name (it overwrites the copies; nothing " & _
               "is duplicated)." & vbLf & "Do NOT email this workbook as a workaround.", _
               vbExclamation, actionName
        Exit Sub
    End If

    Dim cmd As String
    cmd = """" & dvExe & """ """ & synDest & """"
    StampLog "Shell: " & cmd
    Shell cmd, vbNormalFocus
    StampLog "Upload dispatched (DB write happens inside DataViewer)"

    If asCheckpoint Then
        SetNamed "DV_Status", "OK (checkpoint)"
        StampLog "Checkpoint done (no reset)"
        ShowReceipt actionName, baseName, synDest, _
                    "Your sheets were NOT reset - keep testing and upload again any " & _
                    "time. Re-using the same file name overwrites this checkpoint " & _
                    "with the newer data; a new name creates a new file."
        Exit Sub
    End If

    ' --- 5. Upload All only: reset the LIVE workbook ---
    On Error GoTo PostDispatchFailed
    StampLog "Resetting live workbook"
    Dim skipped As Collection
    Set skipped = ResetLiveWorkbookAfterUpload(keep)
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    SetNamed "DV_Status", "OK"
    StampLog "Done"
    Dim extra As String
    extra = "Each uploaded sheet was reset (a '- Review' copy was kept)."
    If skipped.Count > 0 Then
        extra = extra & vbLf & "NOTE - these sheets could NOT be reset and still " & _
                "hold their data (this is safe; see the log):"
        Dim sk As Variant
        For Each sk In skipped
            extra = extra & vbLf & "  - " & CStr(sk)
        Next
    End If
    ShowReceipt actionName, baseName, synDest, extra
    Exit Sub

PostDispatchFailed:
    ' Audit H4a: the upload already succeeded; never fail silently here, or the
    ' tester re-uploads (duplicate ingest) or falls back to emailing the file.
    StampLog "WARN: post-upload reset failed: " & Err.Description
    SetNamed "DV_Status", "OK (upload delivered; live reset partial - see log)"
    MsgBox "Your data WAS uploaded successfully (Synology + Local + DataViewer)." & _
           vbLf & vbLf & "But the automatic sheet reset did not complete (" & _
           Err.Description & "). Do NOT upload again - the data is already " & _
           "delivered. Sheets still holding data can be left for the next upload.", _
           vbExclamation, actionName
    Exit Sub

Failed:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    StampLog "ERROR " & Err.Number & ": " & Err.Description
    SetNamed "DV_Status", "Failed: " & Err.Description
    MsgBox actionName & " failed:" & vbLf & Err.Description & vbLf & vbLf & _
           "Nothing was lost. Fix the issue and run " & actionName & " again." & vbLf & _
           "Do NOT email this workbook as a workaround.", vbCritical, actionName
End Sub
```

- [ ] **Step 3: Add the new helpers** (place after `PromptForFileName`):

```vba
Private Function HasDateToken(ByVal s As String) As Boolean
    ' True if the name already carries a recognizable date (yyyy-mm-dd or
    ' d-m-yyyy style, with -, / or . separators).
    Dim pats As Variant, p As Variant, sep As Variant
    pats = Array("####-##-##", "#-#-####", "#-##-####", "##-#-####", "##-##-####")
    For Each p In pats
        For Each sep In Array("-", "/", ".")
            If s Like "*" & Replace(CStr(p), "-", CStr(sep)) & "*" Then
                HasDateToken = True
                Exit Function
            End If
        Next sep
    Next p
End Function

Private Sub ShowReceipt(ByVal title As String, ByVal baseName As String, _
                        ByVal synDest As String, ByVal extra As String)
    ' P4: the success popup is a delivery receipt -- it tells the tester the data
    ' is ALREADY delivered, where the shareable copy lives, and offers to open
    ' the folder. This is the primary defense against emailing the template.
    Dim msg As String
    msg = "DONE - your data is already delivered." & vbLf & vbLf & _
          baseName & ".xlsx is now:" & vbLf & _
          "  - in the shared Synology folder," & vbLf & _
          "  - in your local folder," & vbLf & _
          "  - open in DataViewer (saved to the shared database)." & vbLf & vbLf & _
          extra & vbLf & vbLf & _
          "NEVER email or share THIS workbook - it is your reusable template." & vbLf & _
          "If someone needs the data, send the uploaded copy:" & vbLf & _
          "  " & synDest & vbLf & vbLf & _
          "Open the Synology folder now?"
    If MsgBox(msg, vbYesNo + vbInformation, title) = vbYes Then
        On Error Resume Next
        Shell "explorer.exe /select,""" & synDest & """", vbNormalFocus
        On Error GoTo 0
    End If
End Sub

Private Function CollectUploadWarnings() As Collection
    ' Poka-yoke warnings (audit M-a + orphan sheets) -- shown once, tester can
    ' always continue. Never a gate.
    Dim warns As New Collection
    Dim selection As Object
    Set selection = ReadSelection()
    Dim sheetName As Variant, norm As String, ticked As Boolean
    For Each sheetName In CanonicalDataSheets()
        norm = NormalizeSheetName(CStr(sheetName))
        ticked = False
        If selection.Exists(norm) Then ticked = selection(norm)
        If Not ticked Then
            If SheetExists(CStr(sheetName)) Then
                If SheetHasPopulatedSamples(ThisWorkbook.Worksheets(CStr(sheetName))) Then
                    warns.Add CStr(sheetName) & " contains data but its box is NOT " & _
                              "ticked - it will NOT be uploaded"
                End If
            End If
        End If
    Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsOrphanTestSheet(ws) Then
            If SheetHasPopulatedSamples(ws) Then
                warns.Add "'" & ws.Name & "' looks like a test sheet but is not one " & _
                          "of the standard tests (renamed or copied?) - it will NOT " & _
                          "be uploaded"
            End If
        End If
    Next
    Set CollectUploadWarnings = warns
End Function

Private Function IsOrphanTestSheet(ws As Worksheet) As Boolean
    ' A visible sheet with the 'puffs' block layout that is neither canonical,
    ' nor a Review copy, nor plumbing: a tester-made rename or copy.
    On Error GoTo NoOrphan
    If ws.Visible <> xlSheetVisible Then GoTo NoOrphan
    If Left$(ws.Name, 1) = "_" Then GoTo NoOrphan
    If InStr(1, ws.Name, " - Review", vbTextCompare) > 0 Then GoTo NoOrphan
    If StrComp(ws.Name, UPLOAD_SHEET_NAME, vbTextCompare) = 0 Then GoTo NoOrphan
    If StrComp(ws.Name, SOPS_SHEET_NAME, vbTextCompare) = 0 Then GoTo NoOrphan
    Dim s As Variant
    For Each s In CanonicalDataSheets()
        If StrComp(NormalizeSheetName(ws.Name), NormalizeSheetName(CStr(s)), vbTextCompare) = 0 Then GoTo NoOrphan
    Next
    IsOrphanTestSheet = (LCase$(Trim$(SafeString(ws.Cells(4, 1).value))) = "puffs")
    Exit Function
NoOrphan:
    IsOrphanTestSheet = False
End Function

Public Function HasUnuploadedData() As Boolean
    ' Drives the BeforeClose nudge (P6). False once this session delivered data.
    If gUploadedThisSession Then Exit Function
    Dim s As Variant
    For Each s In CanonicalDataSheets()
        If SheetExists(CStr(s)) Then
            If SheetHasPopulatedSamples(ThisWorkbook.Worksheets(CStr(s))) Then
                HasUnuploadedData = True
                Exit Function
            End If
        End If
    Next
End Function
```

- [ ] **Step 4: H9 — replace `MakeTempXlsx` (lines 1381–1410) with the fresh-workbook version** (+ M-l cleanup handler):

```vba
Private Function MakeTempXlsx(fso As Object, sourceXlsm As String, _
                              baseName As String) As String
    ' Distributed-copy materializer. v1.2 (audit H9): the staged sheets are
    ' copied into a FRESH workbook before SaveAs -- a new workbook carries no
    ' customUI part and no VBA, so receivers get a plain .xlsx with no dead
    ' "TPM Testing" tab throwing "Cannot run the macro" popups.
    ' Per-run subdirectory so the bare filename is exactly "<baseName>.xlsx"
    ' (DataViewer derives the DB filename from the path).
    Dim runDir As String, outXlsx As String
    runDir = Environ$("TEMP") & "\dvupload_" & Format$(Now, "yyyymmdd_hhnnss")
    If Not fso.FolderExists(runDir) Then fso.CreateFolder runDir
    outXlsx = runDir & "\" & baseName & ".xlsx"

    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wb As Workbook, clean As Workbook
    On Error GoTo Fail
    Set wb = Application.Workbooks.Open(fileName:=sourceXlsm, ReadOnly:=True, _
                                        UpdateLinks:=0)
    On Error Resume Next
    Application.Windows(wb.Name).Visible = False
    On Error GoTo Fail

    wb.Sheets.Copy                       ' all staged sheets -> a brand-new workbook
    Set clean = ActiveWorkbook
    If clean Is Nothing Then Err.Raise vbObjectError + 4, "MakeTempXlsx", _
        "Sheets.Copy produced no new workbook"
    If StrComp(clean.Name, wb.Name, vbTextCompare) = 0 Then Err.Raise vbObjectError + 4, _
        "MakeTempXlsx", "Sheets.Copy did not create a separate workbook"

    MakeWorkbookOpenable clean
    clean.SaveAs fileName:=outXlsx, FileFormat:=51
    clean.Close SaveChanges:=False
    wb.Close SaveChanges:=False

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedScreenUpdating
    MakeTempXlsx = outXlsx
    Exit Function

Fail:
    ' Audit M-l: never leave a hidden workbook open holding a TEMP-file lock.
    Dim eNum As Long, eMsg As String
    eNum = Err.Number: eMsg = Err.Description
    On Error Resume Next
    If Not clean Is Nothing Then clean.Close SaveChanges:=False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedScreenUpdating
    On Error GoTo 0
    Err.Raise eNum, "MakeTempXlsx", eMsg
End Function
```

- [ ] **Step 5: M-l for `TrimSheetsInWorkbook`.** Wrap its body so the staging workbook always closes on error: after `Set wb = Application.Workbooks.Open(...)` (line 553), change error discipline by inserting `On Error GoTo TrimFail` right after the Open (before the keep-set build), and before the final `If deleteFailed Then` block's `Err.Raise`, add the handler at the end of the sub:

```vba
    Exit Sub
TrimFail:
    Dim eNum As Long, eMsg As String
    eNum = Err.Number: eMsg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedScreenUpdating
    On Error GoTo 0
    Err.Raise eNum, "TrimSheetsInWorkbook", eMsg
```

(The existing explicit `wb.Close`/restore/`Err.Raise` paths inside the sub stay as they are — the handler only catches what they don't.)

- [ ] **Step 6: M-b — validate to the trim bound.** Add the shared helper and rewire both users:

```vba
Private Function LastAfterWeightRow(ws As Worksheet, startCol As Long) As Long
    ' The same bound TrimSampleBlockTails ships: the last populated After-Weight
    ' cell. Validation must cover everything that uploads (audit M-b).
    Dim usedLastRow As Long
    usedLastRow = ws.UsedRange.row + ws.UsedRange.Rows.Count - 1
    LastAfterWeightRow = FIRST_DATA_ROW - 1
    If usedLastRow < FIRST_DATA_ROW Then Exit Function
    Dim found As Range
    On Error Resume Next
    Set found = ws.Range(ws.Cells(FIRST_DATA_ROW, startCol + 2), _
                         ws.Cells(usedLastRow, startCol + 2)) _
                  .Find(What:="*", LookIn:=xlValues, _
                        SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    On Error GoTo 0
    If Not found Is Nothing Then LastAfterWeightRow = found.row
End Function
```

In `ValidateSamplePuffs` replace the loop header + first-blank exit (lines 1242–1243):

```vba
    Dim lastRow As Long
    lastRow = LastAfterWeightRow(ws, startCol)
    For row = FIRST_DATA_ROW To lastRow
        If Not IsBlankRow(ws, row, startCol) Then
```

…indent the existing checks inside this `If`, and close it with `End If` before `Next row` (blank rows inside a block are now skipped, not treated as end-of-data). In `TrimSampleBlockTails`, replace the per-block `Find` logic (lines 665–677) with `lastAfterWeightRow = LastAfterWeightRow(ws, startCol)` (delete the local `found` usage for that computation).

- [ ] **Step 7: Lockbox (owner rule 3).** Make `ResetSheetToBlankWithReview` a `Private Function ... As Boolean` (returns `True` only on the success path: add `ResetSheetToBlankWithReview = True` right before its success-path `Exit Sub`→`Exit Function`; convert `Exit Sub`/`End Sub` accordingly — skip/fail paths return False). Then replace `ResetLiveWorkbookAfterUpload` (lines 727–754) with:

```vba
Private Function ResetLiveWorkbookAfterUpload(keep As Object) As Collection
    ' Revert each uploaded canonical sheet to its snapshot (Review copy kept),
    ' restore any canonical sheet the tester deleted or renamed (lockbox), then
    ' reset the selection and visibility. Returns the names that could NOT be
    ' reset so the receipt can say so honestly (audit H4c).
    Dim skipped As New Collection
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Cleanup

    Dim arr As Variant, i As Long, sheetName As String
    arr = CanonicalDataSheets()
    For i = LBound(arr) To UBound(arr)
        sheetName = CStr(arr(i))
        If keep.Exists(sheetName) And SheetExists(sheetName) Then
            If Not ResetSheetToBlankWithReview(ThisWorkbook, sheetName, i) Then
                skipped.Add sheetName
            End If
        End If
    Next

    ' Lockbox reconcile: the workbook always ends canonical, whatever the tester
    ' deleted/renamed/copied during the session.
    For i = LBound(arr) To UBound(arr)
        If Not SheetExists(CStr(arr(i))) Then EnsureCanonicalSheet i
    Next

    ResetSelectionToDefault
    ApplySheetVisibility

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Set ResetLiveWorkbookAfterUpload = skipped
End Function

Private Function EnsureCanonicalSheet(ByVal idx As Long) As Boolean
    ' Lockbox restore: if canonical sheet #idx is missing, stamp a fresh copy
    ' from its very-hidden _Template_NN snapshot. Never touches existing sheets.
    Dim arr As Variant
    arr = CanonicalDataSheets()
    If idx < LBound(arr) Or idx > UBound(arr) Then Exit Function
    Dim sheetName As String
    sheetName = CStr(arr(idx))
    If SheetExists(sheetName) Then
        EnsureCanonicalSheet = True
        Exit Function
    End If
    Dim tplName As String
    tplName = TemplateSheetName(idx)
    If Not WorkbookHasSheetIn(ThisWorkbook, tplName) Then
        StampLog "  '" & sheetName & "': missing and no snapshot (" & tplName & ") - cannot restore."
        Exit Function
    End If
    Dim savedAlerts As Boolean
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Fail
    Dim tpl As Worksheet
    Set tpl = ThisWorkbook.Worksheets(tplName)
    Dim savedVis As XlSheetVisibility
    savedVis = tpl.Visible
    tpl.Visible = xlSheetVisible
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        seen(w.Name) = True
    Next
    tpl.Copy After:=ThisWorkbook.Worksheets(UPLOAD_SHEET_NAME)
    Dim fresh As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If Not seen.Exists(w.Name) Then
            Set fresh = w
            Exit For
        End If
    Next
    tpl.Visible = savedVis
    If fresh Is Nothing Then GoTo Fail
    fresh.Name = sheetName
    fresh.Visible = xlSheetVisible
    Application.DisplayAlerts = savedAlerts
    StampLog "  '" & sheetName & "': restored fresh from " & tplName & "."
    EnsureCanonicalSheet = True
    Exit Function
Fail:
    Application.DisplayAlerts = savedAlerts
    StampLog "  '" & sheetName & "': restore from snapshot FAILED (" & Err.Description & ")."
End Function
```

- [ ] **Step 8: Lockbox on tick.** In `ApplySheetVisibility`, replace the canonical `For Each sheetName In CanonicalDataSheets()` loop (lines 149–165) with an indexed loop that restores a missing ticked sheet:

```vba
    Dim arr As Variant, i As Long
    arr = CanonicalDataSheets()
    Dim canonName As String, wantVisible As Boolean
    For i = LBound(arr) To UBound(arr)
        canonName = CStr(arr(i))
        wantVisible = False
        If selection.Exists(NormalizeSheetName(canonName)) Then
            wantVisible = selection(NormalizeSheetName(canonName))
        End If
        ' Lockbox: ticking a test whose sheet was deleted restores it fresh from
        ' its snapshot (owner rule: tester mutations never break the flow).
        If wantVisible And Not SheetExists(canonName) Then EnsureCanonicalSheet i
        If SheetExists(canonName) Then
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets(canonName)
            If wantVisible Then
                ws.Visible = xlSheetVisible
            Else
                ws.Visible = xlSheetHidden
            End If
        End If
    Next i
```

- [ ] **Step 9: M-c — persist picks.** Add the helper and call it from all three pickers (`Btn_PickSynologyFolder`, `Btn_PickLocalFolder`, `Btn_PickDataViewerExe`) after `RefreshPathLabels`:

```vba
Private Sub PersistSettings()
    ' Picked paths live on the hidden _Settings sheet; Workbook_Open marks the
    ' file Saved, so close + "Don't Save" silently discarded them (audit M-c).
    If ThisWorkbook.ReadOnly Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub
```

- [ ] **Step 10: Ribbon callback for Checkpoint.** Next to `Ribbon_UploadAll` (line 1446) add:

```vba
Public Sub Ribbon_UploadCheckpoint(control As IRibbonControl)
    Btn_UploadCheckpoint
End Sub
```

- [ ] **Step 11: Gate.** `check_sources.py` → ALL PASS (it cross-checks ribbon callbacks against handlers; the customUI button arrives in Task 5, but the handler existing first is fine — the check is XML→handler, not handler→XML; if it IS bidirectional and fails, do Task 5's customUI edit in this task instead and note it in the commit).

- [ ] **Step 12: Commit.** `git commit -am "feat(sidecar): v1.2 upload engine -- Upload Checkpoint, checkpoint-stream collisions, date stamp, delivery receipt, honest failure popups, ribbon-free .xlsx, synDest ingest, lockbox restore, persisted picks (T3)"`

---

### Task 4: ThisWorkbook.cls.txt — BeforeClose nudge (P6)

**Files:**
- Modify: `excel-sidecar/ThisWorkbook.cls.txt`

- [ ] **Step 1: Replace `Workbook_BeforeClose`** (lines 22–28) with:

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Restore Excel default behavior so the bindings don't leak into other
    ' workbooks opened in the same Excel session.
    On Error Resume Next
    Application.OnKey "+^."
    Application.OnKey "+^,"
    ' P6: one-button reminder, never a gate -- closing is always allowed.
    If DataViewerUpload.HasUnuploadedData() Then
        MsgBox "Reminder: this workbook holds test data that has not been " & _
               "uploaded this session." & vbCrLf & "Upload All / Upload Checkpoint " & _
               "(TPM Testing tab) delivers it to DataViewer - emailing this file " & _
               "does not.", vbInformation, "DataViewer Upload"
    End If
End Sub
```

- [ ] **Step 2: Gate + commit.** `check_sources.py` → ALL PASS. `git commit -am "feat(sidecar): BeforeClose unuploaded-data nudge (v1.2 T4, P6)"`

---

### Task 5: customUI14.xml + check_sources.py — Upload Checkpoint button, XML gate (audit M-g)

**Files:**
- Modify: `excel-sidecar/customUI14.xml`
- Modify: `excel-sidecar/check_sources.py`

- [ ] **Step 1: Add the button.** In `boxDvActions` (between `btnUploadAll` and `btnSpecifyName`):

```xml
            <button id="btnUploadCheckpoint" label="Upload Checkpoint" size="normal"
                    onAction="Ribbon_UploadCheckpoint"
                    screentip="Upload a mid-test checkpoint (no reset)"
                    supertip="Same as Upload All, but your sheets are NOT reset and no Review copies are made. Re-using the same file name overwrites the previous checkpoint; a new name creates a new file."/>
```

(The stacked box becomes 3 rows — at the 3-row group cap, OK.)

- [ ] **Step 2: XML well-formedness gate (M-g).** In `check_sources.py`, add a check (registered like the existing ones — read the file's harness first and follow its `ok()/fail()` pattern):

```python
def check_xml_wellformed():
    from xml.dom import minidom
    raw = open(os.path.join(REPO, "customUI14.xml"), "rb").read()
    try:
        dom = minidom.parseString(raw)
    except Exception as e:
        fail("customUI14.xml is not well-formed XML: %s" % e)
        return
    ok("customUI14.xml is well-formed XML")
    root = dom.documentElement
    if root.getAttribute("xmlns") == "http://schemas.microsoft.com/office/2009/07/customui":
        ok("customUI14.xml uses the customui/2009/07 namespace")
    else:
        fail("customUI14.xml namespace is wrong/missing")
```

- [ ] **Step 3: New-callback assertions.** Wherever check_sources lists expected ribbon callbacks/handlers explicitly (the v1.1 spot-list pattern — e.g. its checks for `Ribbon_SpecifyName`), add the same pair of assertions for `Ribbon_UploadCheckpoint` (wired in XML) and `Btn_UploadCheckpoint` (defined in DataViewerUpload.bas).

- [ ] **Step 4: Gate + commit.** `check_sources.py` → ALL PASS (now including the new checks). `git commit -am "feat(sidecar): Upload Checkpoint ribbon button + XML well-formedness gate (v1.2 T5, M-g)"`

---

### Task 6: Help popups — accurate v1.2 texts (audit H8 + new features)

**Files:**
- Modify: `excel-sidecar/DataViewerUpload.bas` (only `InstructionsText`, `SheetGuideText`, `SampleToolsText`)

- [ ] **Step 1: `SheetGuideText` (H8):** in the DROPDOWNS section, replace the puffs and Clog lines with:

```vba
    s = s & "  - Puffs (first column, rows 5+): type ANY number and the" & NL
    s = s & "      column auto-fills below it (row above + your number)." & NL
    s = s & "      Typing in the FIRST row seeds the whole column. Type" & NL
    s = s & "      'custom' to clear it and enter numbers by hand." & NL
    s = s & "  - Smell (the 'Smell' column): a 0-and-up number rating." & NL & NL
```

(the Clog dropdown line is deleted) and in the AUTO-CALCULATED list add Clog:

```vba
    s = s & "  Clog (from Draw Pressure: blank, 'Light Clog' or 'Heavy Clog');" & NL
```

- [ ] **Step 2: `SampleToolsText`:** replace the PUFFS AUTO-FILL paragraph with:

```vba
    s = s & "PUFFS AUTO-FILL" & NL
    s = s & "  Type any number into the puffs column and every row below" & NL
    s = s & "  becomes 'the row above + your number'. In the first data row" & NL
    s = s & "  it seeds the whole block. Type 'custom' to enter puff numbers" & NL
    s = s & "  by hand." & NL & NL
```

- [ ] **Step 3: `InstructionsText`:** replace the whole function body string with:

```vba
    Dim s As String
    s = "How to upload test data to DataViewer" & vbLf & vbLf
    s = s & "0.  If the ribbon buttons do nothing, enable macros: click" & vbLf
    s = s & "     'Enable Content' in the yellow bar (one time)." & vbLf & vbLf
    s = s & "1.  Tick the tests you're running. Each ticked test's sheet" & vbLf
    s = s & "     appears; untick to hide it again. (A deleted sheet comes" & vbLf
    s = s & "     back fresh when you re-tick its box.)" & vbLf & vbLf
    s = s & "2.  Set your three destinations once, from this ribbon:" & vbLf
    s = s & "     Pick Synology Folder, Pick Local Folder, and Pick" & vbLf
    s = s & "     DataViewer File. The paths show under 'Active Folders'." & vbLf & vbLf
    s = s & "3.  Upload All = final upload. Your data goes to the Synology" & vbLf
    s = s & "     + Local folders and into DataViewer; each uploaded sheet" & vbLf
    s = s & "     is reset and a '... - Review' copy is kept." & vbLf & vbLf
    s = s & "4.  Upload Checkpoint = mid-test save. Same delivery, but your" & vbLf
    s = s & "     sheets are NOT reset - keep testing and upload again." & vbLf
    s = s & "     Re-using the same file name overwrites the checkpoint;" & vbLf
    s = s & "     a new name creates a new file." & vbLf & vbLf
    s = s & "5.  Delete All Review Sheets clears review copies when done." & vbLf & vbLf
    s = s & "IMPORTANT: never email or share this workbook - it is your" & vbLf
    s = s & "reusable template. After an upload the data is ALREADY" & vbLf
    s = s & "delivered; send people the uploaded .xlsx copy instead" & vbLf
    s = s & "(the receipt popup shows exactly where it is)." & vbLf & vbLf
    s = s & "Tip:  Test SOP's is always included in every upload." & vbLf
    s = s & "Tip:  'Specify Test Name' labels this FILE for parallel" & vbLf
    s = s & "      projects; it does not affect uploaded data names."
    InstructionsText = s
```

- [ ] **Step 4: Gate + commit.** `check_sources.py` → ALL PASS. `git commit -am "docs(sidecar): v1.2 in-workbook help -- auto-Clog, any-number puffs, checkpoint, never-email rule (T6, H8)"`

---

### Task 7: build_clean_template.py — guards, banner, DV_LastUpload, Clog ISNUMBER (audit H6, M-d, M-k, F12)

**Files:**
- Modify: `excel-sidecar/build_clean_template.py`

- [ ] **Step 1: H6 guards + timestamped backups.** Add `import time` to the imports. In `main()`, add `ap.add_argument("--force", action="store_true", help="overwrite an existing --out")` and pass `args.force` to `build(...)`. Change `build(source, out)` signature to `build(source, out, force=False)` and replace lines 304–310 with:

```python
    assert_not_mip_ciphertext(source)
    if os.path.exists(out):
        if os.path.samefile(source, out):
            raise SystemExit("FATAL: --out is the same file as --source; refusing (audit H6).")
        if not force:
            raise SystemExit("FATAL: --out already exists: %s  (pass --force to overwrite; "
                             "a timestamped backup of it will be kept)" % out)
        out_bak = out + time.strftime(".bak-%Y%m%d_%H%M%S")
        shutil.copyfile(out, out_bak)
        print("Backed up existing --out ->", out_bak)
    backup = source + time.strftime(".bak-%Y%m%d_%H%M%S")
    shutil.copyfile(source, backup)
    print("Backed up source ->", backup)
    if os.path.exists(out):
        os.remove(out)
    shutil.copyfile(source, out)          # work on a copy; never touch the source
    print("Working copy ->", out)
```

- [ ] **Step 2: M-k MIP sniff.** Add near the top (after the constants):

```python
def assert_not_mip_ciphertext(path):
    """The deployed .xlsm is MIP-encrypted at rest; only the allowlisted Python
    reads plaintext. Fail loudly instead of propagating ciphertext (audit M-k)."""
    with open(path, "rb") as f:
        head = f.read(16)
    if head.startswith(b"%TSD-Header-"):
        raise SystemExit("FATAL: %s reads as MIP ciphertext under this interpreter.\n"
                         "Run with the allowlisted Python: "
                         "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" % path)
    if not head.startswith(b"PK\x03\x04"):
        raise SystemExit("FATAL: %s does not look like an OOXML (.xlsm/.xlsx) package." % path)
```

- [ ] **Step 3: DV_LastUpload (B8).** In `NAMED` add `"DV_LastUpload": "'_Settings'!$B$8",` and append `"Last uploaded file"` to `SETTINGS_LABELS`. In `_carry_over_values`, extend the loop tuple to `("DV_FileName", "DV_SynologyPath", "DV_LocalPath", "DV_DataViewerExe", "DV_OrigFileName", "DV_LastUpload")`. In `_build_settings_sheet`, after the row-6 write add:

```python
    st.Cells(7, 2).Value = carried.get("DV_OrigFileName", "") or ""
    st.Cells(8, 2).Value = carried.get("DV_LastUpload", "") or ""
```

- [ ] **Step 4: M-d Clog formula.** Replace `clog_formula` with:

```python
def clog_formula(draw_col_letter, row):
    """Clog formula keyed off the same-row Draw Pressure cell. ISNUMBER guard:
    text like 'n/a' sorts above all numbers in Excel and would otherwise
    classify as Heavy Clog (audit M-d)."""
    d = "%s%d" % (draw_col_letter, row)
    return ('=IF(NOT(ISNUMBER(%s)),"",IF(%s>=15,"Heavy Clog",IF(%s>5,"Light Clog","")))'
            % (d, d, d))
```

- [ ] **Step 5: P5 banner.** Add module constants after the TS_* block:

```python
TS_BANNER_FIRST = TS_LAST_ROW + 2
TS_BANNER_LINES = [
    "Macros required: if the ribbon buttons do nothing, click 'Enable Content' in the yellow bar.",
    "This is your reusable WORKING TEMPLATE - never email or upload this file.",
    "Upload All / Upload Checkpoint deliver your data to Synology + DataViewer automatically.",
]
```

In `_relay_test_selection`, after the 13 selection-row writes (the `for i, (name, checked) ...` loop), add:

```python
    # P5: the only guidance channel that works with macros DISABLED. Test
    # Selection never enters distributed copies, so it can be loud.
    for i, txt in enumerate(TS_BANNER_LINES):
        ws.Cells(TS_BANNER_FIRST + i, cc).Value = txt
```

and inside the guarded styling `try:` block (before the gridline-hiding part) add:

```python
        DARKRED = 0x0000C0
        for i in range(len(TS_BANNER_LINES)):
            r = TS_BANNER_FIRST + i
            try:
                ws.Range(ws.Cells(r, cc), ws.Cells(r, lc + 6)).Merge()
            except Exception as _e:
                print("WARNING: banner merge skipped: %s" % _e)
            cell = ws.Cells(r, cc)
            cell.Font.Size = 10
            cell.Font.Bold = (i == 1)
            cell.Font.Color = DARKRED if i == 1 else GREY
            cell.HorizontalAlignment = XL_LEFT
```

(Note: the existing clears below the table use `lr + 200` and width `52`, which covers the banner zone — idempotent on rebuild.)

- [ ] **Step 6: F12 — verify the surgered zip BEFORE swapping it in.** In `inject_customui`, move the post-`os.replace` assertion block so it runs against `tmp` *before* `os.replace(tmp, target_xlsm)` (open `zipfile.ZipFile(tmp)` instead of `target_xlsm`); on assertion failure delete `tmp` and raise, leaving the target untouched.

- [ ] **Step 7: Gates + commit.** Run `check_sources.py` (ALL PASS) and `"C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" excel-sidecar/test_build_helpers.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.1.xlsm"` (ALL PASS — if it asserts on `NAMED`/`SETTINGS_LABELS` shapes, update those expectations to include `DV_LastUpload`). `git commit -am "feat(sidecar): build guards (samefile/--force/timestamped baks), MIP sniff, on-sheet banner, DV_LastUpload, ISNUMBER Clog, pre-swap zip asserts (v1.2 T7)"`

---

### Task 8: verify_sidecar.py — close the blind spots (audit M-f) + signature gate

**Files:**
- Modify: `excel-sidecar/verify_sidecar.py`

Read the file first; integrate these checks following its existing `[OK]/[FAIL]` reporting pattern, all feeding the final exit code:

- [ ] **Step 1: Structure presence.** Assert every member of: 12 canonical sheets (import `CANON` from build_clean_template like the geometry constants are imported, with the same fallback pattern), `_Template_00`–`_Template_11`, `_Template_Master`, `Test SOP's`, `Test Selection`, `_Settings` exists in the workbook (openpyxl sheetnames).
- [ ] **Step 2: Visibility states.** Using openpyxl `sheet_state`: `Test Selection`, `Test SOP's`, `Lifetime Test` = `visible`; every other canonical = `hidden`; every `_Template_*` and `_Settings` = `veryHidden`.
- [ ] **Step 3: featurePropertyBag (the checkbox carrier).** Zip-level: `any(n.startswith("xl/featurePropertyBag") for n in z.namelist())` must be True — this is the part whose loss degrades checkboxes to TRUE/FALSE text (audit M-f).
- [ ] **Step 4: Literal anchor.** Independently of the imported geometry, assert the exact string `"'Test Selection'!$B$4:$C$16"` == DV_TestSelection's RefersTo (not substring — audit M-f's `$B$40` hole). Keep the derived-geometry check too.
- [ ] **Step 5: DV_LastUpload** joins the named-range checks (`'_Settings'!$B$8`).
- [ ] **Step 6: Banner present.** Read `Test Selection` cell at `(TS_BANNER_FIRST + 1, TS_CHECK_COL)` (the "never email" line) and assert it contains `"never email"` (case-insensitive).
- [ ] **Step 7: Clog formula expectation.** Update the G5 formula check to expect the new `ISNUMBER` form from Task 7.
- [ ] **Step 8: `--require-signature` flag.** New argparse flag; when set, assert the package contains a part matching `xl/vbaProjectSignature` (any of `.bin`/`Agile`/`V3` variants — match prefix). Unsigned → FAIL. (The runbook's final gate uses this; the pre-sign build check runs without it.)
- [ ] **Step 9: MIP sniff** — call the same `assert_not_mip_ciphertext` logic on the `--file` argument before opening (duplicate the 10-line helper; verify_sidecar must stay runnable standalone).
- [ ] **Step 10: Gate + commit.** Run verify against the live v1.1 template: `python.exe excel-sidecar/verify_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.1.xlsm"` — expect: TestingTools DIFFERS (drift now ported — the deployed `pick.Validation.Delete`/IsNumeric variant vs our richer v1.2; this is EXPECTED until the v1.2 build), banner check FAILS (v1.1 file has no banner), DV_LastUpload FAILS — confirm the failures are exactly these expected ones and the NEW checks otherwise run clean (presence/visibility/featurePropertyBag/anchor must PASS against v1.1). `git commit -am "feat(sidecar): verify gains presence/visibility/featurePropertyBag/literal-anchor/banner/signature checks (v1.2 T8, M-f)"`

---

### Task 9: Signing kit (VBA code-signing, owner decision Q3)

**Files:**
- Create: `excel-sidecar/signing/make_cert.ps1`
- Create: `excel-sidecar/signing/tester-setup.ps1`
- Create: `excel-sidecar/signing/README.md`

(Write all three via the Python delete-and-rewrite pattern — they are NEW files. UTF-8, LF.)

- [ ] **Step 1: `make_cert.ps1`** (owner runs ONCE; private key stays in their user cert store, never in the repo):

```powershell
# Creates the long-lived code-signing certificate for the Automated Testing
# Template VBA project and exports the PUBLIC half next to this script.
# Run ONCE as the template owner. The private key stays in YOUR user store
# (certmgr.msc -> Personal). Never commit a .pfx.
$ErrorActionPreference = "Stop"
$existing = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { $_.Subject -eq "CN=SDR DataViewer Templates" }
if ($existing) {
    Write-Host "Certificate already exists (thumbprint $($existing[0].Thumbprint)); re-exporting public key."
    $cert = $existing[0]
} else {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert `
        -Subject "CN=SDR DataViewer Templates" `
        -CertStoreLocation Cert:\CurrentUser\My `
        -KeyUsage DigitalSignature `
        -NotAfter (Get-Date).AddYears(10)
    Write-Host "Created code-signing certificate, thumbprint $($cert.Thumbprint)"
}
$out = Join-Path $PSScriptRoot "DataViewerTemplates.cer"
Export-Certificate -Cert $cert -FilePath $out | Out-Null
Write-Host "Public certificate exported -> $out"
Write-Host "Commit DataViewerTemplates.cer; testers trust it via tester-setup.ps1."
```

- [ ] **Step 2: `tester-setup.ps1`** (each tester runs ONCE; no admin):

```powershell
# One-time tester setup for the Automated Testing Template: trusts the
# template's code-signing certificate so signed macros run with NO prompts,
# even on copies that arrive with Mark of the Web (email/Teams/downloads).
# No admin rights needed (current-user stores only).
$ErrorActionPreference = "Stop"
$cer = Join-Path $PSScriptRoot "DataViewerTemplates.cer"
if (-not (Test-Path $cer)) {
    Write-Error "DataViewerTemplates.cer not found next to this script."
}
certutil -user -addstore Root $cer | Out-Null
certutil -user -addstore TrustedPublisher $cer | Out-Null
Write-Host ""
Write-Host "Done. The DataViewer testing template's macros are now trusted on this"
Write-Host "account. Close and reopen the template; there should be no macro prompts."
```

- [ ] **Step 3: `signing/README.md`** documenting: WHY (zero prompts on every channel incl. MOTW'd email copies; trusted-publisher check precedes the MOTW block); OWNER flow (run `make_cert.ps1` once → after EVERY build, open the built .xlsm → Alt+F11 → Tools → Digital Signature → Choose → "SDR DataViewer Templates" → OK → save → run `verify_sidecar.py --file <out> --require-signature`; note: ANY edit to the VBA inside the workbook strips the signature — which doubles as a drift alarm); TESTER flow (right-click `tester-setup.ps1` → Run with PowerShell, once per Windows account); SECURITY notes (never commit a .pfx; anything signed with this cert auto-runs for testers — guard the owner account; signature does not cover non-VBA workbook content).

- [ ] **Step 4: Commit.** `git add excel-sidecar/signing && git commit -m "feat(sidecar): VBA code-signing kit -- cert script, tester one-time setup, signing runbook (v1.2 T9)"`

---

### Task 10: Docs + retire install_sidecar.py (audit H7, M-j)

**Files:**
- Delete: `excel-sidecar/install_sidecar.py` (git rm — audit H7: attaches to the user's running Excel and `Quit()`s it with alerts off; superseded by build_clean_template.py)
- Modify: `excel-sidecar/README.md`
- Modify: `excel-sidecar/RUNBOOK-migrate-existing-template.md`

- [ ] **Step 1: `git rm excel-sidecar/install_sidecar.py`**
- [ ] **Step 2: README.md** — update: (a) geometry: `DV_TestSelection = 'Test Selection'!$B$4:$C$16`, title B2/hint B3/table B4:C16 (M-j); (b) "How the workbook works" gains Upload Checkpoint (step 3.5: same delivery, no reset, same-name overwrites = checkpoint stream), the lockbox behavior (deleted sheet restored on re-tick; Upload All ends canonical), row-5 puff seeding + any-number steps, the banner, `DV_LastUpload` in the named-range list, and the distributed copies being "macro-free AND ribbon-free `.xlsx`"; (c) the "Known issues" entry about ResetEquations interval-20 is RESOLVED — replace with a note that Reset no longer touches puffs/before-weight columns; (d) signing: one paragraph pointing at `signing/README.md`; (e) MakeTempXlsx known-issue paragraph: update — it now copies sheets into a fresh workbook.
- [ ] **Step 3: RUNBOOK** — (a) fix acceptance item 2 + Option-B references from `A3:A15` to `B4:B16` (M-j); (b) add the sign step between build and verify, and `--require-signature` on the final verify; (c) extend the operator acceptance checklist with v1.2 items:
  - Upload Checkpoint: delivers + does NOT reset; immediate second checkpoint with the same name overwrites silently; a DIFFERENT name that collides with a foreign file prompts.
  - Receipt popup shows paths + opens the Synology folder; "never email" wording present.
  - Receiver test: open the distributed `.xlsx` on a machine/profile without the template → NO "TPM Testing" tab, NO macro popups (H9).
  - Delete a canonical sheet → re-tick its box → fresh sheet restored (lockbox).
  - Rename a canonical sheet, Upload All → warning lists the orphan; after upload the canonical sheet exists again.
  - Type `7` (a non-list number) into a puffs cell row 7 → fills below with +7; type `25` into row 5 → row 5 = 25, rows below = prev+25 (and a `- Review` sheet is NOT affected by typing in its puffs column).
  - Draw Pressure = `n/a` → Clog stays EMPTY (not Heavy Clog).
  - Macros disabled (open a copy without enabling) → the Test Selection banner is visible and legible.
  - Signed build on a tester-setup machine → zero macro prompts; on a non-setup machine → standard prompt (not an error).
  - Close with un-uploaded data → one-button reminder appears; close after an upload → no reminder.
- [ ] **Step 4: Gate + commit.** `check_sources.py` → ALL PASS. `git commit -am "docs(sidecar): v1.2 README/RUNBOOK -- checkpoint, lockbox, signing, B4 geometry; retire install_sidecar.py (T10, H7/M-j)"`

---

### Task 11: Final verification + build (orchestrator, this machine)

- [ ] **Step 1: Full headless gates** in the worktree: `check_sources.py` (ALL PASS) + `test_build_helpers.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.1.xlsm"` (ALL PASS).
- [ ] **Step 2: Code review** (superpowers:requesting-code-review) of the whole branch diff vs `feature/v2.4.0-bugfix-batch`; fix anything Critical/Important.
- [ ] **Step 3: Build v1.2** (Excel must be closed):

```
python excel-sidecar/build_clean_template.py --source "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.1.xlsm" --out "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.2.xlsm"
python excel-sidecar/verify_sidecar.py --file "C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.2.xlsm"
```

Expect: ALL modules MATCH (the C4 drift disappears — repo now carries the ported behavior), all structure checks OK, banner OK, DV_LastUpload OK. Signature check NOT run yet (unsigned until the owner signs).
- [ ] **Step 4: USER steps (hand off — never automated):** run `signing/make_cert.ps1`, sign the project in the VBE, save, re-run verify with `--require-signature`, then the RUNBOOK acceptance checklist in Excel, then rename/distribute + have testers run `tester-setup.ps1`. Branch merges only after acceptance passes (workflow_branch_to_main).

---

## Self-review notes

- Spec coverage: C1–C4 (T1/T3/T2/T1), H1–H9 (T1×3, T3×4, T7, T10, T6), M-a–M-l (T3×5, T7×2, T8, T1, T10), P2–P7 (T2, T3×3, T7, T4), owner features F1 checkpoint (T3/T5), F2 row-5 (T1), F3 lockbox (T3), F4 signing (T9), F5 ribbon-free (T3). Deferred by owner: DataViewer-side rejection (P1). Not in scope: UniqueReviewName line-813 theoretical fallback (audit Low #14, requires 98 review copies).
- Type consistency: `ResetSheetToBlankWithReview` becomes a Function (T3 step 7) — its only caller is rewritten in the same task. `EnsureCanonicalSheet` is used by both T3 step 7 and step 8 — defined once in step 7. `HasUnuploadedData` defined T3, consumed T4 — T3 runs first.
- VBA `Date`/`Now` usage is normal VBA (the Workflow-tool restriction does not apply to VBA).

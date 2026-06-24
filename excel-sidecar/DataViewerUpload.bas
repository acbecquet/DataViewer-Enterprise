Attribute VB_Name = "DataViewerUpload"
Option Explicit

' ============================================================================
' DataViewer Enterprise - upload sidecar module (template v1.2).
'
' Workflow:
'   1. Workbook opens -> only "Lifetime Test" + "Test Selection" + "Test
'      SOP's" are visible. Everything else (incl. other canonical data sheets,
'      _Macro_Install, _Template_Master, _Template_NN, _Settings) is hidden.
'   2. On the "Test Selection" sheet, the user toggles TRUE/FALSE next to
'      each canonical test name in the DV_TestSelection range. Toggling TRUE
'      unhides the matching test sheet; FALSE re-hides it.
'   3. Tech fills out the selected test sheets. Save/close/reopen freely -
'      visibility is driven entirely by the DV_TestSelection cells.
'   4. Click "Upload All" (or "Upload Checkpoint"). The macro:
'        a. runs the checklist on selected sheets,
'        b. stages a trimmed .xlsm (selected + populated + Test SOP's),
'        c. materializes ONE clean .xlsx by copying the staged sheets into a
'           FRESH workbook - macro-free AND ribbon-free, so receivers never
'           see dead macro buttons or "Cannot run the macro" popups,
'        d. copies it to <DV_SynologyPath>\<DV_FileName>.xlsx and
'           <DV_LocalPath>\<DV_FileName>.xlsx,
'        e. launches DataViewer on the SYNOLOGY copy (-> Postgres ingest).
'   5. Upload All then resets each uploaded sheet from its internal blank
'      snapshot (_Template_NN; a "- Review" copy is kept), restores missing
'      canonical sheets (lockbox), resets DV_TestSelection to default and
'      re-hides everything else, so the workbook is fresh for the next run.
'      Upload Checkpoint is the same delivery WITHOUT the reset: re-using the
'      same file name overwrites that checkpoint stream's own copies; a new
'      name starts a new file.
'
' Named ranges that must exist on the workbook:
'   DV_FileName         single-cell text  - base filename (no extension)
'   DV_SynologyPath     single-cell text  - destination folder #1
'   DV_LocalPath        single-cell text  - destination folder #2
'   DV_Status           single-cell text  - macro status output
'   DV_Log              single-cell text  - macro log output (multi-line)
'   DV_LastUpload       single-cell text  - base name of the last delivered
'                                            upload (checkpoint-stream identity)
'   DV_OrigFileName     single-cell text  - original on-disk file name; Specify
'                                            Test Name / pre-upload revert use it
'   DV_TestSelection    2-col range       - col 1 TRUE/FALSE, col 2 sheet name
'                                            (one row per canonical test sheet)
'
' Optional named range:
'   DV_DataViewerExe    single-cell text  - override DataViewer.exe path
'
' ThisWorkbook class module wire-up (REQUIRED):
'
'   Private Sub Workbook_Open()
'       DataViewerUpload.ApplySheetVisibility
'   End Sub
'
'   Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'       DataViewerUpload.OnWorkbookSheetChange Sh, Target
'   End Sub
' ============================================================================

' Default install location. Override by adding a DV_DataViewerExe named range.
Private Const DEFAULT_DATAVIEWER_EXE As String = _
    "C:\Program Files\DataViewer Enterprise\DataViewer.exe"

Private Const COLS_PER_SAMPLE As Long = 12
Private Const MAX_SAMPLES_PER_SHEET As Long = 24   ' up to 24 sample blocks per sheet
Private Const FIRST_DATA_ROW As Long = 5   ' rows 1-3 metadata, row 4 headers
Private Const TPM_MAX_PLAUSIBLE As Double = 50#

Private Const SELECTION_RANGE_NAME As String = "DV_TestSelection"
Private Const UPLOAD_SHEET_NAME As String = "Test Selection"
Private Const DEFAULT_SELECTED_SHEET As String = "Lifetime Test"
Private Const SOPS_SHEET_NAME As String = "Test SOP's"

' Captured at ribbon load so the path rows can be refreshed after a Pick.
Public gRibbon As IRibbonUI

' True once this session has successfully delivered data (Upload All or
' Checkpoint, or copies-landed-but-exe-missing). Drives the BeforeClose nudge.
Private gUploadedThisSession As Boolean

' ----------------------------------------------------------------------------
' Canonical data-sheet list. Order matters - this is the order Btn_UploadAll,
' the checklist, and the selection rendering all iterate in. Keep aligned
' with the DV_TestSelection range layout in the workbook.
' ----------------------------------------------------------------------------
Private Function CanonicalDataSheets() As Variant
    CanonicalDataSheets = Array( _
        "Lifetime Test", _
        "User Test Simulation", _
        "Long Puff Lifetime Test", _
        "Rapid Puff Lifetime Test", _
        "Intense Test", _
        "Big Headspace Serial Test", _
        "Negative Pressure Test", _
        "Temperature Cycling Test #2", _
        "Viscosity Compatibility", _
        "Various Oil Compatibility", _
        "Custom Test Template", _
        "Temperature Cycling Test #1")
End Function

' ============================================================================
' Sheet visibility driven by DV_TestSelection
' ============================================================================

Public Sub OnWorkbookSheetChange(Sh As Object, Target As Range)
    ' Forwarded from ThisWorkbook.Workbook_SheetChange. We only care about
    ' edits inside DV_TestSelection on the DataViewer Upload sheet.
    On Error GoTo Bail
    If Sh Is Nothing Then Exit Sub
    If StrComp(Sh.Name, UPLOAD_SHEET_NAME, vbTextCompare) <> 0 Then Exit Sub

    Dim selRange As Range
    Set selRange = SelectionRange()
    If selRange Is Nothing Then Exit Sub

    ' Only react if the change touches the first column of the selection
    ' range (the TRUE/FALSE column).
    Dim hit As Range
    Set hit = Application.Intersect(Target, selRange.Columns(1))
    If hit Is Nothing Then Exit Sub

    ApplySheetVisibility
    Exit Sub
Bail:
    ' Swallow - never let an event handler break user typing.
End Sub

Public Sub ApplySheetVisibility()
    ' Walks DV_TestSelection and sets each canonical sheet visible/hidden
    ' accordingly. Also enforces:
    '   DataViewer Upload + Test SOP's = always xlSheetVisible
    '   _Macro_Install + _Template_* = always xlSheetVeryHidden
    ' A trimmed/distributed copy has no "DataViewer Upload" sheet and hence no
    ' DV_TestSelection range. With no selection, do nothing - otherwise an empty
    ' selection would hide every data sheet and the copy would open blank.
    If SelectionRange() Is Nothing Then Exit Sub

    Dim selection As Object
    Set selection = ReadSelection()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Cleanup

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

    ' Canonical data sheets: visibility from selection.
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

    ' Hidden template / utility sheets stay xlSheetVeryHidden.
    EnsureVeryHidden "_Macro_Install"
    EnsureVeryHidden "_Template_Master"
    EnsureVeryHidden "_Settings"
    Dim wsAll As Worksheet
    For Each wsAll In ThisWorkbook.Worksheets
        If Left$(wsAll.Name, 10) = "_Template_" Then wsAll.Visible = xlSheetVeryHidden
    Next

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub EnsureVisible(sheetName As String, vis As XlSheetVisibility)
    If Not SheetExists(sheetName) Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Visible = vis
    On Error GoTo 0
End Sub

Private Sub EnsureVeryHidden(sheetName As String)
    EnsureVisible sheetName, xlSheetVeryHidden
End Sub

Private Function ReadSelection() As Object
    ' Returns Dictionary[normalized sheet name -> Boolean].
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim selRange As Range
    Set selRange = SelectionRange()
    If selRange Is Nothing Then
        Set ReadSelection = d
        Exit Function
    End If

    Dim row As Long
    For row = 1 To selRange.Rows.Count
        Dim flagCell As Range, nameCell As Range
        Set flagCell = selRange.Cells(row, 1)
        Set nameCell = selRange.Cells(row, 2)
        Dim sheetName As String
        sheetName = Trim$(SafeString(nameCell.value))
        If Len(sheetName) > 0 Then
            d(NormalizeSheetName(sheetName)) = ParseBool(flagCell.value)
        End If
    Next row

    Set ReadSelection = d
End Function

Private Function SelectionRange() As Range
    On Error Resume Next
    Set SelectionRange = ThisWorkbook.Names(SELECTION_RANGE_NAME).RefersToRange
    On Error GoTo 0
End Function

Private Function ParseBool(v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If TypeName(v) = "Boolean" Then
        ParseBool = CBool(v)
    ElseIf IsNumeric(v) Then
        ParseBool = (CDbl(v) <> 0)
    Else
        Dim s As String
        s = UCase$(Trim$(SafeString(v)))
        ParseBool = (s = "TRUE" Or s = "YES" Or s = "Y" Or s = "1" Or s = "X")
    End If
End Function

Public Sub ResetSelectionToDefault()
    ' Sets all selection rows to FALSE except DEFAULT_SELECTED_SHEET.
    Dim selRange As Range
    Set selRange = SelectionRange()
    If selRange Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo Cleanup

    Dim row As Long
    For row = 1 To selRange.Rows.Count
        Dim sheetName As String
        sheetName = Trim$(SafeString(selRange.Cells(row, 2).value))
        If IsDefaultVisibleSheet(sheetName) Then
            selRange.Cells(row, 1).value = True
        Else
            selRange.Cells(row, 1).value = False
        End If
    Next row

Cleanup:
    Application.EnableEvents = True
End Sub

Private Function IsDefaultVisibleSheet(sheetName As String) As Boolean
    ' Sheets that should be TRUE after a post-upload reset: Lifetime + Test SOP's.
    Dim n As String
    n = NormalizeSheetName(sheetName)
    IsDefaultVisibleSheet = _
        (StrComp(n, NormalizeSheetName(DEFAULT_SELECTED_SHEET), vbTextCompare) = 0) Or _
        (StrComp(n, NormalizeSheetName(SOPS_SHEET_NAME), vbTextCompare) = 0)
End Function

' ============================================================================
' Public button handlers
' ============================================================================

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

Private Function HasDateToken(ByVal s As String) As Boolean
    ' True if the name already carries a recognizable date (yyyy-mm-dd or
    ' d-m-yyyy style, with -, / or . separators).
    ' Heuristic: digit runs like lot codes (12-34-5678) can false-positive,
    ' which merely suppresses the auto-stamp - harmless.
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
          "  - opening in DataViewer now (it saves to the shared database)." & vbLf & vbLf & _
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
    Dim warningsConfirmed As Boolean   ' set once the tester OKs the warnings dialog
    On Error GoTo Failed
    ClearLog
    SetNamed "DV_Status", "Starting " & actionName & "..."
    StampLog actionName & " started"

    ' Bug 1 (owner: "Checkpoint keeps name; All reverts"): the on-disk file is
    ' NOT renamed here. A checkpoint is mid-test -- the workbook must keep its
    ' descriptive working name so the tester carries on. Only Upload All reverts
    ' to the clean template name (silently), and only AFTER it has reset the data
    ' (see section 5) so the file genuinely ends up a fresh template.

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
    ' Separate guard: VBA's Like char-list can't express ]. Brackets are legal
    ' on disk but Excel's SaveAs/Open reject them later with a cryptic 1004.
    If InStr(fName, "[") > 0 Or InStr(fName, "]") > 0 Then
        SetNamed "DV_Status", "Cancelled (invalid file name)"
        MsgBox "The file name can't contain square brackets [ ] (Excel rejects them).", _
               vbExclamation, "Upload file name"
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
    If warns.Count > 0 And Not warningsConfirmed Then
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
        ' A later GoTo AskName re-runs the checklist but must not nag with the
        ' same already-confirmed warnings again.
        warningsConfirmed = True
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
    ' The shared copy is delivered from this moment on: record the stream name
    ' immediately so a retry after a later failure overwrites silently instead
    ' of tripping the cross-stream confirm (quality review #3).
    SetNamed "DV_LastUpload", baseName
    gUploadedThisSession = True
    StampLog "Copy -> " & locDest
    fso.CopyFile cleanXlsx, locDest, True
    ' Persist DV_LastUpload now -- a checkpoint run does no further save, so
    ' close + "Don't Save" would otherwise forget the stream name (review #2).
    PersistSettings
    ' From here on the data is already in the shared folders: every later error
    ' (temp cleanup, Shell, checkpoint branch, reset phase) must report through
    ' the delivered-state handler, never as a failed upload (review #6).
    On Error GoTo PostDeliveryFailed

    ' Staging artifacts are no longer needed: DataViewer ingests the Synology
    ' copy (audit H5), so the per-run TEMP dir can go immediately (audit M-l/11).
    On Error Resume Next
    fso.DeleteFolder Left$(cleanXlsx, InStrRev(cleanXlsx, "\") - 1), True
    On Error GoTo PostDeliveryFailed

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

    ' --- 5. Upload All only: reset the LIVE workbook (PostDeliveryFailed armed) ---
    StampLog "Resetting live workbook"
    Dim skipped As Collection
    Set skipped = ResetLiveWorkbookAfterUpload(keep)
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    ' Bug 1: now that the data is reset to a clean template, revert the on-disk
    ' file name to the original template name -- SILENTLY (overwriting the prior
    ' template is intended and safe: both are now blank), leaving exactly one
    ' file (the old "(do not send)" working copy is renamed away, no duplicate).
    RevertToOriginalName

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

PostDeliveryFailed:
    ' The data is ALREADY in the shared folders -- never report this as a failed
    ' upload, and never let the tester re-enter data (quality review #6).
    Application.DisplayAlerts = True
    StampLog "WARN: post-delivery step failed: " & Err.Description
    SetNamed "DV_Status", "OK (delivered; a follow-up step failed - see log)"
    MsgBox "Your data WAS delivered to the Synology and Local folders." & vbLf & vbLf & _
           "But a follow-up step failed: " & Err.Description & vbLf & vbLf & _
           "Do NOT re-enter your data. If DataViewer did not open, run " & actionName & _
           " again with the SAME file name - it overwrites the copies and nothing " & _
           "is duplicated.", vbExclamation, actionName
    Exit Sub

Failed:
    ' Pre-delivery errors only: nothing has reached the shared folders yet.
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    StampLog "ERROR " & Err.Number & ": " & Err.Description
    SetNamed "DV_Status", "Failed: " & Err.Description
    MsgBox actionName & " failed:" & vbLf & Err.Description & vbLf & vbLf & _
           "Nothing was lost. Fix the issue and run " & actionName & " again." & vbLf & _
           "Do NOT email this workbook as a workaround.", vbCritical, actionName
End Sub

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

' ============================================================================
' Keep-list construction (selection-driven, no session tracking)
' ============================================================================

Private Function BuildKeepList() As Object
    Dim keep As Object
    Set keep = CreateObject("Scripting.Dictionary")
    keep.CompareMode = vbTextCompare
    keep.Add SOPS_SHEET_NAME, True

    Dim selection As Object
    Set selection = ReadSelection()

    Dim sheetName As Variant
    For Each sheetName In CanonicalDataSheets()
        Dim norm As String
        norm = NormalizeSheetName(CStr(sheetName))
        If selection.Exists(norm) Then
            If selection(norm) And SheetExists(CStr(sheetName)) Then
                If SheetHasPopulatedSamples(ThisWorkbook.Worksheets(CStr(sheetName))) Then
                    If Not keep.Exists(CStr(sheetName)) Then
                        keep.Add CStr(sheetName), True
                    End If
                End If
            End If
        End If
    Next

    Set BuildKeepList = keep
End Function

Private Function SheetHasPopulatedSamples(ws As Worksheet) As Boolean
    Dim sampleIdx As Long, startCol As Long, sampleID As String
    For sampleIdx = 0 To MAX_SAMPLES_PER_SHEET - 1
        startCol = sampleIdx * COLS_PER_SAMPLE + 1
        sampleID = Trim$(SafeString(ws.Cells(1, startCol + 5).value))
        If IsRealSampleId(sampleID) Then
            SheetHasPopulatedSamples = True
            Exit Function
        End If
    Next
End Function

Private Function IsRealSampleId(ByVal sampleID As String) As Boolean
    ' Bug 3: a tester-entered sample ID, NOT a static template label. On the
    ' irregular sheets (Temperature Cycling, etc.) the row1/startCol+5 cell holds
    ' a fixed label like "Heater Technology:" or "Sample ID:" -- which made an
    ' EMPTY sheet read as "populated" and fire a false "has data" warning, and
    ' could pull empty sheets into the upload. Every such label ends with a colon;
    ' a real sample ID never does (":" is also a banned file-name character).
    If Len(sampleID) = 0 Then Exit Function
    If Right$(sampleID, 1) = ":" Then Exit Function
    IsRealSampleId = True
End Function

' ============================================================================
' Trim staged copy: drop unwanted sheets + trim sample-block tails
' ============================================================================

Private Sub TrimSheetsInWorkbook(filePath As String, keep As Object)
    ' Guard: never trim the active workbook.
    If StrComp(filePath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 2, "TrimSheetsInWorkbook", _
                  "Refusing to trim ThisWorkbook.FullName"
    End If

    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(fileName:=filePath, UpdateLinks:=0)
    ' Audit M-l: from here on, any unexpected error must close the staging
    ' workbook (else it stays open, hidden, holding a TEMP-file lock).
    On Error GoTo TrimFail

    ' Hide the staging workbook window so the user never sees the trimmed copy.
    On Error Resume Next
    Application.Windows(wb.Name).Visible = False
    On Error GoTo TrimFail

    ' Build a normalized keep-set so curly-vs-straight apostrophes etc. match.
    Dim keepNorm As Object
    Set keepNorm = CreateObject("Scripting.Dictionary")
    keepNorm.CompareMode = vbTextCompare
    Dim k As Variant
    For Each k In keep.Keys
        If Not keepNorm.Exists(NormalizeSheetName(CStr(k))) Then
            keepNorm.Add NormalizeSheetName(CStr(k)), True
        End If
    Next

    Dim victims As Collection
    Set victims = New Collection
    ' wb.Sheets, not wb.Worksheets: tester-created CHART sheets must not
    ' survive into distributed copies. Object - chart sheets aren't Worksheets.
    Dim sh As Object
    For Each sh In wb.Sheets
        If Not keepNorm.Exists(NormalizeSheetName(sh.Name)) Then
            victims.Add sh.Name
        End If
    Next

    If victims.Count >= wb.Sheets.Count Then
        wb.Close SaveChanges:=False
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = savedScreenUpdating
        Err.Raise vbObjectError + 1, "TrimSheetsInWorkbook", _
                  "Keep list matched no sheets in workbook; aborting trim"
    End If

    ' Delete victims with per-sheet logging.
    Dim victim As Variant
    Dim deleteFailed As Boolean
    deleteFailed = False
    For Each victim In victims
        On Error Resume Next
        wb.Sheets(CStr(victim)).Visible = xlSheetVisible
        wb.Sheets(CStr(victim)).Delete
        If Err.Number <> 0 Then
            StampLog "  Could not delete sheet '" & CStr(victim) & "': " & Err.Description
            deleteFailed = True
            Err.Clear
        End If
        On Error GoTo TrimFail
    Next

    ' Trim trailing rows per sample block on each surviving canonical sheet.
    Dim sheetName As Variant
    For Each sheetName In CanonicalDataSheets()
        If WorkbookHasSheetIn(wb, CStr(sheetName)) Then
            TrimSampleBlockTails wb.Worksheets(CStr(sheetName))
        End If
    Next

    ' Re-show the window + activate a data sheet so the distributed copy does
    ' not open to a blank gray screen (Excel persists the hidden-window state).
    MakeWorkbookOpenable wb
    wb.Save
    wb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedScreenUpdating

    If deleteFailed Then
        Err.Raise vbObjectError + 3, "TrimSheetsInWorkbook", _
                  "One or more sheets could not be deleted (see log)"
    End If
    Exit Sub
TrimFail:
    ' 'eNum' was a VBA reserved word (== 'Enum', case-insensitive) -> the whole
    ' Dim line was a Compile error, so this proc never compiled and every Upload
    ' died here (bug 2). Renamed to errNum.
    Dim errNum As Long, eMsg As String
    errNum = Err.Number: eMsg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedScreenUpdating
    On Error GoTo 0
    Err.Raise errNum, "TrimSheetsInWorkbook", eMsg
End Sub

Private Function WorkbookHasSheetIn(wb As Workbook, sheetName As String) As Boolean
    Dim Target As String
    Target = NormalizeSheetName(sheetName)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(NormalizeSheetName(ws.Name), Target, vbTextCompare) = 0 Then
            WorkbookHasSheetIn = True
            Exit Function
        End If
    Next
End Function

Private Function NormalizeSheetName(s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, ChrW(8216), "'")
    t = Replace(t, ChrW(8217), "'")
    t = Replace(t, ChrW(8220), """")
    t = Replace(t, ChrW(8221), """")
    NormalizeSheetName = t
End Function

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

Private Sub TrimSampleBlockTails(ws As Worksheet)
    ' For each sample block (12 columns wide), find the last row
    ' with a populated "After Weight" cell (col offset +2) and clear every
    ' cell in the block below that row.
    Dim usedLastRow As Long
    usedLastRow = ws.UsedRange.row + ws.UsedRange.Rows.Count - 1
    If usedLastRow < FIRST_DATA_ROW Then Exit Sub

    Dim sampleIdx As Long, startCol As Long
    Dim lastDataRow As Long
    Dim clearStart As Long

    For sampleIdx = 0 To MAX_SAMPLES_PER_SHEET - 1
        startCol = sampleIdx * COLS_PER_SAMPLE + 1
        ' Shared bound with ValidateSamplePuffs (audit M-b). Local NOT named
        ' lastAfterWeightRow: VBA names are case-insensitive, so that local
        ' would shadow the LastAfterWeightRow function and break the call.
        lastDataRow = LastAfterWeightRow(ws, startCol)

        clearStart = lastDataRow + 1
        If clearStart <= usedLastRow Then
            ws.Range(ws.Cells(clearStart, startCol), _
                     ws.Cells(usedLastRow, startCol + COLS_PER_SAMPLE - 1)) _
              .ClearContents
        End If
    Next sampleIdx
End Sub

Private Sub MakeWorkbookOpenable(wb As Workbook)
    ' Guarantee a saved copy opens to its data, not a blank gray screen.
    ' Excel bakes the workbook-window-hidden state into the file, so a file
    ' saved while its window is hidden (we hide the staging copy so the user
    ' never sees it) opens hidden - you would have to Window > Unhide to see
    ' anything. Restore a visible window, make every data sheet visible, and
    ' open on a real test sheet.
    On Error Resume Next
    Application.Windows(wb.Name).Visible = True
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 10) <> "_Template_" And ws.Name <> "_Macro_Install" Then
            ws.Visible = xlSheetVisible
        End If
    Next
    ' Prefer to land on a real test-data sheet; fall back to the first visible.
    Dim target As Worksheet, nm As Variant
    For Each nm In CanonicalDataSheets()
        If WorkbookHasSheetIn(wb, CStr(nm)) Then
            Set target = wb.Worksheets(CStr(nm))
            Exit For
        End If
    Next
    If target Is Nothing Then
        For Each ws In wb.Worksheets
            If ws.Visible = xlSheetVisible Then
                Set target = ws
                Exit For
            End If
        Next
    End If
    If Not target Is Nothing Then target.Activate
    On Error GoTo 0
End Sub

' ============================================================================
' Post-upload reset of the LIVE workbook
' ============================================================================

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
    If Err.Number <> 0 Then StampLog "WARN: reset scaffolding error: " & Err.Description
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
    If fresh Is Nothing Then
        Application.DisplayAlerts = savedAlerts
        StampLog "  '" & sheetName & "': snapshot copy added no sheet - cannot restore."
        Exit Function
    End If
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

' ============================================================================
' Post-upload reset: revert each uploaded sheet from an internal blank snapshot
' ============================================================================

Private Function TemplateSheetName(idx As Long) As String
    ' Internal snapshot sheet name for canonical sheet #idx. The "_Template_"
    ' prefix makes ApplySheetVisibility auto-very-hide it and the trim step
    ' auto-exclude it from distributed copies. Positional (idx ties to the
    ' CanonicalDataSheets order) to stay within Excel's 31-char sheet-name
    ' limit; re-run RebuildBlankTemplates if you change CanonicalDataSheets.
    TemplateSheetName = "_Template_" & Format$(idx, "00")
End Function

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

Private Function AddSnapshotFrom(srcWs As Worksheet, idx As Long) As Boolean
    ' Copy srcWs into THIS workbook as _Template_<idx> (left visible; the caller
    ' very-hides snapshots afterwards). Returns True on success.
    '
    ' Robust against two failure modes that previously clobbered data:
    '   * copies AFTER a VISIBLE anchor (the upload sheet) - copying relative to
    '     a very-hidden last sheet can make Excel misplace the copy;
    '   * finds the newly added sheet by NAME-DIFF, so it can never rename a
    '     pre-existing sheet, and returns False (touching nothing) if the copy
    '     added no sheet at all.
    AddSnapshotFrom = False

    Dim tplName As String
    tplName = TemplateSheetName(idx)
    If WorkbookHasSheetIn(ThisWorkbook, tplName) Then
        ThisWorkbook.Worksheets(tplName).Delete   ' caller has DisplayAlerts=False
    End If

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        seen(ws.Name) = True
    Next

    On Error Resume Next
    srcWs.Copy After:=ThisWorkbook.Worksheets(UPLOAD_SHEET_NAME)
    On Error GoTo 0

    Dim snap As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not seen.Exists(ws.Name) Then
            Set snap = ws
            Exit For
        End If
    Next
    If snap Is Nothing Then Exit Function   ' copy added no sheet - report via caller

    snap.Name = tplName
    AddSnapshotFrom = True
End Function

Private Function ResetSheetToBlankWithReview(liveWb As Workbook, sheetName As String, idx As Long) As Boolean
    ' Non-destructive reset: the live data sheet BECOMES a "<base> - Review" copy
    ' (a rename, not a delete), and a pristine blank from the internal snapshot
    ' (_Template_NN) takes its place + tab position. Transactional: on any failure
    ' the sheet is renamed back to its canonical name and left fully intact.
    ' Returns True only when the reset fully succeeded; every skip/fail path
    ' leaves the sheet intact and returns False.
    Dim tplName As String
    tplName = TemplateSheetName(idx)
    If Not WorkbookHasSheetIn(liveWb, tplName) Then
        StampLog "  '" & sheetName & "': no internal snapshot (" & tplName & _
                 ") - not reset (left intact). Build snapshots first (RebuildBlankTemplates)."
        Exit Function
    End If
    If Not SheetExists(sheetName) Then Exit Function

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
        Exit Function
    End If

    fresh.Name = sheetName
    fresh.Visible = xlSheetVisible        ' ApplySheetVisibility sets the real state

    ' 3) Move the Review sheet to the end to keep the working area uncluttered.
    orig.Move After:=liveWb.Worksheets(liveWb.Worksheets.Count)

    Application.DisplayAlerts = savedAlerts
    StampLog "  '" & sheetName & "': reset; data preserved as '" & reviewName & "'."
    ResetSheetToBlankWithReview = True
    Exit Function

Fail:
    On Error Resume Next
    If SheetExists(reviewName) And Not SheetExists(sheetName) Then
        liveWb.Worksheets(reviewName).Name = sheetName
    End If
    Application.DisplayAlerts = savedAlerts
    On Error GoTo 0
    StampLog "  '" & sheetName & "': reset FAILED (" & Err.Description & ") - left as-is."
End Function

Public Sub RebuildBlankTemplates()
    ' Build/refresh the internal blank-template snapshots the post-upload reset
    ' restores from. Snapshots each canonical data sheet's CURRENT state into a
    ' very-hidden _Template_NN sheet.
    '
    ' POKA-YOKE: refuses if any canonical sheet contains entered samples, so you
    ' cannot bake real data into the templates. Run on a BLANK workbook (a fresh
    ' template, or right after clearing). Re-run after changing a sheet's layout.
    Dim arr As Variant, i As Long, sheetName As String
    arr = CanonicalDataSheets()

    Dim dirty As String
    dirty = ""
    For i = LBound(arr) To UBound(arr)
        sheetName = CStr(arr(i))
        If SheetExists(sheetName) Then
            If SheetHasPopulatedSamples(ThisWorkbook.Worksheets(sheetName)) Then
                dirty = dirty & vbLf & "  - " & sheetName
            End If
        End If
    Next
    If Len(dirty) > 0 Then
        MsgBox "Rebuild aborted - these sheets contain data:" & dirty & vbLf & vbLf & _
               "Snapshot only a BLANK workbook (clear the data first).", _
               vbExclamation, "Rebuild Blank Templates"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim savedAlerts As Boolean
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Cleanup

    Dim made As Long, failed As String
    made = 0
    failed = ""
    For i = LBound(arr) To UBound(arr)
        sheetName = CStr(arr(i))
        If SheetExists(sheetName) Then
            If AddSnapshotFrom(ThisWorkbook.Worksheets(sheetName), i) Then
                made = made + 1
            Else
                failed = failed & vbLf & "  - " & sheetName & " (" & TemplateSheetName(i) & ")"
            End If
        End If
    Next

    ' Hide the snapshots. Activate a non-snapshot sheet first so we never try to
    ' hide the active sheet (Excel refuses to very-hide the active sheet).
    On Error Resume Next
    ThisWorkbook.Worksheets(UPLOAD_SHEET_NAME).Activate
    On Error GoTo Cleanup
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If Left$(w.Name, 10) = "_Template_" And w.Name <> "_Template_Master" Then
            On Error Resume Next
            w.Visible = xlSheetVeryHidden
            On Error GoTo Cleanup
        End If
    Next

    StampLog "RebuildBlankTemplates: created " & made & " snapshot(s)." & _
             IIf(Len(failed) > 0, " FAILED:" & Replace(failed, vbLf, " "), "")
    If Len(failed) > 0 Then
        MsgBox "Created " & made & " snapshot(s), but the COPY FAILED for:" & failed & _
               vbLf & vbLf & "Nothing else was changed - please report this.", _
               vbExclamation, "Rebuild Blank Templates"
    Else
        MsgBox "Created " & made & " blank-template snapshot(s).", vbInformation, _
               "Rebuild Blank Templates"
    End If

Cleanup:
    Application.DisplayAlerts = savedAlerts
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub SeedBlankTemplatesFromFile()
    ' Seed the internal blank-template snapshots (_Template_NN) from an EXTERNAL
    ' blank workbook (e.g. a fresh "Automated Testing Template - DVE.xlsm"). Use
    ' this to add reset snapshots to a workbook that ALREADY has real data, where
    ' RebuildBlankTemplates would refuse. Your data sheets are NOT touched - only
    ' the hidden _Template_NN snapshots are (re)created from the source workbook.
    Dim srcPath As String
    srcPath = PickWorkbookFile("Pick a BLANK template workbook to snapshot from")
    If Len(srcPath) = 0 Then Exit Sub
    If StrComp(srcPath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "Pick a different, BLANK workbook - not this one." & vbLf & _
               "(To snapshot THIS workbook, use RebuildBlankTemplates.)", vbExclamation
        Exit Sub
    End If

    Dim savedSec As Long
    savedSec = Application.AutomationSecurity
    Application.AutomationSecurity = 3   ' msoAutomationSecurityForceDisable: no macro prompt on Open
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim savedAlerts As Boolean
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Dim src As Workbook
    On Error Resume Next
    Set src = Application.Workbooks.Open(fileName:=srcPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo Cleanup
    If src Is Nothing Then
        MsgBox "Could not open:" & vbLf & srcPath, vbExclamation
        GoTo Cleanup
    End If

    Dim arr As Variant, i As Long, sheetName As String
    arr = CanonicalDataSheets()

    ' Warn (don't hard-block) if the source carries entered samples.
    Dim dirty As String
    dirty = ""
    For i = LBound(arr) To UBound(arr)
        sheetName = CStr(arr(i))
        If WorkbookHasSheetIn(src, sheetName) Then
            If SheetHasPopulatedSamples(src.Worksheets(sheetName)) Then
                dirty = dirty & vbLf & "  - " & sheetName
            End If
        End If
    Next
    If Len(dirty) > 0 Then
        If MsgBox("The source has entered samples in:" & dirty & vbLf & vbLf & _
                  "Seed anyway? (Recommended: pick a BLANK source.)", _
                  vbYesNo + vbExclamation, "Seed Blank Templates") <> vbYes Then
            GoTo Cleanup
        End If
    End If

    Dim made As Long, failed As String
    made = 0
    failed = ""
    For i = LBound(arr) To UBound(arr)
        sheetName = CStr(arr(i))
        If WorkbookHasSheetIn(src, sheetName) Then
            If AddSnapshotFrom(src.Worksheets(sheetName), i) Then
                made = made + 1
            Else
                failed = failed & vbLf & "  - " & sheetName & " (" & TemplateSheetName(i) & ")"
            End If
        End If
    Next

    src.Close SaveChanges:=False
    Set src = Nothing

    ' Hide snapshots; activate a non-snapshot sheet first.
    On Error Resume Next
    ThisWorkbook.Worksheets(UPLOAD_SHEET_NAME).Activate
    On Error GoTo Cleanup
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If Left$(w.Name, 10) = "_Template_" And w.Name <> "_Template_Master" Then
            On Error Resume Next
            w.Visible = xlSheetVeryHidden
            On Error GoTo Cleanup
        End If
    Next

    StampLog "SeedBlankTemplatesFromFile: created " & made & " snapshot(s) from " & srcPath & _
             IIf(Len(failed) > 0, " FAILED:" & Replace(failed, vbLf, " "), "")
    If Len(failed) > 0 Then
        MsgBox "Created " & made & " snapshot(s) from:" & vbLf & srcPath & vbLf & vbLf & _
               "But the COPY FAILED for:" & failed & vbLf & "Nothing else was changed.", _
               vbExclamation, "Seed Blank Templates"
    Else
        MsgBox "Created " & made & " blank-template snapshot(s) from:" & vbLf & srcPath & _
               vbLf & vbLf & "Your data sheets were not touched. Save the workbook now.", _
               vbInformation, "Seed Blank Templates"
    End If

Cleanup:
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    Application.AutomationSecurity = savedSec
    Application.DisplayAlerts = savedAlerts
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Function PickWorkbookFile(promptTitle As String) As String
    Dim r As Variant
    r = Application.GetOpenFilename( _
        "Excel workbooks (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , promptTitle)
    If VarType(r) = vbBoolean Then
        PickWorkbookFile = ""
    Else
        PickWorkbookFile = CStr(r)
    End If
End Function

' ============================================================================
' Checklist (validates only sheets the user selected to upload)
' ============================================================================

Public Function RunChecklist() As Collection
    Dim failures As New Collection

    ' Required inputs
    If Len(Trim$(GetNamed("DV_SynologyPath"))) = 0 Then failures.Add "DV_SynologyPath is empty"
    If Len(Trim$(GetNamed("DV_LocalPath"))) = 0 Then failures.Add "DV_LocalPath is empty"

    ' Selection must define at least one TRUE that maps to a real, populated sheet.
    Dim selection As Object
    Set selection = ReadSelection()
    Dim selectedAny As Boolean: selectedAny = False
    Dim populatedSelected As Boolean: populatedSelected = False

    Dim sheetName As Variant
    For Each sheetName In CanonicalDataSheets()
        Dim norm As String
        norm = NormalizeSheetName(CStr(sheetName))
        If selection.Exists(norm) Then
            If selection(norm) Then
                selectedAny = True
                If SheetExists(CStr(sheetName)) Then
                    If SheetHasPopulatedSamples(ThisWorkbook.Worksheets(CStr(sheetName))) Then
                        populatedSelected = True
                        ValidateSheet ThisWorkbook.Worksheets(CStr(sheetName)), failures
                    Else
                        failures.Add CStr(sheetName) & ": selected but no populated samples"
                    End If
                Else
                    failures.Add CStr(sheetName) & ": selected but sheet is missing"
                End If
            End If
        End If
    Next

    If Not selectedAny Then failures.Add "No tests selected in DV_TestSelection"

    Set RunChecklist = failures
End Function

Private Sub ValidateSheet(ws As Worksheet, failures As Collection)
    Dim sampleIdx As Long, startCol As Long, sidCol As Long
    Dim sampleID As String, heatTech As String, thisAddr As String
    Dim seenIDs As Object
    Set seenIDs = CreateObject("Scripting.Dictionary")

    For sampleIdx = 0 To MAX_SAMPLES_PER_SHEET - 1
        startCol = sampleIdx * COLS_PER_SAMPLE + 1
        sidCol = startCol + 5
        sampleID = Trim$(SafeString(ws.Cells(1, sidCol).value))
        heatTech = Trim$(SafeString(ws.Cells(1, startCol + 7).value))

        If Len(sampleID) > 0 Then
            thisAddr = ws.Cells(1, sidCol).Address(False, False)
            If seenIDs.Exists(sampleID) Then
                ' Two sample blocks share one name - they would collide in the
                ' database. Point the operator at both cells so it is a 5-second fix.
                failures.Add ws.Name & ": duplicate sample ID '" & sampleID & "'" & _
                    " - sample #" & (sampleIdx + 1) & " (cell " & thisAddr & _
                    ") repeats sample #" & seenIDs(sampleID)(0) & " (cell " & _
                    seenIDs(sampleID)(1) & "). Check for duplicate sample names; " & _
                    "rename one so every sample on the sheet is unique."
            Else
                seenIDs.Add sampleID, Array(sampleIdx + 1, thisAddr)
            End If
            If Len(heatTech) = 0 Then
                failures.Add ws.Name & " / " & sampleID & ": heating technology empty"
            End If
            ValidateSamplePuffs ws, sampleID, startCol, failures
        End If
    Next sampleIdx
End Sub

Private Sub ValidateSamplePuffs(ws As Worksheet, sampleID As String, _
                                startCol As Long, failures As Collection)
    Dim row As Long
    Dim puffs As Double, prevPuffs As Double
    Dim bWeight As Double, aWeight As Double
    Dim tpm As Double
    Dim seenPuff As Boolean

    seenPuff = False
    prevPuffs = 0#

    Dim lastRow As Long
    lastRow = LastAfterWeightRow(ws, startCol)
    For row = FIRST_DATA_ROW To lastRow
        If Not IsBlankRow(ws, row, startCol) Then
            puffs = SafeDouble(ws.Cells(row, startCol).value)
            bWeight = SafeDouble(ws.Cells(row, startCol + 1).value)
            aWeight = SafeDouble(ws.Cells(row, startCol + 2).value)
            tpm = SafeDouble(ws.Cells(row, startCol + 8).value)

            If puffs > 0# Then
                If seenPuff And puffs <= prevPuffs Then
                    failures.Add ws.Name & " / " & sampleID & _
                                 ": puff sequence not strictly increasing at row " & row
                End If
                prevPuffs = puffs
                seenPuff = True
            End If

            If bWeight > 0# And aWeight > 0# And bWeight <= aWeight Then
                failures.Add ws.Name & " / " & sampleID & _
                             ": mass-before <= mass-after at row " & row
            End If

            If tpm < 0# Or tpm > TPM_MAX_PLAUSIBLE Then
                failures.Add ws.Name & " / " & sampleID & _
                             ": implausible TPM (" & tpm & ") at row " & row
            End If
        End If
    Next row
End Sub

Private Function IsBlankRow(ws As Worksheet, row As Long, startCol As Long) As Boolean
    IsBlankRow = IsEmptyCell(ws.Cells(row, startCol)) And _
                 IsEmptyCell(ws.Cells(row, startCol + 1)) And _
                 IsEmptyCell(ws.Cells(row, startCol + 2))
End Function

Private Function IsEmptyCell(c As Range) As Boolean
    IsEmptyCell = (IsEmpty(c.value) Or Len(Trim$(SafeString(c.value))) = 0)
End Function

' ============================================================================
' Named-range + logging helpers
' ============================================================================

Private Function GetNamed(rangeName As String) As String
    On Error Resume Next
    GetNamed = SafeString(ThisWorkbook.Names(rangeName).RefersToRange.value)
    On Error GoTo 0
End Function

Private Sub SetNamed(rangeName As String, value As String)
    On Error Resume Next
    ThisWorkbook.Names(rangeName).RefersToRange.value = value
    On Error GoTo 0
End Sub

Private Sub ClearLog()
    SetNamed "DV_Log", ""
End Sub

Private Sub StampLog(msg As String)
    Dim line As String
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & msg
    Dim current As String
    current = GetNamed("DV_Log")
    If Len(current) > 0 Then
        SetNamed "DV_Log", current & vbLf & line
    Else
        SetNamed "DV_Log", line
    End If
End Sub

Private Sub WriteFailures(failures As Collection)
    Dim f As Variant
    For Each f In failures
        StampLog "  - " & CStr(f)
    Next
End Sub

Private Sub ShowFailures(title As String, failures As Collection)
    Dim msg As String, n As Long, shown As Long
    shown = 0
    For n = 1 To failures.Count
        If shown >= 20 Then
            msg = msg & vbLf & "   ...and " & (failures.Count - shown) & _
                  " more. Fix the issues above and run again - the rest will surface."
            Exit For
        End If
        msg = msg & vbLf & "  - " & CStr(failures(n))
        shown = shown + 1
    Next
    MsgBox "Found " & failures.Count & " issue(s):" & vbLf & msg, vbExclamation, title
End Sub

' ============================================================================
' Utilities
' ============================================================================

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

Private Function SafeString(v As Variant) As String
    If IsError(v) Then
        SafeString = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeString = ""
    Else
        SafeString = CStr(v)
    End If
End Function

Private Function SafeDouble(v As Variant) As Double
    If IsNumeric(v) Then SafeDouble = CDbl(v)
End Function

Private Function AppendBackslash(p As String) As String
    If Right$(p, 1) = "\" Then
        AppendBackslash = p
    Else
        AppendBackslash = p & "\"
    End If
End Function

Private Function ResolveDataViewerExe() As String
    Dim override As String
    override = GetNamed("DV_DataViewerExe")
    If Len(Trim$(override)) > 0 Then
        ResolveDataViewerExe = override
    Else
        ResolveDataViewerExe = DEFAULT_DATAVIEWER_EXE
    End If
End Function

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
    If StrComp(clean.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then Err.Raise vbObjectError + 4, _
        "MakeTempXlsx", "Sheets.Copy left the live workbook active"

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
    ' 'eNum' == reserved word 'Enum' -> Compile error (bug 2). Renamed to errNum.
    Dim errNum As Long, eMsg As String
    errNum = Err.Number: eMsg = Err.Description
    On Error Resume Next
    If Not clean Is Nothing Then clean.Close SaveChanges:=False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedScreenUpdating
    On Error GoTo 0
    Err.Raise errNum, "MakeTempXlsx", eMsg
End Function


' ============================================================================
' Path-picker button handlers
' ============================================================================
Public Sub Btn_PickSynologyFolder()
    PickFolderInto "DV_SynologyPath", "Choose the Synology data folder"
    RefreshPathLabels
    PersistSettings
End Sub

Public Sub Btn_PickLocalFolder()
    PickFolderInto "DV_LocalPath", "Choose the local data folder"
    RefreshPathLabels
    PersistSettings
End Sub



Private Sub PickFolderInto(ByVal namedRange As String, ByVal title As String)
    ' Modern Windows folder picker (same dialog family as the file picker),
    ' not the old Shell.BrowseForFolder tree.
    Dim fd As Object
    Set fd = Application.FileDialog(4)        ' msoFileDialogFolderPicker
    fd.title = title
    fd.AllowMultiSelect = False
    Dim startPath As String
    startPath = GetNamed(namedRange)
    If Len(startPath) > 0 Then fd.InitialFileName = AppendBackslash(startPath)
    If fd.Show = -1 Then
        If fd.SelectedItems.Count > 0 Then SetNamed namedRange, fd.SelectedItems(1)
    End If
End Sub

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

' ============================================================================
' Ribbon callbacks (TPM Testing tab -> "DataViewer Upload" group)
' ============================================================================
Public Sub Ribbon_UploadAll(control As IRibbonControl)
    Btn_UploadAll
End Sub
Public Sub Ribbon_UploadCheckpoint(control As IRibbonControl)
    Btn_UploadCheckpoint
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

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

' ===== v1.1 Feature 2: Specify Test Name (replaces Dry-Run Checklist) =====

Private Function BaseFileName(ByVal nm As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(nm, ".")
    If dotPos > 0 Then BaseFileName = Left$(nm, dotPos - 1) Else BaseFileName = nm
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    ' Strips Windows-illegal filename characters PLUS [ and ]: brackets are
    ' legal on disk but Excel's SaveAs rejects them in workbook names.
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "[", "]")
    For Each ch In bad
        s = Replace$(s, CStr(ch), "")
    Next ch
    SanitizeFileName = Trim$(s)
End Function

Private Function RenameWorkbookTo(ByVal newFullPath As String, _
                                  Optional ByVal silent As Boolean = False) As Boolean
    ' Returns True when the workbook ended up at newFullPath (including the
    ' already-there short-circuit); False when the user declined the overwrite
    ' confirm or the rename failed.
    '   silent=True is the Upload-All auto-revert (bug 1): the data has already
    '   been reset to a blank template, so reverting to the template name and
    '   overwriting the prior (also-blank) template is intended -- no prompt, and
    '   the old working copy is renamed away so exactly one file remains.
    Dim oldPath As String: oldPath = ThisWorkbook.FullName
    If StrComp(oldPath, newFullPath, vbTextCompare) = 0 Then
        RenameWorkbookTo = True
        Exit Function
    End If
    On Error GoTo Fail
    ' POKA-YOKE (audit C2): the user-initiated "Specify Test Name" path never
    ' silently SaveAs over an existing workbook (a sibling project copy could
    ' share the name). The silent auto-revert skips this: see above.
    If Not silent And Len(Dir$(newFullPath)) > 0 Then
        If MsgBox("A file already exists at:" & vbCrLf & newFullPath & vbCrLf & vbCrLf & _
                  "Overwrite it? If another project copy uses this name, choose No " & _
                  "and pick a different test name.", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Rename workbook") <> vbYes Then Exit Function
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
    RenameWorkbookTo = True
    Exit Function
Fail:
    Application.DisplayAlerts = True
    MsgBox "Could not rename the file:" & vbCrLf & Err.Description, vbExclamation, "Specify Test Name"
End Function

Private Sub RevertToOriginalName()
    Dim orig As String: orig = Trim$(GetNamed("DV_OrigFileName"))
    If Len(orig) = 0 Then Exit Sub
    Dim target As String: target = ThisWorkbook.Path & "\" & orig & ".xlsm"
    If StrComp(ThisWorkbook.FullName, target, vbTextCompare) = 0 Then Exit Sub
    ' Bug 1: silent revert (the workbook is already reset to a blank template).
    If Not RenameWorkbookTo(target, silent:=True) Then
        StampLog "WARN: revert to original name failed; continuing under current name"
    End If
End Sub

Public Sub Btn_SpecifyName()
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Please save the workbook once before naming it.", vbExclamation, "Specify Test Name"
        Exit Sub
    End If
    Dim orig As String: orig = Trim$(GetNamed("DV_OrigFileName"))
    If Len(orig) = 0 Then
        orig = BaseFileName(ThisWorkbook.Name)
        SetNamed "DV_OrigFileName", orig
    End If
    Dim r As Variant
    r = Application.InputBox( _
        Prompt:="Enter a project / test name to label this file." & vbCrLf & _
                "Leave blank to restore the original name (" & orig & ").", _
        Title:="Specify Test Name", Default:="", Type:=2)
    If VarType(r) = vbBoolean Then Exit Sub
    Dim proj As String: proj = SanitizeFileName(CStr(r))
    Dim target As String
    If Len(proj) = 0 Then
        target = orig & ".xlsm"
    Else
        ' "(do not send)" marks the renamed working copy as a non-deliverable at
        ' the exact place the send-the-template mistake happens: the Explorer /
        ' Outlook attach dialog (audit P2). Upload All still reverts to orig.
        target = proj & " - " & orig & " (do not send).xlsm"
    End If
    If Not RenameWorkbookTo(ThisWorkbook.Path & "\" & target) Then
        StampLog "WARN: rename declined/failed"
    End If
End Sub

Public Sub Ribbon_SpecifyName(control As IRibbonControl)
    Btn_SpecifyName
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
    returnedVal = ResolveDataViewerExe()   ' show the effective exe (default if unset)
End Sub
Public Sub GetSynPathTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Synology folder:" & vbLf & GetNamed("DV_SynologyPath")
End Sub
Public Sub GetLocPathTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Local folder:" & vbLf & GetNamed("DV_LocalPath")
End Sub
Public Sub GetExePathTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "DataViewer.exe:" & vbLf & ResolveDataViewerExe()
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
        PersistSettings
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
End Function

Public Sub Ribbon_PickDataViewer(control As IRibbonControl)
    Btn_PickDataViewerExe
End Sub
Public Sub Ribbon_Instructions(control As IRibbonControl)
    Btn_ShowInstructions
End Sub

Public Sub Btn_ShowSheetGuide()
    MsgBox SheetGuideText(), vbInformation, "TPM Test Sheet - Layout & Dropdowns"
End Sub

Private Function SheetGuideText() As String
    Dim s As String, NL As String
    NL = vbCrLf
    s = "HOW THE TPM TEST SHEET WORKS" & NL & NL
    s = s & "Each test sheet is a row of 12-column 'sample blocks' placed" & NL
    s = s & "side by side - one device per block. Rows 1-3 hold the device" & NL
    s = s & "info, row 4 the column headings, rows 5 and down your readings." & NL & NL
    s = s & "DROPDOWNS (click the arrow that appears in the cell)" & NL
    s = s & "  - Heating Technology (top of each block):" & NL
    s = s & "      CCELL3.0, EVOMAX, EVO, SE, T51, Competitor." & NL
    s = s & "  - Voltage (header, row 3): choose 'Voltage (Constant)' or" & NL
    s = s & "      'Voltage (Variable)', then type the value beside it." & NL
    s = s & "  - Puffs (first column, rows 5+): type ANY number and the" & NL
    s = s & "      column auto-fills below it (row above + your number)." & NL
    s = s & "      Typing in the FIRST row seeds the whole column. Pick" & NL
    s = s & "      'custom' to clear the column, then PASTE your own" & NL
    s = s & "      numbers (typing a single number auto-fills below it)." & NL
    s = s & "  - Smell (the 'Smell' column): a 0-and-up number rating." & NL & NL
    s = s & "AUTO-CALCULATED - DON'T TYPE IN THESE" & NL
    s = s & "  Clog (from Draw Pressure: blank, 'Light Clog' or 'Heavy Clog');" & NL
    s = s & "  Power (from voltage + resistance); TPM, TPM Power Density," & NL
    s = s & "  Variation % and Oil Consumed (the right-hand columns); and the" & NL
    s = s & "  Average TPM / Std Dev. They recalculate from what you enter." & NL
    s = s & "  If a formula gets overwritten, use Sample Tools > Reset" & NL
    s = s & "  Formulas to restore it."
    SheetGuideText = s
End Function

Public Sub Ribbon_SheetGuide(control As IRibbonControl)
    Btn_ShowSheetGuide
End Sub

Public Sub Btn_ShowSampleTools()
    MsgBox SampleToolsText(), vbInformation, "Sample Tools & Shortcuts"
End Sub

Private Function SampleToolsText() As String
    Dim s As String, NL As String
    NL = vbCrLf
    s = "SAMPLE TOOLS & SHORTCUTS" & NL & NL
    s = s & "PUFFS AUTO-FILL" & NL
    s = s & "  Type any number into the puffs column and every row below" & NL
    s = s & "  becomes 'the row above + your number'. In the first data row" & NL
    s = s & "  it seeds the whole block. Pick 'custom' to clear the column," & NL
    s = s & "  then PASTE your own numbers (typing a single number" & NL
    s = s & "  auto-fills below it)." & NL & NL
    s = s & "RIBBON (TPM Testing tab)" & NL
    s = s & "  - Add Sample: adds a fresh, empty block on the right" & NL
    s = s & "    (dropdowns and formulas included)." & NL
    s = s & "  - Remove Sample: deletes the right-most block (asks first)." & NL
    s = s & "  - Reset Formulas: click inside a block first; restores that" & NL
    s = s & "    block's calculated formulas only - your data is untouched." & NL & NL
    s = s & "JUMP BETWEEN SAMPLES (keeps your current row)" & NL
    s = s & "  - Ctrl + Shift + .   next sample block" & NL
    s = s & "  - Ctrl + Shift + ,   previous sample block" & NL
    s = s & "  - First / Prev / Next / Last buttons do the same; 'Last'" & NL
    s = s & "    jumps to the right-most block that has a header." & NL & NL
    s = s & "  Up to 24 sample blocks per sheet."
    SampleToolsText = s
End Function

Public Sub Ribbon_SampleTools(control As IRibbonControl)
    Btn_ShowSampleTools
End Sub

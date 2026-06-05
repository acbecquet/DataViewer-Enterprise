Attribute VB_Name = "DataViewerUpload"
Option Explicit

' ============================================================================
' DataViewer Enterprise - upload sidecar module.
'
' Workflow:
'   1. Workbook opens -> only "Lifetime Test" + "DataViewer Upload" + "Test
'      SOP's" are visible. Everything else (incl. other canonical data sheets,
'      _Macro_Install, _Template_Master, _Template_<Test>) is hidden.
'   2. On the "DataViewer Upload" sheet, the user toggles TRUE/FALSE next to
'      each canonical test name in the DV_TestSelection range. Toggling TRUE
'      unhides the matching test sheet; FALSE re-hides it.
'   3. Tech fills out the selected test sheets. Save/close/reopen freely -
'      visibility is driven entirely by the DV_TestSelection cells.
'   4. Click "Upload All". The macro:
'        a. runs the checklist on selected sheets,
'        b. stages a trimmed .xlsm (selected + populated + Test SOP's),
'        c. copies it to <DV_SynologyPath>\<DV_FileName>.xlsm and
'           <DV_LocalPath>\<DV_FileName>.xlsm,
'        d. materializes a .xlsx for DataViewer ingestion (-> Postgres),
'        e. reverts each selected sheet to its internal blank snapshot
'           (_Template_NN, built by RebuildBlankTemplates); a sheet with no
'           snapshot is left intact and logged,
'        f. resets DV_TestSelection to default (only Lifetime Test = TRUE),
'        g. re-hides all sheets except Lifetime Test + Test SOP's + DataViewer
'           Upload, so the workbook is fresh for the next session.
'
' Named ranges that must exist on the workbook:
'   DV_FileName         single-cell text  - base filename (no extension)
'   DV_SynologyPath     single-cell text  - destination folder #1
'   DV_LocalPath        single-cell text  - destination folder #2
'   DV_Status           single-cell text  - macro status output
'   DV_Log              single-cell text  - macro log output (multi-line)
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
    Dim sheetName As Variant
    For Each sheetName In CanonicalDataSheets()
        Dim wantVisible As Boolean
        wantVisible = False
        If selection.Exists(NormalizeSheetName(CStr(sheetName))) Then
            wantVisible = selection(NormalizeSheetName(CStr(sheetName)))
        End If
        If SheetExists(CStr(sheetName)) Then
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
            If wantVisible Then
                ws.Visible = xlSheetVisible
            Else
                ws.Visible = xlSheetHidden
            End If
        End If
    Next

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

Public Sub Btn_DryRunChecklist()
    ClearLog
    SetNamed "DV_Status", "Running checklist..."
    StampLog "Checklist dry-run started"

    Dim failures As Collection
    Set failures = RunChecklist()

    If failures.Count = 0 Then
        SetNamed "DV_Status", "OK"
        StampLog "Checklist passed"
        MsgBox "Checklist passed - ready to upload.", vbInformation, "Dry-Run Checklist"
    Else
        WriteFailures failures
        SetNamed "DV_Status", "Failed: " & failures.Count & " checklist issue(s)"
        ShowFailures "Dry-Run Checklist", failures
    End If
End Sub

Public Sub Btn_UploadAll()
    On Error GoTo Failed
    ClearLog
    SetNamed "DV_Status", "Starting upload..."
    StampLog "Upload started"

    ' Ask for the descriptive file name up front (pre-filled with the last one).
    Dim fName As String
    fName = PromptForFileName()
    If Len(fName) = 0 Then
        SetNamed "DV_Status", "Cancelled (no file name)"
        StampLog "Upload cancelled - no file name entered"
        Exit Sub
    End If
    ' Reject characters that are illegal in a Windows filename, so the operator
    ' gets a clear message now instead of a cryptic file-copy error later.
    If fName Like "*[" & "\/:*?<>|" & Chr$(34) & "]*" Then
        SetNamed "DV_Status", "Cancelled (invalid file name)"
        MsgBox "The file name can't contain any of these characters:" & vbLf & _
               "   \ / : * ? " & Chr$(34) & " < > |", vbExclamation, "Upload file name"
        Exit Sub
    End If
    SetNamed "DV_FileName", fName

    ' --- 1. Checklist (abort on failure) ---
    Dim failures As Collection
    Set failures = RunChecklist()
    If failures.Count > 0 Then
        WriteFailures failures
        SetNamed "DV_Status", "Failed: checklist has " & failures.Count & " issue(s)"
        ShowFailures "Upload All", failures
        Exit Sub
    End If
    StampLog "Checklist passed"

    Dim baseName As String, synPath As String, locPath As String
    baseName = Trim$(GetNamed("DV_FileName"))
    synPath = Trim$(GetNamed("DV_SynologyPath"))
    locPath = Trim$(GetNamed("DV_LocalPath"))

    ' --- 2. Persist in-memory edits so the copies pick them up ---
    StampLog "Saving workbook"
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    ' --- 3. Resolve destinations. The distributed copies are macro-free
    '        .xlsx: only the source template carries VBA. A copy WITH macros
    '        opened to a blank gray window and re-hid its own sheets on open;
    '        MakeTempXlsx's SaveAs FileFormat:=51 strips the VBA entirely.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim synDest As String, locDest As String
    synDest = AppendBackslash(synPath) & baseName & ".xlsx"
    locDest = AppendBackslash(locPath) & baseName & ".xlsx"

    If Not fso.FolderExists(synPath) Then
        SetNamed "DV_Status", "Failed: Synology path not accessible: " & synPath
        MsgBox "Synology path not accessible:" & vbLf & synPath, vbExclamation, "Upload All"
        Exit Sub
    End If
    If Not fso.FolderExists(locPath) Then
        SetNamed "DV_Status", "Failed: Local path not accessible: " & locPath
        MsgBox "Local path not accessible:" & vbLf & locPath, vbExclamation, "Upload All"
        Exit Sub
    End If

    ' --- 3a. Build keep-list (selected + populated + Test SOP's) ---
    Dim keep As Object
    Set keep = BuildKeepList()
    StampLog "Trim keep-list size: " & keep.Count

    ' Abort if nothing real to upload (only Test SOP's in the keep-list).
    If keep.Count <= 1 Then
        SetNamed "DV_Status", "Failed: no selected sheets contain data"
        StampLog "Aborting: keep-list contains only Test SOP's"
        MsgBox "No selected sheets contain data.", vbExclamation, "Upload All"
        Exit Sub
    End If

    ' --- 3b. Stage a trimmed .xlsm in TEMP (intermediate only) ---
    Dim trimmedXlsm As String
    trimmedXlsm = Environ$("TEMP") & "\dvupload_staging_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & "_" & baseName & ".xlsm"
    fso.CopyFile ThisWorkbook.FullName, trimmedXlsm, True
    StampLog "Staging trimmed copy: " & trimmedXlsm
    TrimSheetsInWorkbook trimmedXlsm, keep

    ' --- 3c. Materialize ONE clean, macro-free .xlsx (FileFormat 51 strips
    '         the VBA). This single file is distributed to both destinations
    '         AND ingested by DataViewer. ---
    Dim cleanXlsx As String
    cleanXlsx = MakeTempXlsx(fso, trimmedXlsm, baseName)
    StampLog "Clean .xlsx: " & cleanXlsx

    On Error Resume Next
    fso.DeleteFile trimmedXlsm, True
    On Error GoTo 0

    ' --- 3d. Distribute the macro-free .xlsx to both destinations ---
    StampLog "Copy -> " & synDest
    fso.CopyFile cleanXlsx, synDest, True
    StampLog "Copy -> " & locDest
    fso.CopyFile cleanXlsx, locDest, True

    ' --- 5. Launch DataViewer; SingleInstance hands off to a running window ---
    Dim dvExe As String
    dvExe = ResolveDataViewerExe()
    If Not fso.FileExists(dvExe) Then
        SetNamed "DV_Status", "Failed: DataViewer.exe not found at " & dvExe
        MsgBox "DataViewer.exe not found at:" & vbLf & dvExe, vbExclamation, "Upload All"
        Exit Sub
    End If

    Dim cmd As String
    cmd = """" & dvExe & """ """ & cleanXlsx & """"
    StampLog "Shell: " & cmd
    Shell cmd, vbNormalFocus
    StampLog "Upload dispatched (DB write happens inside DataViewer)"

    ' --- 6. Reset the LIVE workbook so the technician starts fresh ---
    On Error GoTo PostDispatchFailed
    StampLog "Resetting live workbook"
    ResetLiveWorkbookAfterUpload keep
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    SetNamed "DV_Status", "OK"
    StampLog "Done"
    MsgBox "Upload complete." & vbLf & vbLf & baseName & ".xlsx was sent to the " & _
           "Synology and Local folders and opened in DataViewer." & vbLf & _
           "Each uploaded sheet was reset (a '- Review' copy was kept).", _
           vbInformation, "Upload All"
    Exit Sub

PostDispatchFailed:
    ' Upload itself succeeded; only the live-workbook reset hiccuped.
    StampLog "WARN: post-upload reset failed: " & Err.Description
    SetNamed "DV_Status", "OK (live reset partial - see log)"
    Exit Sub

Failed:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    StampLog "ERROR " & Err.Number & ": " & Err.Description
    SetNamed "DV_Status", "Failed: " & Err.Description
    MsgBox "Upload failed:" & vbLf & Err.Description, vbCritical, "Upload All"
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
        If Len(sampleID) > 0 Then
            SheetHasPopulatedSamples = True
            Exit Function
        End If
    Next
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

    ' Hide the staging workbook window so the user never sees the trimmed copy.
    On Error Resume Next
    Application.Windows(wb.Name).Visible = False
    On Error GoTo 0

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
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not keepNorm.Exists(NormalizeSheetName(ws.Name)) Then
            victims.Add ws.Name
        End If
    Next

    If victims.Count >= wb.Worksheets.Count Then
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
        wb.Worksheets(CStr(victim)).Visible = xlSheetVisible
        wb.Worksheets(CStr(victim)).Delete
        If Err.Number <> 0 Then
            StampLog "  Could not delete sheet '" & CStr(victim) & "': " & Err.Description
            deleteFailed = True
            Err.Clear
        End If
        On Error GoTo 0
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

Private Sub TrimSampleBlockTails(ws As Worksheet)
    ' For each sample block (12 columns wide), find the last row
    ' with a populated "After Weight" cell (col offset +2) and clear every
    ' cell in the block below that row.
    Dim usedLastRow As Long
    usedLastRow = ws.UsedRange.row + ws.UsedRange.Rows.Count - 1
    If usedLastRow < FIRST_DATA_ROW Then Exit Sub

    Dim sampleIdx As Long, startCol As Long
    Dim lastAfterWeightRow As Long
    Dim found As Range
    Dim clearStart As Long

    For sampleIdx = 0 To MAX_SAMPLES_PER_SHEET - 1
        startCol = sampleIdx * COLS_PER_SAMPLE + 1
        Set found = Nothing
        On Error Resume Next
        Set found = ws.Range(ws.Cells(FIRST_DATA_ROW, startCol + 2), _
                             ws.Cells(usedLastRow, startCol + 2)) _
                      .Find(What:="*", LookIn:=xlValues, _
                            SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
        On Error GoTo 0

        If found Is Nothing Then
            lastAfterWeightRow = FIRST_DATA_ROW - 1
        Else
            lastAfterWeightRow = found.row
        End If

        clearStart = lastAfterWeightRow + 1
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

Private Sub ResetLiveWorkbookAfterUpload(keep As Object)
    ' Revert each uploaded canonical data sheet to its pristine internal snapshot
    ' (_Template_NN, very hidden), then reset DV_TestSelection to default and
    ' re-hide everything so the workbook is fresh for the next session.
    '
    ' The snapshots live INSIDE this workbook (built by RebuildBlankTemplates),
    ' so the reset has NO external dependency. If a sheet's snapshot is missing,
    ' that sheet is left intact and logged - we never clear without a source.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Cleanup

    Dim arr As Variant, i As Long, sheetName As String
    arr = CanonicalDataSheets()
    For i = LBound(arr) To UBound(arr)
        sheetName = CStr(arr(i))
        If keep.Exists(sheetName) And SheetExists(sheetName) Then
            ResetSheetToBlankWithReview ThisWorkbook, sheetName, i
        End If
    Next

    ResetSelectionToDefault
    ApplySheetVisibility

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

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

    For row = FIRST_DATA_ROW To 10000
        If IsBlankRow(ws, row, startCol) Then Exit For

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
                  " more (see the log on the hidden _Settings sheet)."
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
    ' Per-run subdirectory so the bare filename is exactly "<baseName>.xlsx".
    ' DataViewer derives the DB filename from the path - no prefix wanted.
    Dim runDir As String, outXlsx As String
    runDir = Environ$("TEMP") & "\dvupload_" & Format$(Now, "yyyymmdd_hhnnss")
    If Not fso.FolderExists(runDir) Then fso.CreateFolder runDir
    outXlsx = runDir & "\" & baseName & ".xlsx"

    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(fileName:=sourceXlsm, ReadOnly:=True, _
                                        UpdateLinks:=0)
    ' Ensure the .xlsx opens cleanly if a human double-clicks it (DataViewer
    ' reads it headlessly regardless). ScreenUpdating is already off.
    MakeWorkbookOpenable wb
    wb.SaveAs fileName:=outXlsx, FileFormat:=51
    wb.Close SaveChanges:=False

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedScreenUpdating

    MakeTempXlsx = outXlsx
End Function


' ============================================================================
' Path-picker button handlers
' ============================================================================
Public Sub Btn_PickSynologyFolder()
    PickFolderInto "DV_SynologyPath", "Choose the Synology data folder"
    RefreshPathLabels
End Sub

Public Sub Btn_PickLocalFolder()
    PickFolderInto "DV_LocalPath", "Choose the local data folder"
    RefreshPathLabels
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
    s = s & "1.  Tick the tests you're running. Each ticked test's" & vbLf
    s = s & "     sheet appears so you can fill it in; untick to hide" & vbLf
    s = s & "     a sheet again." & vbLf & vbLf
    s = s & "2.  Set your three destinations once, from this ribbon:" & vbLf
    s = s & "     Pick Synology Folder, Pick Local Folder, and" & vbLf
    s = s & "     Pick DataViewer File. The chosen paths appear in" & vbLf
    s = s & "     the 'Active Folders' box." & vbLf & vbLf
    s = s & "3.  Optional: click Dry-Run Checklist to validate your" & vbLf
    s = s & "     data without uploading anything." & vbLf & vbLf
    s = s & "4.  Click Upload All and enter a file name when asked" & vbLf
    s = s & "     (Product + Test + Date). Your data is copied to the" & vbLf
    s = s & "     Synology and Local folders, opened in DataViewer," & vbLf
    s = s & "     and each uploaded sheet is reset; a '... - Review'" & vbLf
    s = s & "     copy is kept so you can see what was sent." & vbLf & vbLf
    s = s & "5.  Click Delete All Review Sheets to clear those" & vbLf
    s = s & "     review copies once you're finished with them." & vbLf & vbLf
    s = s & "Tip:  Test SOP's is always included in every upload," & vbLf
    s = s & "      whether or not its box is ticked."
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
    s = s & "  - Puffs (first column, rows 5+): pick a step - 1, 2, 5, 10," & NL
    s = s & "      20 or 50 - and the column auto-fills the running puff" & NL
    s = s & "      count (20 -> 20, 40, 60, ...). Pick 'custom' to clear it" & NL
    s = s & "      and type your own numbers." & NL
    s = s & "  - Clog (the 'Clog' column): Y or N per reading." & NL
    s = s & "  - Smell (the 'Smell' column): a 0-and-up number rating." & NL & NL
    s = s & "AUTO-CALCULATED - DON'T TYPE IN THESE" & NL
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
    s = s & "  Type a step (1, 2, 5, 10, 20 or 50) into the puffs column" & NL
    s = s & "  and every row below becomes 'the row above + step'. Put the" & NL
    s = s & "  step in the first data row to fill the whole block. Type" & NL
    s = s & "  'custom' to enter puff numbers by hand." & NL & NL
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

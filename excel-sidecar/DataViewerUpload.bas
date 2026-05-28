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
Private Const UPLOAD_SHEET_NAME As String = "DataViewer Upload"
Private Const DEFAULT_SELECTED_SHEET As String = "Lifetime Test"
Private Const SOPS_SHEET_NAME As String = "Test SOP's"

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

    ' Always-visible utility sheets.
    EnsureVisible UPLOAD_SHEET_NAME, xlSheetVisible
    EnsureVisible SOPS_SHEET_NAME, xlSheetVisible

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
        If StrComp(NormalizeSheetName(sheetName), _
                   NormalizeSheetName(DEFAULT_SELECTED_SHEET), _
                   vbTextCompare) = 0 Then
            selRange.Cells(row, 1).value = True
        Else
            selRange.Cells(row, 1).value = False
        End If
    Next row

Cleanup:
    Application.EnableEvents = True
End Sub

' ============================================================================
' Public button handlers
' ============================================================================

Public Sub Btn_DryRunChecklist()
    ClearLog
    SetNamed "DV_Status", "Running checklist..."
    StampLog "Checklist dry-run started"

    Dim failures As Collection
    Set failures = RunChecklist()

    If failures.Count = 0 Then
        SetNamed "DV_Status", "OK"
        StampLog "Checklist passed"
    Else
        WriteFailures failures
        SetNamed "DV_Status", "Failed: " & failures.Count & " checklist issue(s)"
    End If
End Sub

Public Sub Btn_UploadAll()
    On Error GoTo Failed
    ClearLog
    SetNamed "DV_Status", "Starting upload..."
    StampLog "Upload started"

    ' --- 1. Checklist (abort on failure) ---
    Dim failures As Collection
    Set failures = RunChecklist()
    If failures.Count > 0 Then
        WriteFailures failures
        SetNamed "DV_Status", "Failed: checklist has " & failures.Count & " issue(s)"
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
        Exit Sub
    End If
    If Not fso.FolderExists(locPath) Then
        SetNamed "DV_Status", "Failed: Local path not accessible: " & locPath
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
            RestoreSheetFromTemplate ThisWorkbook, sheetName, i
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

Private Sub RestoreSheetFromTemplate(liveWb As Workbook, sheetName As String, idx As Long)
    ' Replace one live data sheet with its pristine internal snapshot, preserving
    ' tab position. Transactional: the original is removed only AFTER the
    ' replacement is in place, so a mid-failure never loses the sheet.
    Dim tplName As String
    tplName = TemplateSheetName(idx)
    If Not WorkbookHasSheetIn(liveWb, tplName) Then
        StampLog "  '" & sheetName & "': no internal snapshot (" & tplName & _
                 ") - not reset. Create snapshots first: SeedBlankTemplatesFromFile" & _
                 " (from a blank copy) or RebuildBlankTemplates (on a blank workbook)."
        Exit Sub
    End If

    Dim orig As Worksheet, tpl As Worksheet
    Set orig = liveWb.Worksheets(sheetName)
    Set tpl = liveWb.Worksheets(tplName)

    ' Sanity: the snapshot's A1 title must match the live sheet's, else the
    ' positional mapping is stale - skip rather than restore the wrong template.
    If Len(CStr(tpl.Cells(1, 1).Value)) > 0 And Len(CStr(orig.Cells(1, 1).Value)) > 0 Then
        If StrComp(CStr(tpl.Cells(1, 1).Value), CStr(orig.Cells(1, 1).Value), vbTextCompare) <> 0 Then
            StampLog "  '" & sheetName & "': snapshot title mismatch (" & tplName & _
                     ") - not reset. Re-run RebuildBlankTemplates."
            Exit Sub
        End If
    End If

    Dim savedAlerts As Boolean
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Fail

    ' The snapshot is very hidden; make it visible so the copy becomes the active
    ' sheet, then re-hide the snapshot. Copy lands just before the original.
    Dim savedVis As XlSheetVisibility
    savedVis = tpl.Visible
    tpl.Visible = xlSheetVisible
    tpl.Copy Before:=orig
    Dim fresh As Worksheet
    Set fresh = liveWb.ActiveSheet
    tpl.Visible = savedVis

    orig.Delete
    fresh.Name = sheetName
    fresh.Visible = xlSheetVisible   ' ApplySheetVisibility sets the real state

    Application.DisplayAlerts = savedAlerts
    Exit Sub

Fail:
    On Error Resume Next
    tpl.Visible = xlSheetVeryHidden
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

    Dim made As Long, tplName As String, snap As Worksheet, src As Worksheet
    Dim srcVis As XlSheetVisibility
    made = 0
    For i = LBound(arr) To UBound(arr)
        sheetName = CStr(arr(i))
        If SheetExists(sheetName) Then
            tplName = TemplateSheetName(i)
            If WorkbookHasSheetIn(ThisWorkbook, tplName) Then
                ThisWorkbook.Worksheets(tplName).Delete
            End If
            ' Make the source visible so its copy becomes the active sheet (most
            ' canonical sheets are hidden per DV_TestSelection); restore after.
            Set src = ThisWorkbook.Worksheets(sheetName)
            srcVis = src.Visible
            src.Visible = xlSheetVisible
            src.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
            Set snap = ThisWorkbook.ActiveSheet
            src.Visible = srcVis
            snap.Name = tplName
            made = made + 1
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

    StampLog "RebuildBlankTemplates: created " & made & " internal snapshot(s)."
    MsgBox "Created " & made & " blank-template snapshot(s).", vbInformation, _
           "Rebuild Blank Templates"

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

    Dim made As Long, tplName As String, snap As Worksheet
    made = 0
    For i = LBound(arr) To UBound(arr)
        sheetName = CStr(arr(i))
        If WorkbookHasSheetIn(src, sheetName) Then
            tplName = TemplateSheetName(i)
            If WorkbookHasSheetIn(ThisWorkbook, tplName) Then
                ThisWorkbook.Worksheets(tplName).Delete
            End If
            src.Worksheets(sheetName).Copy _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
            Set snap = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
            snap.Name = tplName
            made = made + 1
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

    StampLog "SeedBlankTemplatesFromFile: created " & made & " snapshot(s) from " & srcPath
    MsgBox "Created " & made & " blank-template snapshot(s) from:" & vbLf & srcPath & _
           vbLf & vbLf & "Your data sheets were not touched. Save the workbook now.", _
           vbInformation, "Seed Blank Templates"

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
    If Len(Trim$(GetNamed("DV_FileName"))) = 0 Then failures.Add "DV_FileName is empty"
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
End Sub

Public Sub Btn_PickLocalFolder()
    PickFolderInto "DV_LocalPath", "Choose the local data folder"
End Sub



Private Sub PickFolderInto(ByVal namedRange As String, ByVal title As String)
    Dim startPath As String, chosen As String
    startPath = GetNamed(namedRange)
    chosen = BrowseForFolderOwned(title, startPath)
    If Len(chosen) > 0 Then SetNamed namedRange, chosen
End Sub

Private Function BrowseForFolderOwned(ByVal title As String, ByVal startPath As String) As String
    Dim shellApp As Object, fldr As Object
    Const BIF_RETURNONLYFSDIRS As Long = &H1
    Const BIF_NEWDIALOGSTYLE As Long = &H40
    Dim flgs As Long: flgs = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    Set shellApp = CreateObject("Shell.Application")
    If Len(startPath) > 0 Then
        Set fldr = shellApp.BrowseForFolder(Application.hWnd, title, flgs, startPath)
    Else
        Set fldr = shellApp.BrowseForFolder(Application.hWnd, title, flgs)
    End If
    If Not fldr Is Nothing Then BrowseForFolderOwned = fldr.Self.path
End Function



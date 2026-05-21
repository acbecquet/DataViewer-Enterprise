Attribute VB_Name = "DataViewerUpload"
Option Explicit

' ============================================================================
' DataViewer Enterprise - upload sidecar module.
'
' Drop this file into the macro-enabled template (.xlsm) via VBE:
'   Alt+F11 -> File -> Import File -> select DataViewerUpload.bas
'
' Wire the two buttons on the "DataViewer Upload" sheet:
'   Button 1 -> Btn_UploadAll
'   Button 2 -> Btn_DryRunChecklist
'
' Named ranges that must already exist on the workbook:
'   DV_FileName      base filename (no extension)
'   DV_SynologyPath  destination folder #1
'   DV_LocalPath     destination folder #2
'   DV_Status        single-cell status output
'   DV_Log           single-cell log output (multi-line text)
' ============================================================================

' Default install location. Override by adding a DV_DataViewerExe named range.
Private Const DEFAULT_DATAVIEWER_EXE As String = _
    "C:\Program Files\DataViewer Enterprise\DataViewer.exe"

Private Const COLS_PER_SAMPLE As Long = 12
Private Const MAX_SAMPLES_PER_SHEET As Long = 8
Private Const FIRST_DATA_ROW As Long = 5   ' rows 1-3 metadata, row 4 headers
Private Const TPM_MAX_PLAUSIBLE As Double = 50#

' ----------------------------------------------------------------------------
' Canonical data-sheet list (must match the layout DataViewer expects).
' Procedure / SOP sheets are excluded - they aren't sample-block sheets.
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
        "Various Oil Compatibility")
End Function

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

    ' --- 3. Copy .xlsm to both destinations ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim synDest As String, locDest As String
    synDest = AppendBackslash(synPath) & baseName & ".xlsm"
    locDest = AppendBackslash(locPath) & baseName & ".xlsm"

    If Not fso.FolderExists(synPath) Then
        SetNamed "DV_Status", "Failed: Synology path not accessible: " & synPath
        Exit Sub
    End If
    If Not fso.FolderExists(locPath) Then
        SetNamed "DV_Status", "Failed: Local path not accessible: " & locPath
        Exit Sub
    End If

    StampLog "Copy -> " & synDest
    fso.CopyFile ThisWorkbook.FullName, synDest, True
    StampLog "Copy -> " & locDest
    fso.CopyFile ThisWorkbook.FullName, locDest, True

    ' --- 4. Materialize a temp .xlsx for DataViewer ingestion ---
    ' Shape A: DataViewer's ingestion code is .xlsx-native and macro-free; the
    ' temp .xlsx avoids any macro-prompting risk from the COM reader path.
    Dim tempXlsx As String
    tempXlsx = MakeTempXlsx(fso, baseName)
    StampLog "Temp .xlsx: " & tempXlsx

    ' --- 5. Launch DataViewer; SingleInstance hands off to a running window ---
    Dim dvExe As String
    dvExe = ResolveDataViewerExe()
    If Not fso.FileExists(dvExe) Then
        SetNamed "DV_Status", "Failed: DataViewer.exe not found at " & dvExe
        Exit Sub
    End If

    Dim cmd As String
    cmd = """" & dvExe & """ """ & tempXlsx & """"
    StampLog "Shell: " & cmd
    Shell cmd, vbNormalFocus

    SetNamed "DV_Status", "OK"
    StampLog "Upload dispatched (DB write happens inside DataViewer)"
    Exit Sub

Failed:
    Application.DisplayAlerts = True
    StampLog "ERROR " & Err.Number & ": " & Err.Description
    SetNamed "DV_Status", "Failed: " & Err.Description
End Sub

' ============================================================================
' Checklist
' ============================================================================

Public Function RunChecklist() As Collection
    Dim failures As New Collection
    Dim sheetName As Variant
    Dim cs As Variant
    cs = CanonicalDataSheets()

    ' Required inputs
    If Len(Trim$(GetNamed("DV_FileName"))) = 0 Then failures.Add "DV_FileName is empty"
    If Len(Trim$(GetNamed("DV_SynologyPath"))) = 0 Then failures.Add "DV_SynologyPath is empty"
    If Len(Trim$(GetNamed("DV_LocalPath"))) = 0 Then failures.Add "DV_LocalPath is empty"

    ' Every canonical data sheet must exist
    For Each sheetName In cs
        If Not SheetExists(CStr(sheetName)) Then
            failures.Add "Missing sheet: " & sheetName
        End If
    Next

    ' Per-sheet validation (skip sheets that are absent)
    For Each sheetName In cs
        If SheetExists(CStr(sheetName)) Then
            ValidateSheet ThisWorkbook.Worksheets(CStr(sheetName)), failures
        End If
    Next

    Set RunChecklist = failures
End Function

Private Sub ValidateSheet(ws As Worksheet, failures As Collection)
    Dim sampleIdx As Long, startCol As Long
    Dim sampleID As String, heatTech As String
    Dim seenIDs As Object
    Set seenIDs = CreateObject("Scripting.Dictionary")

    For sampleIdx = 0 To MAX_SAMPLES_PER_SHEET - 1
        ' 1-based: block-start column for sample N is N*12 + 1
        startCol = sampleIdx * COLS_PER_SAMPLE + 1

        ' Sample ID at row 1, offset +5; heating tech at row 1, offset +7.
        ' Matches the "new" December-2025 template layout DataViewer reads.
        sampleID = Trim$(SafeString(ws.Cells(1, startCol + 5).Value))
        heatTech = Trim$(SafeString(ws.Cells(1, startCol + 7).Value))

        If Len(sampleID) > 0 Then
            If seenIDs.Exists(sampleID) Then
                failures.Add ws.Name & ": duplicate sample ID '" & sampleID & "'"
            Else
                seenIDs.Add sampleID, True
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

        puffs = SafeDouble(ws.Cells(row, startCol).Value)
        bWeight = SafeDouble(ws.Cells(row, startCol + 1).Value)
        aWeight = SafeDouble(ws.Cells(row, startCol + 2).Value)
        tpm = SafeDouble(ws.Cells(row, startCol + 8).Value)

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
    IsEmptyCell = (IsEmpty(c.Value) Or Len(Trim$(SafeString(c.Value))) = 0)
End Function

' ============================================================================
' Named-range helpers
' ============================================================================

Private Function GetNamed(rangeName As String) As String
    On Error Resume Next
    GetNamed = SafeString(ThisWorkbook.Names(rangeName).RefersToRange.Value)
    On Error GoTo 0
End Function

Private Sub SetNamed(rangeName As String, value As String)
    On Error Resume Next
    ThisWorkbook.Names(rangeName).RefersToRange.Value = value
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

Private Function MakeTempXlsx(fso As Object, baseName As String) As String
    Dim tempDir As String, stagedXlsm As String, outXlsx As String
    tempDir = Environ$("TEMP")
    stagedXlsm = tempDir & "\dvupload_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                 baseName & ".xlsm"
    outXlsx = tempDir & "\dvupload_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & _
              baseName & ".xlsx"

    ' Copy the current .xlsm to a staging path, open it read-only, SaveAs as
    ' .xlsx (FileFormat 51 = xlOpenXMLWorkbook), close, drop the staged .xlsm.
    fso.CopyFile ThisWorkbook.FullName, stagedXlsm, True

    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(Filename:=stagedXlsm, ReadOnly:=True, _
                                        UpdateLinks:=0)
    wb.SaveAs Filename:=outXlsx, FileFormat:=51
    wb.Close SaveChanges:=False

    Application.EnableEvents = True
    Application.DisplayAlerts = True

    On Error Resume Next
    fso.DeleteFile stagedXlsm, True
    On Error GoTo 0

    MakeTempXlsx = outXlsx
End Function

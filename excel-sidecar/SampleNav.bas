Attribute VB_Name = "SampleNav"
Option Explicit

' ============================================================================
' Sample Navigation - hotkey + ribbon navigation across 12-column sample
' blocks in the TPM testing template workbook.
'
' Sample columns: C, O, AA, AM, AY, BK, BW, CI (cols 3, 15, 27, 39, 51, 63,
' 75, 87) - 8 samples max per sheet, every 12 columns starting at column C.
'
' Hotkeys (wired in ThisWorkbook.Workbook_Open):
'   Ctrl+Shift+. -> JumpRight12 (next sample)
'   Ctrl+Shift+, -> JumpLeft12  (prev sample)
'
' Ribbon buttons (customUI14.xml, "Sample Navigation" group on the
' "TPM Testing" tab):
'   First | Prev Sample | Next Sample | Last
'
' Each main Sub is no-arg so Application.OnKey can call it. A thin
' _Ribbon(control As Object) wrapper exists per macro to satisfy the
' customUI onAction signature requirement.
' ============================================================================

Private Const COLS_PER_SAMPLE As Long = 12
Private Const FIRST_SAMPLE_COL As Long = 3       ' column C
Private Const MAX_SAMPLES As Long = 8
Private Const HEADER_ROW As Long = 4
' Last possible sample col = 3 + (8-1)*12 = 87 (column CI)
Private Const LAST_SAMPLE_COL As Long = _
    FIRST_SAMPLE_COL + (MAX_SAMPLES - 1) * COLS_PER_SAMPLE

' ----------------------------------------------------------------------------
' Public no-arg macros - direct OnKey targets, also called by ribbon wrappers
' ----------------------------------------------------------------------------

Public Sub JumpRight12()
    Dim r As Range
    Set r = SafeActiveCell()
    If r Is Nothing Then Exit Sub

    Dim targetCol As Long
    targetCol = r.Column + COLS_PER_SAMPLE
    If targetCol > LAST_SAMPLE_COL Then targetCol = LAST_SAMPLE_COL

    On Error Resume Next
    ActiveSheet.Cells(r.Row, targetCol).Select
End Sub

Public Sub JumpLeft12()
    Dim r As Range
    Set r = SafeActiveCell()
    If r Is Nothing Then Exit Sub

    Dim targetCol As Long
    targetCol = r.Column - COLS_PER_SAMPLE
    If targetCol < FIRST_SAMPLE_COL Then targetCol = FIRST_SAMPLE_COL

    On Error Resume Next
    ActiveSheet.Cells(r.Row, targetCol).Select
End Sub

Public Sub JumpFirstSample()
    Dim r As Range
    Set r = SafeActiveCell()
    If r Is Nothing Then Exit Sub

    On Error Resume Next
    ActiveSheet.Cells(r.Row, FIRST_SAMPLE_COL).Select
End Sub

Public Sub JumpLastSample()
    Dim r As Range
    Set r = SafeActiveCell()
    If r Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim sampleIdx As Long
    Dim col As Long
    Dim lastCol As Long
    Dim v As Variant
    lastCol = 0

    For sampleIdx = MAX_SAMPLES To 1 Step -1
        col = FIRST_SAMPLE_COL + (sampleIdx - 1) * COLS_PER_SAMPLE
        On Error Resume Next
        v = ws.Cells(HEADER_ROW, col).Value
        On Error GoTo 0
        If Not IsEmpty(v) Then
            If Len(Trim(CStr(v))) > 0 Then
                lastCol = col
                Exit For
            End If
        End If
    Next sampleIdx

    If lastCol = 0 Then Exit Sub   ' no samples present, do nothing

    On Error Resume Next
    ws.Cells(r.Row, lastCol).Select
End Sub

' ----------------------------------------------------------------------------
' Ribbon-callable wrappers (customUI onAction requires this signature)
' ----------------------------------------------------------------------------

Public Sub JumpRight12_Ribbon(control As Object)
    JumpRight12
End Sub

Public Sub JumpLeft12_Ribbon(control As Object)
    JumpLeft12
End Sub

Public Sub JumpFirstSample_Ribbon(control As Object)
    JumpFirstSample
End Sub

Public Sub JumpLastSample_Ribbon(control As Object)
    JumpLastSample
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

Private Function SafeActiveCell() As Range
    ' Returns Nothing on chart sheet, protected view, or empty workbook.
    On Error Resume Next
    If ActiveSheet Is Nothing Then Exit Function
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
    If ActiveCell Is Nothing Then Exit Function
    Set SafeActiveCell = ActiveCell
End Function

Attribute VB_Name = "TestingTools"
'==========================================================
' TPM Testing - Sample Block Macros
' Module: TestingTools
'==========================================================
Option Explicit

Private Const ALLOWED_SHEETS As String = "|Lifetime Test|User Test Simulation|Long Puff Lifetime Test|Rapid Puff Lifetime Test|Intense Test|Big Headspace Serial Test|Negative Pressure Test|Temperature Cycling Test #2|Viscosity Compatibility|Various Oil Compatibility|Custom Test Template|"
Private Const MASTER_SHEET As String = "_Template_Master"
Private Const BLOCK_COLS As Long = 12
Private Const BLOCK_ROWS As Long = 115

Private Function IsAllowedSheet(ws As Worksheet) As Boolean
    IsAllowedSheet = InStr(1, ALLOWED_SHEETS, "|" & ws.Name & "|", vbTextCompare) > 0
End Function

Private Function MasterSheet() As Worksheet
    On Error Resume Next
    Set MasterSheet = ThisWorkbook.Worksheets(MASTER_SHEET)
    On Error GoTo 0
    If MasterSheet Is Nothing Then
        MsgBox "Hidden template sheet '" & MASTER_SHEET & "' is missing.", vbCritical, "TPM Testing"
    End If
End Function

Private Function CountBlocks(ws As Worksheet) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then CountBlocks = 0 Else CountBlocks = Int((lastCol + BLOCK_COLS - 1) / BLOCK_COLS)
End Function

Private Function BlockIndexOfColumn(col As Long) As Long
    BlockIndexOfColumn = Int((col - 1) / BLOCK_COLS)
End Function

Private Function ColLetter(col As Long) As String
    Dim s As String, n As Long
    n = col
    Do While n > 0
        s = Chr(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Public Sub AddSample()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsAllowedSheet(ws) Then MsgBox "Add Sample can only run on a test data sheet.", vbExclamation, "TPM Testing": Exit Sub
    Dim master As Worksheet: Set master = MasterSheet()
    If master Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim nBlocks As Long: nBlocks = CountBlocks(ws)
    Dim destStartCol As Long: destStartCol = nBlocks * BLOCK_COLS + 1

    master.Range(master.Cells(1, 1), master.Cells(BLOCK_ROWS, BLOCK_COLS)).Copy
    ws.Cells(1, destStartCol).PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    Dim i As Long
    For i = 0 To BLOCK_COLS - 1
        ws.Columns(destStartCol + i).ColumnWidth = master.Columns(i + 1).ColumnWidth
    Next i

    Dim titleCell As Range: Set titleCell = ws.Cells(1, destStartCol)
    If Len(titleCell.value) > 0 Then titleCell.value = titleCell.value & " " & (nBlocks + 1)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate

    ws.Cells(1, destStartCol).Select
    MsgBox "Added sample block " & (nBlocks + 1) & " in columns " & ColLetter(destStartCol) & ":" & ColLetter(destStartCol + BLOCK_COLS - 1) & ".", vbInformation, "TPM Testing"
End Sub

Public Sub RemoveSample()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsAllowedSheet(ws) Then MsgBox "Remove Sample can only run on a test data sheet.", vbExclamation, "TPM Testing": Exit Sub

    Dim nBlocks As Long: nBlocks = CountBlocks(ws)
    If nBlocks <= 1 Then MsgBox "Cannot remove the last sample block.", vbExclamation, "TPM Testing": Exit Sub

    Dim startCol As Long: startCol = (nBlocks - 1) * BLOCK_COLS + 1
    Dim endCol As Long: endCol = nBlocks * BLOCK_COLS
    Dim titleVal As String: titleVal = CStr(ws.Cells(1, startCol).value)

    If MsgBox("Remove rightmost sample block?" & vbCrLf & vbCrLf & _
              "Sheet: " & ws.Name & vbCrLf & _
              "Columns: " & ColLetter(startCol) & ":" & ColLetter(endCol) & vbCrLf & _
              "Title: " & titleVal & vbCrLf & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbExclamation + vbDefaultButton2, "TPM Testing - Confirm Remove") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    ws.Range(ws.Columns(startCol), ws.Columns(endCol)).Delete Shift:=xlToLeft
    Application.ScreenUpdating = True
    MsgBox "Sample block removed.", vbInformation, "TPM Testing"
End Sub

Public Sub ResetEquations()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsAllowedSheet(ws) Then MsgBox "Reset can only run on a test data sheet.", vbExclamation, "TPM Testing": Exit Sub

    Dim master As Worksheet: Set master = MasterSheet()
    If master Is Nothing Then Exit Sub

    Dim blockIdx As Long: blockIdx = BlockIndexOfColumn(ActiveCell.Column)
    Dim startCol As Long: startCol = blockIdx * BLOCK_COLS + 1
    Dim endCol As Long: endCol = startCol + BLOCK_COLS - 1

    If CountBlocks(ws) <= blockIdx Then MsgBox "Click inside a sample block first.", vbExclamation, "TPM Testing": Exit Sub

    If MsgBox("Reset formulas in sample block " & (blockIdx + 1) & " (cols " & ColLetter(startCol) & ":" & ColLetter(endCol) & ")?" & vbCrLf & vbCrLf & _
              "Your entered data (puffs, weights, draw pressure, resistance, smell, clog, notes, voltage, oil mass) will NOT be touched." & vbCrLf & vbCrLf & _
              "Only the calculated formulas will be restored.", vbYesNo + vbQuestion, "TPM Testing - Confirm Reset") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Formula cell ranges within a block: (relCol, rowStart, rowEnd)
    Dim specs As Variant
    specs = Array(Array(1, 6, 115), Array(2, 6, 115), Array(6, 2, 2), Array(9, 3, 3), Array(9, 5, 115), _
                  Array(10, 5, 115), Array(11, 6, 115), Array(12, 2, 2), Array(12, 3, 3), Array(12, 5, 115))

    Dim spec As Variant, relCol As Long, r1 As Long, r2 As Long
    For Each spec In specs
        relCol = spec(0): r1 = spec(1): r2 = spec(2)
        ws.Range(ws.Cells(r1, startCol + relCol - 1), ws.Cells(r2, startCol + relCol - 1)).Formula = _
            master.Range(master.Cells(r1, relCol), master.Cells(r2, relCol)).Formula
    Next spec

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate
    MsgBox "Formulas reset in block " & (blockIdx + 1) & ". Your data is intact.", vbInformation, "TPM Testing"
End Sub

' Ribbon callbacks
Public Sub Ribbon_AddSample(control As IRibbonControl): AddSample: End Sub
Public Sub Ribbon_RemoveSample(control As IRibbonControl): RemoveSample: End Sub
Public Sub Ribbon_ResetEquations(control As IRibbonControl): ResetEquations: End Sub



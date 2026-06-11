Attribute VB_Name = "TestingTools"
'==========================================================
' TPM Testing - Sample Block Macros
' Module: TestingTools
'==========================================================
Option Explicit

Private Const MASTER_SHEET As String = "_Template_Master"
Private Const BLOCK_COLS As Long = 12
Private Const BLOCK_ROWS As Long = 115

Private Function IsAllowedSheet(ws As Worksheet) As Boolean
    ' A test data sheet is any sheet whose first sample block has the canonical
    ' "puffs" header at row 4, column 1 (the same marker TryPuffStepPicker uses).
    ' This auto-covers every block-layout test sheet and excludes step-checklist
    ' sheets (Temperature Cycling Test #1), prose sheets (Test SOP's), and
    ' "- Review" archive copies -- no hard-coded name list to drift.
    On Error GoTo NotAllowed
    If ws Is Nothing Then GoTo NotAllowed
    If InStr(1, ws.Name, " - Review", vbTextCompare) > 0 Then GoTo NotAllowed
    If Left$(ws.Name, 1) = "_" Then GoTo NotAllowed
    IsAllowedSheet = (LCase$(Trim$(CStr(ws.Cells(4, 1).Value))) = "puffs")
    Exit Function
NotAllowed:
    IsAllowedSheet = False
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
    ' Count contiguous 12-col blocks by their row-4 "puffs" header at each
    ' block-start column (1, 13, 25, ...). Robust against stray cells beyond
    ' the last block, which the old row-4 last-used-column scan mistook for
    ' extra blocks.
    Dim n As Long, col As Long
    col = 1
    Do While col <= ws.Columns.Count
        If LCase$(Trim$(CStr(ws.Cells(4, col).Value))) = "puffs" Then
            n = n + 1
            col = col + BLOCK_COLS
        Else
            Exit Do
        End If
    Loop
    CountBlocks = n
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

    Dim nBlocks As Long: nBlocks = CountBlocks(ws)
    Dim destStartCol As Long: destStartCol = nBlocks * BLOCK_COLS + 1

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
    Exit Sub
AddFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Add Sample hit an error and stopped: " & Err.Description, vbExclamation, "TPM Testing"
End Sub

Public Sub RemoveSample()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsAllowedSheet(ws) Then MsgBox "Remove Sample can only run on a test data sheet.", vbExclamation, "TPM Testing": Exit Sub

    Dim nBlocks As Long: nBlocks = CountBlocks(ws)
    If nBlocks <= 1 Then MsgBox "Cannot remove the last sample block.", vbExclamation, "TPM Testing": Exit Sub

    Dim startCol As Long: startCol = (nBlocks - 1) * BLOCK_COLS + 1
    Dim endCol As Long: endCol = nBlocks * BLOCK_COLS
    Dim titleVal As String: titleVal = CStr(ws.Cells(1, startCol).value)

    Dim filled As Long
    filled = Application.WorksheetFunction.CountA( _
        ws.Range(ws.Cells(5, startCol), ws.Cells(BLOCK_ROWS, endCol)))

    If MsgBox("Remove rightmost sample block?" & vbCrLf & vbCrLf & _
              "Sheet: " & ws.Name & vbCrLf & _
              "Columns: " & ColLetter(startCol) & ":" & ColLetter(endCol) & vbCrLf & _
              "Title: " & titleVal & vbCrLf & _
              IIf(filled > 0, "THIS BLOCK CONTAINS " & filled & " FILLED DATA CELL(S).", _
                  "This block is empty.") & vbCrLf & vbCrLf & _
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
              "Your entered data (puffs, weights, draw pressure, resistance, smell, notes, voltage, oil mass) will NOT be touched." & vbCrLf & vbCrLf & _
              "Only the calculated formulas will be restored.", vbYesNo + vbQuestion, "TPM Testing - Confirm Reset") <> vbYes Then Exit Sub

    On Error GoTo ResetFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

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

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate
    MsgBox "Formulas reset in block " & (blockIdx + 1) & ". Your data is intact.", vbInformation, "TPM Testing"
    Exit Sub
ResetFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Reset Formulas hit an error and stopped: " & Err.Description, vbExclamation, "TPM Testing"
End Sub

Public Function TryPuffStepPicker(Sh As Object, Target As Range) As Boolean
    ' Auto-fills a sample block's puff column when the user types a step value
    ' (any positive number) or "custom" in the puffs column (each block's 1st column,
    ' header row 4 = "puffs"), rows 5..BLOCK_ROWS. Returns True if it handled the
    ' edit. CRASH-SAFE: Application.EnableEvents is ALWAYS restored at PuffDone,
    ' so a mid-edit error can never leave events disabled for the session.
    TryPuffStepPicker = False
    If Sh Is Nothing Then Exit Function
    If TypeName(Sh) <> "Worksheet" Then Exit Function
    ' Never run on archived Review copies or hidden plumbing (audit H1): they
    ' share the test-sheet layout but must not be mutated by the step picker.
    If InStr(1, Sh.Name, " - Review", vbTextCompare) > 0 Then Exit Function
    If Left$(Sh.Name, 1) = "_" Then Exit Function

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
        Case "custom": stp = "custom"
        Case Else
            ' v1.2 (ports the owner's in-workbook v1.1 fix): ANY positive number
            ' is a step. Poka-yokes, never gates -- testers may use any interval.
            If IsNumeric(v) And CDbl(v) > 0 Then
                stp = CDbl(v)
            Else
                Exit Function             ' nothing to do - let other handlers run
            End If
    End Select

    Dim savedEvents As Boolean
    savedEvents = Application.EnableEvents
    On Error GoTo PuffDone
    Application.EnableEvents = False

    If VarType(stp) = vbString Then        ' "custom"
        pick.ClearContents
        pick.Select
    Else
        ' The list dropdown would reject arbitrary steps on the next edit; the
        ' column is free-form by design now (owner decision, v1.2).
        On Error Resume Next
        Sh.Range(Sh.Cells(pick.row, c), Sh.Cells(BLOCK_ROWS, c)).Validation.Delete
        On Error GoTo PuffDone
        Dim startFill As Long
        If pick.row = 5 Then
            ' Row-5 seeding (v1.2): the first puff value also becomes the step --
            ' rows below fill with prev + value, same as typing on a later row.
            Sh.Cells(5, c).value = stp
            startFill = 6
        Else
            startFill = pick.row
        End If
        Dim colL As String
        colL = ColLetter(c)
        Dim rr As Long
        For rr = startFill To BLOCK_ROWS
            ' Str$ always renders a period decimal separator (locale-safe formula text).
            Sh.Cells(rr, c).Formula = "=" & colL & (rr - 1) & "+" & Trim$(Str$(stp))
        Next rr
    End If
    TryPuffStepPicker = True

PuffDone:
    Application.EnableEvents = savedEvents
    On Error GoTo 0
End Function

' Ribbon callbacks
Public Sub Ribbon_AddSample(control As IRibbonControl): AddSample: End Sub
Public Sub Ribbon_RemoveSample(control As IRibbonControl): RemoveSample: End Sub
Public Sub Ribbon_ResetEquations(control As IRibbonControl): ResetEquations: End Sub



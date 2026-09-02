Attribute VB_Name = "Error_Handling_Demo"
'Most macros that "break" in the real world are not wrong - they simply
'stopped halfway through and left Excel in a strange state, with screen
'updating off or alerts suppressed. The pattern below is the fix: one
'error handler, one cleanup block, and a single exit path that always
'runs whether the macro succeeded or failed.
Sub Error_Handling_Demo()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim badRows As Long
    Dim result As Double

    'Remember the settings we are about to change so we can put them back
    Dim savedScreenUpdating As Boolean
    Dim savedCalculation As XlCalculation
    savedScreenUpdating = Application.ScreenUpdating
    savedCalculation = Application.Calculation

    'From this line on, any unexpected error jumps straight to ErrorHandler
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For i = 3 To lastRow
        'EXPECTED errors are better handled with a test than with a trap.
        'A blank or text score is a normal thing to meet in real data, so
        'check for it rather than letting it raise an error.
        If Not IsNumeric(ws.Cells(i, "D").Value) Or ws.Cells(i, "D").Value = "" Then
            badRows = badRows + 1
        Else
            result = result + ws.Cells(i, "D").Value
        End If
    Next i

    'For a genuinely unexpected error, On Error Resume Next lets you carry
    'on and inspect Err yourself - but you must clear it afterwards, or a
    'stale error number will confuse the next check.
    On Error Resume Next
    Dim divided As Double
    divided = result / (lastRow - 2 - badRows)
    If Err.Number <> 0 Then
        divided = 0
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    MsgBox "Valid rows: " & (lastRow - 2 - badRows) & vbNewLine & _
           "Skipped rows: " & badRows & vbNewLine & _
           "Average: " & Format(divided, "0.00"), _
           vbInformation, "Finished cleanly"

CleanExit:
    'Everything that must happen either way lives here. Note the Exit Sub
    'below - without it, a successful run would fall through into the
    'error handler and report an error that never happened.
    Application.ScreenUpdating = savedScreenUpdating
    Application.Calculation = savedCalculation
    Exit Sub

ErrorHandler:
    'Err.Number and Err.Description tell you what went wrong. Showing the
    'line-free message is fine for training; in production you would log
    'it somewhere you can read later.
    MsgBox "Something went wrong and the macro stopped safely." & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Error_Handling_Demo"

    'Resume CleanExit sends us through the cleanup block on the way out,
    'so Excel is never left with calculation on manual
    Resume CleanExit
End Sub

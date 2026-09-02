Attribute VB_Name = "IfElse_SelectCase"
'Sample file: "Sample Files\Student Exams.xlsx"
'Turns the raw scores in column D into a letter grade and a pass/fail note.
'The same decision is written twice - once with If/ElseIf and once with
'Select Case - so you can compare the two styles side by side.
Sub IfElse_SelectCase()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstDataRow As Long
    Dim i As Long
    Dim score As Double

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    firstDataRow = 3
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    'Write headers for the two new columns
    ws.Cells(2, "H").Value = "Grade"
    ws.Cells(2, "I").Value = "Result"

    For i = firstDataRow To lastRow
        score = ws.Cells(i, "D").Value

        'STYLE 1 - If / ElseIf / Else
        'Each test is checked in order and the first true one wins, so the
        'bands must be written from highest to lowest. Reversing the order
        'would give every student an "F".
        If score >= 80 Then
            ws.Cells(i, "H").Value = "A"
        ElseIf score >= 70 Then
            ws.Cells(i, "H").Value = "B"
        ElseIf score >= 60 Then
            ws.Cells(i, "H").Value = "C"
        ElseIf score >= 50 Then
            ws.Cells(i, "H").Value = "D"
        Else
            ws.Cells(i, "H").Value = "F"
        End If

        'STYLE 2 - Select Case
        'Easier to read when you are testing one value against many ranges.
        'Case Else catches anything that did not match, which is where a
        'blank or negative score would end up.
        Select Case score
            Case Is >= 50
                ws.Cells(i, "I").Value = "Pass"
            Case 0 To 49
                ws.Cells(i, "I").Value = "Fail"
            Case Else
                ws.Cells(i, "I").Value = "Check score"
        End Select
    Next i

    MsgBox "Graded " & (lastRow - firstDataRow + 1) & " students.", _
           vbInformation, "If/Else and Select Case"
End Sub

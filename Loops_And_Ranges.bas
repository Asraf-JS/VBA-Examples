Attribute VB_Name = "Loops_And_Ranges"
'Sample file: "Sample Files\Student Exams.xlsx"
'That sheet keeps its headers in row 2, so the data starts in row 3.
'Columns: B = Student ID, C = Student Name, D = Score, E = Classroom, F = Date, G = Exam
Sub Loops_And_Ranges()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstDataRow As Long
    Dim i As Long
    Dim cell As Range
    Dim total As Double
    Dim count As Long
    Dim highestScore As Double
    Dim topStudent As String

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    firstDataRow = 3

    'STEP 1 - Find the last row that actually holds data.
    'Start at the very bottom of column B and "press Ctrl+Up" (xlUp).
    'Always do this instead of hard-coding a row number, because the
    'list will grow or shrink over time.
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    'If the sheet only contains the header row there is nothing to process
    If lastRow < firstDataRow Then
        MsgBox "No data found below the header row.", vbExclamation
        Exit Sub
    End If

    'STEP 2 - A counter loop visits one row at a time.
    'Use this style when you need values from several columns in the same row.
    For i = firstDataRow To lastRow
        total = total + ws.Cells(i, "D").Value
        count = count + 1

        'Keep track of the best score and who earned it
        If ws.Cells(i, "D").Value > highestScore Then
            highestScore = ws.Cells(i, "D").Value
            topStudent = Trim(ws.Cells(i, "C").Value)
        End If
    Next i

    'STEP 3 - A For Each loop visits every cell in a range.
    'Use this style when you only care about the cells themselves and
    'not about which row they sit on. Here we colour every failing score.
    For Each cell In ws.Range(ws.Cells(firstDataRow, "D"), ws.Cells(lastRow, "D"))
        If cell.Value < 50 Then
            cell.Interior.Color = RGB(255, 199, 206)   'light red
        Else
            cell.Interior.ColorIndex = xlNone
        End If
    Next cell

    'STEP 4 - Report what the loops found
    MsgBox "Rows processed: " & count & vbNewLine & _
           "Average score: " & Format(total / count, "0.00") & vbNewLine & _
           "Highest score: " & highestScore & " (" & topStudent & ")", _
           vbInformation, "Loops And Ranges"
End Sub

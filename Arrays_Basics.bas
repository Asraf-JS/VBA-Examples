Attribute VB_Name = "Arrays_Basics"
'Sample file: "Sample Files\Student Exams.xlsx"
'Reading and writing cells one at a time is slow because every single
'touch crosses between VBA and Excel. Pulling the whole range into an
'array, working on it in memory, then writing it back in one go crosses
'that boundary only twice. This macro does the same job both ways and
'times them so you can see the difference.
Sub Arrays_Basics()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstDataRow As Long
    Dim i As Long
    Dim data As Variant
    Dim results() As Variant
    Dim startTime As Double
    Dim slowTime As Double
    Dim fastTime As Double

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    firstDataRow = 3
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    'METHOD 1 - the slow way, one cell at a time
    startTime = Timer
    For i = firstDataRow To lastRow
        ws.Cells(i, "K").Value = ws.Cells(i, "D").Value * 1.05
    Next i
    slowTime = Timer - startTime

    'METHOD 2 - the fast way, using an array
    startTime = Timer

    'Reading a multi-cell range gives a 2-dimensional array that is
    'always 1-based: data(row, column), starting at data(1, 1).
    data = ws.Range(ws.Cells(firstDataRow, "D"), ws.Cells(lastRow, "D")).Value

    'Size the output array to match, then fill it in memory.
    'No worksheet is touched inside this loop, which is why it is quick.
    ReDim results(1 To UBound(data, 1), 1 To 1)
    For i = 1 To UBound(data, 1)
        results(i, 1) = data(i, 1) * 1.05
    Next i

    'Write the finished array back in a single operation. The destination
    'range must be exactly the same size as the array.
    ws.Range(ws.Cells(firstDataRow, "L"), ws.Cells(lastRow, "L")).Value = results
    fastTime = Timer - startTime

    ws.Cells(2, "K").Value = "Cell by cell"
    ws.Cells(2, "L").Value = "Via array"

    MsgBox "Rows processed: " & UBound(data, 1) & vbNewLine & _
           "Cell by cell: " & Format(slowTime, "0.000") & " seconds" & vbNewLine & _
           "Via array: " & Format(fastTime, "0.000") & " seconds" & vbNewLine & vbNewLine & _
           "The gap grows quickly as the row count rises.", _
           vbInformation, "Arrays Basics"
End Sub

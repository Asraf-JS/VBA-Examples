Attribute VB_Name = "Format_Report"
'Sample file: "Sample Files\Invoice List.xlsx"
'Takes the raw invoice rows and turns them into something you would be
'happy to hand to someone else: a styled header, readable number and date
'formats, borders, a totals row, frozen panes and a highlight rule for
'the largest amounts.
Sub Format_Report()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalRow As Long
    Dim rngData As Range

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data to format.", vbExclamation
        Exit Sub
    End If

    'Turning off screen updating stops the sheet flickering through every
    'change and makes the macro noticeably faster
    Application.ScreenUpdating = False

    Set rngData = ws.Range("A1:I" & lastRow)

    'HEADER ROW - dark fill, white bold text, centred
    With ws.Range("A1:I1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 84, 150)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With

    'NUMBER FORMATS - a date that reads as a date and money that lines up.
    'Formatting changes only how a value is displayed, never the value
    'stored underneath it.
    ws.Range("B2:B" & lastRow).NumberFormat = "dd-mmm-yyyy"
    ws.Range("H2:H" & lastRow).NumberFormat = "#,##0"
    ws.Range("I2:I" & lastRow).NumberFormat = "#,##0.00"

    'BORDERS around every cell in the block
    With rngData.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 180, 180)
    End With

    'BANDED ROWS make long lists easier to read across
    Dim i As Long
    For i = 2 To lastRow
        If i Mod 2 = 0 Then
            ws.Range("A" & i & ":I" & i).Interior.Color = RGB(242, 245, 250)
        End If
    Next i

    'CONDITIONAL FORMATTING - highlight amounts above the average.
    'Delete the existing rules first so repeated runs do not stack them up.
    ws.Range("I2:I" & lastRow).FormatConditions.Delete
    With ws.Range("I2:I" & lastRow).FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=AVERAGE($I$2:$I$" & lastRow & ")")
        .Font.Bold = True
        .Font.Color = RGB(0, 97, 0)
        .Interior.Color = RGB(198, 239, 206)
    End With

    'TOTALS ROW below the data
    totalRow = lastRow + 1
    ws.Cells(totalRow, "G").Value = "TOTAL"
    ws.Cells(totalRow, "H").Formula = "=SUM(H2:H" & lastRow & ")"
    ws.Cells(totalRow, "I").Formula = "=SUM(I2:I" & lastRow & ")"
    ws.Cells(totalRow, "I").NumberFormat = "#,##0.00"

    With ws.Range("G" & totalRow & ":I" & totalRow)
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
    End With

    'FINISHING TOUCHES
    ws.Columns("A:I").AutoFit

    'Freeze the header so it stays visible while scrolling. Panes are a
    'property of the window, not the sheet, so the sheet must be active.
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    Application.ScreenUpdating = True

    MsgBox "Formatted " & (lastRow - 1) & " invoice rows.", _
           vbInformation, "Format Report"
End Sub

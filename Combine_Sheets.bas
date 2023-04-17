Attribute VB_Name = "Combine_Sheets"
Sub Combinesheets()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsCombined As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastCol As Long
    
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    
    ' Create a new worksheet named "CombinedSheet"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("CombinedSheet").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsCombined = ThisWorkbook.Worksheets.Add
    wsCombined.Name = "CombinedSheet"
    
    ' Find the last row and last column in both sheets
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastCol = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    
    ' Copy the header row from Sheet1 to CombinedSheet
    ws1.Rows(1).Copy Destination:=wsCombined.Rows(1)
    
    ' Copy the data from both sheets to CombinedSheet (excluding header rows)
    ws1.Range(ws1.Cells(2, 1), ws1.Cells(lastRow1, lastCol)).Copy Destination:=wsCombined.Cells(2, 1)
    ws2.Range(ws2.Cells(2, 1), ws2.Cells(lastRow2, lastCol)).Copy Destination:=wsCombined.Cells(lastRow1 + 1, 1)
    
    ' Inform the user that the process is complete
    MsgBox "Data from Sheet1 and Sheet2 has been combined in CombinedSheet.", vbInformation, "Process Complete"
End Sub


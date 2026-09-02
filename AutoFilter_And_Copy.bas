Attribute VB_Name = "AutoFilter_And_Copy"
'Sample file: "Sample Files\Invoice List.xlsx"
'Columns: A = Invoice No, B = Date, C = Customer Name, D = Customer ID,
'         E = Customer Address 1, F = Customer State, G = Description,
'         H = Quantity, I = Amount
'Filtering with AutoFilter and copying the visible rows is far quicker
'than looping through every row and testing it yourself, and it is how
'most real-world "extract these records" macros are written.
Sub AutoFilter_And_Copy()
    Dim wsData As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim rngData As Range
    Dim rngVisible As Range
    Dim targetState As String
    Dim minAmount As Double

    Set wsData = ThisWorkbook.Worksheets("Sheet1")
    targetState = "Selangor"
    minAmount = 100

    'Clear any filter left over from a previous run, otherwise the new
    'criteria are applied on top of the old ones and you get no rows.
    If wsData.AutoFilterMode Then wsData.AutoFilterMode = False

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No invoice rows found.", vbExclamation
        Exit Sub
    End If

    'The range handed to AutoFilter must include the header row
    Set rngData = wsData.Range("A1:I" & lastRow)

    'Apply two criteria at once: a matching state AND a large enough amount.
    'Field numbers are counted from the left edge of rngData, so column F
    'is field 6 and column I is field 9.
    rngData.AutoFilter Field:=6, Criteria1:=targetState
    rngData.AutoFilter Field:=9, Criteria1:=">=" & minAmount

    'Rebuild the output sheet from scratch each run
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Filtered").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = ThisWorkbook.Worksheets.Add(After:=wsData)
    wsOut.Name = "Filtered"

    'SpecialCells(xlCellTypeVisible) returns only the rows the filter left
    'on screen. It raises an error when nothing matched, so trap that case
    'instead of letting the macro fall over.
    On Error Resume Next
    Set rngVisible = rngData.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rngVisible Is Nothing Then
        MsgBox "No invoices matched " & targetState & " with an amount of " & _
               minAmount & " or more.", vbInformation
    Else
        rngVisible.Copy Destination:=wsOut.Range("A1")
        wsOut.Columns("A:I").AutoFit
        MsgBox "Copied " & (wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row - 1) & _
               " matching invoices to the Filtered sheet.", vbInformation
    End If

    'Always leave the source sheet the way you found it
    wsData.AutoFilterMode = False
End Sub

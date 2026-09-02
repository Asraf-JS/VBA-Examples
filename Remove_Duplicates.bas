Attribute VB_Name = "Remove_Duplicates"
'Sample file: "Sample Files\Invoice List.xlsx"
'The same customer appears on several invoices. This macro builds a list
'of unique customers in two different ways so you can see the trade-off:
'   1. Range.RemoveDuplicates - one line, but it deletes rows in place
'   2. A Dictionary - more code, but it leaves the source data untouched
'      and can count how many times each value appeared
Sub Remove_Duplicates()
    Dim wsData As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim customer As String
    Dim dict As Object
    Dim key As Variant
    Dim outRow As Long

    Set wsData = ThisWorkbook.Worksheets("Sheet1")
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    'Rebuild the output sheet each run
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Unique Customers").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = ThisWorkbook.Worksheets.Add(After:=wsData)
    wsOut.Name = "Unique Customers"

    'METHOD 1 - RemoveDuplicates on a copy of the data.
    'We copy first because RemoveDuplicates permanently deletes rows, and
    'you almost never want that to happen to your original list.
    wsData.Range("C1:C" & lastRow).Copy Destination:=wsOut.Range("A1")
    wsOut.Range("A1:A" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    wsOut.Range("A1").Value = "Customer (RemoveDuplicates)"

    'METHOD 2 - a Dictionary.
    'CreateObject is used here (late binding) so the macro runs without
    'anyone having to add a reference to the Microsoft Scripting Runtime.
    'A Dictionary key can only exist once, which is what makes it a
    'natural fit for finding unique values.
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1    'vbTextCompare - treat "ABC" and "abc" as the same

    For i = 2 To lastRow
        customer = Trim(wsData.Cells(i, "C").Value)
        If customer <> "" Then
            If dict.Exists(customer) Then
                'Already seen it - just add to the running count
                dict(customer) = dict(customer) + 1
            Else
                dict.Add customer, 1
            End If
        End If
    Next i

    wsOut.Range("C1").Value = "Customer (Dictionary)"
    wsOut.Range("D1").Value = "Invoice Count"

    outRow = 2
    For Each key In dict.Keys
        wsOut.Cells(outRow, "C").Value = key
        wsOut.Cells(outRow, "D").Value = dict(key)
        outRow = outRow + 1
    Next key

    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("C1:D1").Font.Bold = True
    wsOut.Columns("A:D").AutoFit

    MsgBox "Found " & dict.Count & " unique customers across " & _
           (lastRow - 1) & " invoices.", vbInformation, "Remove Duplicates"
End Sub

Attribute VB_Name = "Loop_Through_Files"
'The natural next step after Combine_Sheets: instead of merging two sheets
'in one workbook, this walks through every workbook in a folder you pick,
'copies the data out of each one, and stacks it all into a summary sheet
'with a column recording which file each row came from.
Sub Loop_Through_Files()
    Dim folderPath As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRowSource As Long
    Dim lastCol As Long
    Dim nextRow As Long
    Dim filesProcessed As Long
    Dim headerWritten As Boolean

    'Let the user choose the folder rather than hard-coding a path
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder containing the workbooks to combine"
        If .Show <> -1 Then Exit Sub        '-1 means the user pressed OK
        folderPath = .SelectedItems(1)
    End With

    'Dir needs a trailing backslash before the file mask
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    'Rebuild the summary sheet each run
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsSummary = ThisWorkbook.Worksheets.Add
    wsSummary.Name = "Summary"
    nextRow = 1

    Application.ScreenUpdating = False

    'Dir returns the first matching file name, then each later call with
    'no arguments returns the next one. An empty string means we are done.
    'Never open or close workbooks inside a Dir loop without storing the
    'names first if you also change the folder - here we only read.
    fileName = Dir(folderPath & "*.xls*")

    Do While fileName <> ""
        'Skip this workbook if it happens to live in the same folder
        If fileName <> ThisWorkbook.Name Then

            Set wbSource = Workbooks.Open(fileName:=folderPath & fileName, ReadOnly:=True)
            Set wsSource = wbSource.Worksheets(1)

            lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
            lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

            If lastRowSource >= 2 Then
                'Copy the header only once, from the first file
                If Not headerWritten Then
                    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastCol)).Copy _
                        Destination:=wsSummary.Cells(1, 1)
                    wsSummary.Cells(1, lastCol + 1).Value = "Source File"
                    headerWritten = True
                    nextRow = 2
                End If

                'Copy the data rows, skipping that file's own header
                wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRowSource, lastCol)).Copy _
                    Destination:=wsSummary.Cells(nextRow, 1)

                'Stamp every copied row with the file it came from, so the
                'combined data can still be traced back to its origin
                wsSummary.Range(wsSummary.Cells(nextRow, lastCol + 1), _
                                wsSummary.Cells(nextRow + lastRowSource - 2, lastCol + 1)).Value = fileName

                nextRow = nextRow + lastRowSource - 1
            End If

            'Close without saving - we opened it read-only and changed nothing
            wbSource.Close SaveChanges:=False
            filesProcessed = filesProcessed + 1
        End If

        fileName = Dir     'move to the next file
    Loop

    wsSummary.Rows(1).Font.Bold = True
    wsSummary.Columns.AutoFit
    Application.ScreenUpdating = True

    If filesProcessed = 0 Then
        MsgBox "No Excel files were found in that folder.", vbExclamation
    Else
        MsgBox "Combined " & filesProcessed & " files into " & _
               (nextRow - 2) & " rows on the Summary sheet.", _
               vbInformation, "Loop Through Files"
    End If
End Sub

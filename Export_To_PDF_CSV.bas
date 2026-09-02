Attribute VB_Name = "Export_To_PDF_CSV"
'Two jobs people ask for constantly: "email me this as a PDF" and "give
'me a CSV I can load somewhere else". Both are shown here, saving into
'the same folder as the workbook, with the date in the file name so a
'second run does not silently overwrite yesterday's copy.
Sub Export_To_PDF_CSV()
    Dim ws As Worksheet
    Dim wbTemp As Workbook
    Dim savePath As String
    Dim pdfName As String
    Dim csvName As String
    Dim lastRow As Long
    Dim stamp As String

    Set ws = ThisWorkbook.Worksheets("Sheet1")

    'A workbook that has never been saved has no path to export into
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save this workbook first so the macro knows where " & _
               "to put the exported files.", vbExclamation
        Exit Sub
    End If

    savePath = ThisWorkbook.Path & "\"
    stamp = Format(Date, "yyyy-mm-dd")
    pdfName = savePath & "Invoice Report " & stamp & ".pdf"
    csvName = savePath & "Invoice Data " & stamp & ".csv"

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    'PART 1 - export to PDF.
    'Set the print area and fit everything onto one page wide, otherwise a
    'wide sheet splits across pages in a way nobody wants to read.
    With ws.PageSetup
        .PrintArea = "A1:I" & lastRow
        .Orientation = xlLandscape
        .Zoom = False               'Zoom must be False for FitToPages to apply
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"   'repeat the header on every page
    End With

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        OpenAfterPublish:=False

    'PART 2 - export to CSV.
    'A CSV file can only hold one sheet, so copying the sheet into a new
    'workbook and saving that keeps the original workbook untouched.
    'Saving ThisWorkbook directly as CSV would throw away every other sheet.
    ws.Copy                                  'no argument = copy to a new workbook
    Set wbTemp = ActiveWorkbook

    Application.DisplayAlerts = False        'suppress the "features not compatible" prompt
    wbTemp.SaveAs fileName:=csvName, FileFormat:=xlCSV, CreateBackup:=False
    wbTemp.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "Exported " & (lastRow - 1) & " rows to:" & vbNewLine & vbNewLine & _
           pdfName & vbNewLine & csvName, _
           vbInformation, "Export Complete"

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    MsgBox "Export failed." & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & _
           "The usual cause is the PDF or CSV already being open in " & _
           "another program, which locks the file.", _
           vbCritical, "Export_To_PDF_CSV"
    Resume CleanExit
End Sub

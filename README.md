# VBA-Examples
This Github repository is a collection of VBA (Visual Basic for Applications) code examples and snippets for Excel and other Microsoft Office applications. The examples cover a range of use cases and functions, from automating repetitive tasks to data analysis and reporting.

**extractNumbers**

In this code, we're using a For loop to iterate through each character in the input string (inputString). We use the IsNumeric function to check if each character is a number. If the character is a number, we append it to the outputString variable. Finally, we write the outputString to cell B1.

Note that this code assumes that the input string contains only alphanumeric characters and numeric digits. If the input string contains other types of characters, such as symbols or special characters, those characters will be excluded from the output string.

**extractCityAndState**

Note that this code assumes that the state keyword is present in the address and that the keyword is spelled correctly. If the address does not contain a state keyword or the keyword is spelled incorrectly, the city and state variables will be left blank. Additionally, there may be other approaches to identifying the state in an address, depending on the specific structure and content of the address data.

**changeTextCase**

If the input text is in all caps, we use the LCase function to convert it to all lowercase. If the input text is in title case or mixed case, we use the StrConv function with the vbProperCase argument to convert it to title case. If the input text is already in all uppercase, we leave it unchanged.

**SimpleAI**

In this example, the SimpleAI macro implements a simple decision tree algorithm that predicts whether or not someone should play golf based on the humidity level and wind conditions. The decision tree is represented as a two-dimensional array, where each row represents a node in the tree and the columns represent the decision features and outcomes.

**Combine_Sheets**

The VBA code provided combines data from two worksheets, Sheet1 and Sheet2, into a new worksheet called "CombinedSheet." You can find the files used in Sample Files folder called "Combined Sheets"

## Training Examples

The modules below are arranged in learning order, from the basics through to the patterns you meet in day-to-day work. Each one runs against a workbook in the Sample Files folder, so you can open the file, paste the module into the VBA editor and run it straight away.

Note that "Student Exams.xlsx" keeps its headers in row 2 with the data starting in row 3, while "Invoice List.xlsx" uses a normal header in row 1. The macros are written around those actual layouts, so if you point them at your own data, check the starting row first.

### Fundamentals

**Loops_And_Ranges**

Covers the three things nearly every macro needs: finding the last row that actually contains data with `End(xlUp)` instead of guessing a row number, walking the sheet a row at a time with a counter loop, and visiting cells with `For Each`. It reports the average and highest score, and shades any failing score. Uses "Student Exams.xlsx".

**IfElse_SelectCase**

Writes the same grading decision twice, once with `If`/`ElseIf` and once with `Select Case`, so you can compare them. The If version shows why the bands have to be tested from highest to lowest, and the Select Case version shows how `Case Else` catches the blank or unexpected values that real data always contains. Uses "Student Exams.xlsx".

**Arrays_Basics**

Reads a range into a two-dimensional array, changes the values in memory, and writes the whole thing back in a single operation. It runs the same calculation cell-by-cell first and times both, which shows why array processing is the usual answer when a macro feels slow. Uses "Student Exams.xlsx".

### Everyday tasks

**AutoFilter_And_Copy**

Applies two AutoFilter criteria at once (a matching state and a minimum amount), copies only the visible rows to a new sheet, and clears the filter afterwards. Also shows how to handle the case where nothing matched, since `SpecialCells(xlCellTypeVisible)` raises an error rather than returning an empty range. Uses "Invoice List.xlsx".

**Remove_Duplicates**

Builds a list of unique customers two ways: `Range.RemoveDuplicates`, which is a single line but permanently deletes rows, and a Dictionary, which leaves the source data alone and can also count how many times each value appeared. The Dictionary is created with `CreateObject` so no library reference has to be added first. Uses "Invoice List.xlsx".

**Format_Report**

Turns raw rows into something presentable: a styled header, date and currency number formats, borders, banded rows, a conditional format highlighting above-average amounts, a totals row with live SUM formulas, and frozen panes. Uses "Invoice List.xlsx".

### Going further

**Error_Handling_Demo**

The `On Error GoTo` pattern with a single cleanup block, so the macro cannot leave Excel with screen updating off or calculation stuck on manual. Shows the difference between trapping an unexpected error and simply testing for an expected one, why `Exit Sub` has to sit before the handler, and how `Resume` sends a failed run back through the cleanup on its way out.

**Loop_Through_Files**

The natural sequel to Combine_Sheets. Asks the user to pick a folder, then uses `Dir` to walk every workbook in it, copying the data from each into one Summary sheet and stamping each row with the file it came from. Opens the source files read-only and closes them without saving.

**Export_To_PDF_CSV**

Exports a sheet to PDF with `ExportAsFixedFormat`, using PageSetup to fit it to one page wide and repeat the header row on every page. Then exports the same data to CSV by copying the sheet into a new workbook first, because saving the workbook itself as CSV would discard every other sheet. File names carry the date so a second run does not overwrite the first.

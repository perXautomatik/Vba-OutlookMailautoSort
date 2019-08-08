Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    'do whatever you need
    ws.Cells(1, 1) = 1 'this sets cell A1 of each sheet to "1"
Next

starting_ws.Activate 'activate the worksheet that was originally active




Sub dantheramBatch()
Dim wb As Workbook, sh As Worksheet, fName As String, fPath As String, bSh As Worksheet
Set bSh = ThisWorkbook.Sheets(1) 'Edit sheet name
fPath = "C:\Temp\myXLdocs"
If Right(fPath, 1) <> "\" Then fPath = fPath & "\"
fName = Dir(fPath & "*.xl*") 'Assumes all Excel workbooks will be evaluated.
    Do
        Set wb = Workbooks.Open(fPath & fName)
        On Error GoTo SKIP:
        Set sh = wb.Sheets("ABC")
        On Error GoTo 0
        sh.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count) 'Assumes code will run from merged workbook.
        wb.Close False
        fName = Dir
SKIP:
        If Err.Number = 9 Then
            MsgBox "Sheet name not found in file " & wb.Name & "."
            Err.Clear
        Else
            MsgBox Err.Number & ":  " & Err.Description
            Err.Clear
        End If
    Loop While fName <> ""
ThisWorkbook.SaveAs "C:\SomeDir\ABCmerged.csv", 6  'File format for xlCSV
End Sub
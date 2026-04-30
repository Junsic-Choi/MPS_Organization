On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("DEBUG_ROWS.csv", True, True)

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(4)

For r = 7 To 50
    outFile.WriteLine "Row " & r & ": " & ws.Cells(r, 4).Value
Next

outFile.Close
workbook.Close False
excel.Quit

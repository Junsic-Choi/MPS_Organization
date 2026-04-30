On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("header_check.txt", True)

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(4)

cols = Array(9, 13, 18, 23, 29, 35)
For Each c In cols
    logFile.WriteLine c & ": " & ws.Cells(3, c).Value
Next

workbook.Close False
excel.Quit
logFile.Close

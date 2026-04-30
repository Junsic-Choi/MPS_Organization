On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("row_count.txt", True)
Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws2 = workbook.Sheets(2)
Set ws4 = workbook.Sheets(4)
logFile.WriteLine "Sheet 2 UsedRows: " & ws2.UsedRange.Rows.Count
logFile.WriteLine "MPS Sheet 4 UsedRows: " & ws4.UsedRange.Rows.Count
workbook.Close False
excel.Quit
logFile.Close

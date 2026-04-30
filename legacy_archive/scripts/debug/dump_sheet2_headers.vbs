On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("sheet2_headers_v2.txt", True)
Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(2)

logFile.WriteLine "Row 3:"
For c = 1 To 50
    logFile.Write ws.Cells(3, c).Value & " | "
Next
logFile.WriteLine ""

logFile.WriteLine "Row 4:"
For c = 1 To 50
    logFile.Write ws.Cells(4, c).Value & " | "
Next
logFile.WriteLine ""

workbook.Close False
excel.Quit
logFile.Close

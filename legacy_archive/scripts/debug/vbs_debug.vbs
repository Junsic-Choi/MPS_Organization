On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("vbs_debug.txt", True)

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")

logFile.WriteLine "Workbook opened. Sheets count: " & workbook.Sheets.Count

Set wsMeta = workbook.Sheets(2)
logFile.WriteLine "Meta Sheet: " & wsMeta.Name & " (Rows: " & wsMeta.UsedRange.Rows.Count & ")"

Set wsMps = workbook.Sheets(4)
logFile.WriteLine "MPS Sheet: " & wsMps.Name & " (Rows: " & wsMps.UsedRange.Rows.Count & ")"

' Test reading one cell from Meta
logFile.WriteLine "Meta(7,3): " & wsMeta.Cells(7, 3).Value

' Test reading one cell from MPS
logFile.WriteLine "MPS(7,4): " & wsMps.Cells(7, 4).Value

workbook.Close False
excel.Quit
logFile.Close

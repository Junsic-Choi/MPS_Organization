On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("col15_sum.txt", True)
Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(2)

sum15 = 0
For r = 7 To 1000
    val = ws.Cells(r, 15).Value
    If IsNumeric(val) Then
        sum15 = sum15 + val
    End If
Next

logFile.WriteLine "Col 15 Sum: " & sum15
workbook.Close False
excel.Quit
logFile.Close

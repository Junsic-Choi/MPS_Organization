On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("col5_sum.txt", True)
Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(2)

sum5 = 0
For r = 7 To 1000
    val = ws.Cells(r, 5).Value
    If IsNumeric(val) Then
        sum5 = sum5 + val
    End If
Next

logFile.WriteLine "Col 5 Sum: " & sum5
workbook.Close False
excel.Quit
logFile.Close

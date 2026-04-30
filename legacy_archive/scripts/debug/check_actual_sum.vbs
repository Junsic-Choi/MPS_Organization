On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("true_sum.txt", True)

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(4)

tCols = Array(9, 13, 18, 23, 29, 35)
total = 0
For r = 7 To 1600
    For i = 0 To 5
        v = ws.Cells(r, tCols(i)).Value
        If IsNumeric(v) Then
            total = total + CDbl(v)
        End If
    Next
Next

logFile.WriteLine total
workbook.Close False
excel.Quit
logFile.Close

On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("sheet4_sum.txt", True)
Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(4)

totalSum = 0
targetCols = Array(9, 13, 18, 23, 29, 35)
For r = 7 To 1500
    For Each c In targetCols
        val = ws.Cells(r, c).Value
        If IsNumeric(val) Then
            totalSum = totalSum + val
        End If
    Next
Next

logFile.WriteLine "Sheet 4 Total Sum: " & totalSum
workbook.Close False
excel.Quit
logFile.Close

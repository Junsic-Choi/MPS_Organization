On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("scan_qtys.txt", True)

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set wsMps = workbook.Sheets(4)

tCols = Array(9, 13, 18, 23, 29, 35)
foundCount = 0

For r = 1 To 1000
    rowText = ""
    hadData = False
    For i = 0 To 5
        v = wsMps.Cells(r, tCols(i)).Value
        If IsNumeric(v) Then
            If v > 0 Then
                hadData = True
                rowText = rowText & "Col" & tCols(i) & ":" & v & " "
            End If
        End If
    Next
    If hadData Then
        logFile.WriteLine "Row " & r & ": " & rowText
        foundCount = foundCount + 1
    End If
Next

logFile.WriteLine "Total Data Rows found: " & foundCount
workbook.Close False
excel.Quit
logFile.Close

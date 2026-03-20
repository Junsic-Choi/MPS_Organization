On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("sheet2_sum_final.txt", True)

Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

strPath = fso.GetAbsolutePathName(".") & "\data_working.xlsx"
Set workbook = excel.Workbooks.Open(strPath)
Set ws = workbook.Sheets(2) ' 생산배포용

logFile.WriteLine "Sheet Name: " & ws.Name
logFile.WriteLine "Scanning Row 4 for '생산'..."

Dim tCols()
ReDim tCols(0)
count = 0
For c = 5 To 100 ' Adjusted range
    val4 = ws.Cells(4, c).Value
    If Not IsEmpty(val4) Then
        If InStr(val4, "생산") > 0 Then
            ReDim Preserve tCols(count)
            tCols(count) = c
            count = count + 1
            logFile.WriteLine "  Target Col Found: " & c & " (" & ws.Cells(3, c).Value & ")"
        End If
    End If
Next

logFile.WriteLine "Total Target Columns: " & count
logFile.WriteLine "Summing rows 7 to 2000..."

totalSum = 0
For r = 7 To 2000
    rowHasData = False
    For i = 0 To count - 1
        v = ws.Cells(r, tCols(i)).Value
        If IsNumeric(v) Then
            totalSum = totalSum + CDbl(v)
            rowHasData = True
        End If
    Next
    ' Check if we hit the end
    If Not rowHasData And r > 1600 Then Exit For
Next

logFile.WriteLine "FINAL TOTAL SUM: " & totalSum
workbook.Close False
excel.Quit
logFile.Close

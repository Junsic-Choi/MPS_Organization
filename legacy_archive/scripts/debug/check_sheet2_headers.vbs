On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("sheet2_header_check.txt", True)

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(2) ' 생산배포용

logFile.WriteLine "Sheet Name: " & ws.Name
logFile.WriteLine "Row 4 Scan (1 to 200):"

For c = 1 To 200
    v4 = ws.Cells(4, c).Value
    v3 = ws.Cells(3, c).Value
    If Not IsEmpty(v4) Then
        If InStr(v4, "생산") > 0 Or InStr(v4, "Production") > 0 Then
            logFile.WriteLine "Col " & c & ": R3=[" & v3 & "] R4=[" & v4 & "]"
        End If
    End If
Next

workbook.Close False
excel.Quit
logFile.Close

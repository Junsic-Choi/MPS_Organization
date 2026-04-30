Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("all_sums.txt", True)
Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")
Set ws = workbook.Sheets(2)

For c = 5 To 40
    v3 = ws.Cells(3, c).Value
    v4 = ws.Cells(4, c).Value
    isProd = False
    If InStr(1, v4, "생산", 1) > 0 Then isProd = True
    If InStr(1, v3, "생산", 1) > 0 Then isProd = True
    
    If isProd Then
        colSum = 0
        For r = 7 To 1000
            val = ws.Cells(r, c).Value
            If IsNumeric(val) Then
                colSum = colSum + val
            End If
        Next
        logFile.WriteLine "Col " & c & " (" & v3 & "|" & v4 & "): " & colSum
    End If
Next

workbook.Close False
excel.Quit
logFile.Close

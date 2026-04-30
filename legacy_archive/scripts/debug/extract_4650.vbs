On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("_FinalList_4650.csv", True, True) ' UTF-16 for safety
outFile.WriteLine "Site,Group,Model,RPM,Month,Code,Product"

Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

strPath = fso.GetAbsolutePathName(".") & "\data_working.xlsx"
Set workbook = excel.Workbooks.Open(strPath)
If Err.Number <> 0 Then
    fso.CreateTextFile("vbs_error.txt", True).WriteLine "Error Opening: " & Err.Description
    excel.Quit
    WScript.Quit
End If

Set ws = workbook.Sheets(2) ' 생산배포용

' 1. Identify "생산" columns in Row 4
Dim tCols()
ReDim tCols(0)
count = 0
For c = 5 To 100
    v4 = ws.Cells(4, c).Value
    v3 = ws.Cells(3, c).Value
    If InStr(v4, "생산") > 0 Then
        ReDim Preserve tCols(count)
        tCols(count) = Array(c, v3)
        count = count + 1
    End If
Next

' 2. Extraction Loop
total = 0
For r = 7 To 2000
    site = ws.Cells(r, 1).Value
    group = ws.Cells(r, 2).Value
    model = ws.Cells(r, 3).Value
    rpm = ws.Cells(r, 4).Value
    
    If IsEmpty(model) And r > 1600 Then Exit For
    If Not IsEmpty(model) Then
        For i = 0 To count - 1
            colIdx = tCols(i)(0)
            monthName = tCols(i)(1)
            qty = ws.Cells(r, colIdx).Value
            If IsNumeric(qty) Then
                q = CDbl(qty)
                If q > 0 Then
                    For n = 1 To q
                        line = """" & site & """,""" & group & """,""" & model & """,""" & rpm & """,""" & monthName & ""","""","""""
                        outFile.WriteLine line
                        total = total + 1
                    Next
                End If
            End If
        Next
    End If
Next

outFile.Close
workbook.Close False
excel.Quit

Set resFile = fso.CreateTextFile("vbs_final_result.txt", True)
resFile.WriteLine "TOTAL: " & total
resFile.Close

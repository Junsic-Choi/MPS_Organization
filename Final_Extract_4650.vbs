Sub LogErr(msg)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fl = fso.OpenTextFile("vbs_error_final.txt", 8, True)
    fl.WriteLine Now & " - " & msg
    fl.Close
End Sub

On Error Resume Next
LogErr "--- VBS Extraction Started ---"

Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("_FinalList.csv", True, True)
If Err.Number <> 0 Then LogErr "Failed to create CSV: " & Err.Description : Err.Clear

Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then LogErr "Failed to create Excel object: " & Err.Description : Err.Clear
excel.Visible = False
excel.DisplayAlerts = False

strPath = fso.GetAbsolutePathName(".") & "\data_working.xlsx"
LogErr "Opening path: " & strPath
Set workbook = excel.Workbooks.Open(strPath)
If Err.Number <> 0 Then LogErr "Failed to open workbook: " & Err.Description : Err.Clear

Set ws = workbook.Sheets(2)
If Err.Number <> 0 Then LogErr "Failed to access Sheet 2: " & Err.Description : Err.Clear

targetCols = Array(5, 8, 9, 10, 11, 13)
months = Array("2월", "3월", "4월", "5월", "6월", "7월")

total = 0
For r = 7 To 2000
    site = ws.Cells(r, 1).Value
    group = ws.Cells(r, 2).Value
    model = ws.Cells(r, 3).Value
    rpm = ws.Cells(r, 4).Value
    
    If IsEmpty(model) And r > 1600 Then Exit For
    
    If Not IsEmpty(model) Then
        For i = 0 To 5
            colIdx = targetCols(i)
            monthName = months(i)
            v = ws.Cells(r, colIdx).Value
            If IsNumeric(v) Then
                qVal = CDbl(v)
                If qVal > 0 Then
                    For q = 1 To qVal
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
LogErr "Finished with total: " & total

Set fLog = fso.CreateTextFile("final_vbs_stats.txt", True)
fLog.WriteLine "TOTAL_ROWS: " & total
fLog.Close

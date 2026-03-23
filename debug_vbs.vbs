Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("vbs_step_log.txt", True)

Sub Log(msg)
    logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "Creating Excel Object..."
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    Log "ERROR Creating Excel: " & Err.Description
    WScript.Quit 1
End If

excel.Visible = False
excel.DisplayAlerts = False

strPath = fso.GetAbsolutePathName(".") & "\일반비_MPS2603-1(생산배포용).xlsx"
Log "Opening Workbook: " & strPath
Set wb = excel.Workbooks.Open(strPath, 0, True) ' ReadOnly
If Err.Number <> 0 Then
    Log "ERROR Opening Workbook: " & Err.Description
    excel.Quit
    WScript.Quit 1
End If

Log "Workbook opened successfully: " & wb.Name
Set ws = wb.Sheets(2)
Log "Accessing Sheet 2: " & ws.Name

Set outFile = fso.CreateTextFile("_FinalList.csv", True)
Log "CSV Created. Starting scan..."

targetCols = Array(5, 8, 9, 10, 11, 13)
months = Array("2월", "3월", "4월", "5월", "6월", "7월")

total = 0
For r = 7 To 3000
    model = ws.Cells(r, 3).Value
    If IsEmpty(model) And r > 2000 Then Exit For
    
    If Not IsEmpty(model) Then
        site = ws.Cells(r, 1).Value
        group = ws.Cells(r, 2).Value
        rpm = ws.Cells(r, 4).Value
        
        For i = 0 To 5
            v = ws.Cells(r, targetCols(i)).Value
            If IsNumeric(v) Then
                qty = CDbl(v)
                If qty > 0 Then
                    For q = 1 To qty
                        line = """" & site & """,""" & group & """,""" & model & """,""" & rpm & """,""" & months(i) & ""","""","""""
                        outFile.WriteLine line
                        total = total + 1
                    Next
                End If
            End If
        Next
    End If
    
    If r Mod 100 = 0 Then Log "Processed row " & r
Next

outFile.Close
Log "Finished. Total rows: " & total
wb.Close False
excel.Quit
logFile.Close

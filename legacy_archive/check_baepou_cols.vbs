Dim fso, xlApp, wb, ws, f, txt, i, j

Set fso = CreateObject("Scripting.FileSystemObject")
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim dir
dir = fso.GetAbsolutePathName(".")

' MPS2603-1.xlsx 배포용 탭
Dim mpsPath
mpsPath = dir & "\MPS2603-1.xlsx"
txt = "=== MPS2603-1.xlsx ===" & vbCrLf

On Error Resume Next
Set wb = xlApp.Workbooks.Open(mpsPath, False, True)
If Err.Number = 0 Then
    txt = txt & "Sheets: "
    For i = 1 To wb.Sheets.Count
        txt = txt & "[" & i & "]" & wb.Sheets(i).Name & "  "
    Next
    txt = txt & vbCrLf
    
    Set ws = wb.Sheets(1) ' 배포용
    txt = txt & "배포용 탭 Row 5-12, Col A-H:" & vbCrLf
    For i = 5 To 12
        Dim rowTxt
        rowTxt = "Row" & i & ": "
        For j = 1 To 8
            rowTxt = rowTxt & "[" & j & "]=" & ws.Cells(i, j).Value & " | "
        Next
        txt = txt & rowTxt & vbCrLf
    Next
    wb.Close False
Else
    txt = txt & "ERROR: " & Err.Description & vbCrLf
    Err.Clear
End If

' Real site.xlsx
Dim realPath
realPath = dir & "\Real site.xlsx"
txt = txt & vbCrLf & "=== Real site.xlsx ===" & vbCrLf
Set wb = xlApp.Workbooks.Open(realPath, False, True)
If Err.Number = 0 Then
    Set ws = wb.Sheets(1)
    txt = txt & "Sheet: " & ws.Name & vbCrLf
    txt = txt & "Row 1-5, Col A-H:" & vbCrLf
    For i = 1 To 5
        Dim rowTxt2
        rowTxt2 = "Row" & i & ": "
        For j = 1 To 8
            rowTxt2 = rowTxt2 & "[" & j & "]=" & ws.Cells(i, j).Value & " | "
        Next
        txt = txt & rowTxt2 & vbCrLf
    Next
    wb.Close False
Else
    txt = txt & "ERROR: " & Err.Description & vbCrLf
End If

xlApp.Quit

Set f = fso.CreateTextFile(dir & "\col_check_result.txt", True)
f.Write txt
f.Close

WScript.Echo "Done"

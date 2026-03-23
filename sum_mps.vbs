On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Dim fso, curDir
Set fso = CreateObject("Scripting.FileSystemObject")
curDir = fso.GetAbsolutePathName(".")

Dim wb
Set wb = objExcel.Workbooks.Open(curDir & "\data_working.xlsx", False, True)

Dim ws, s
Dim out
out = "Sheets: "
Dim mpsIndex
mpsIndex = 0

For s = 1 To wb.Sheets.Count
    out = out & wb.Sheets(s).Name & ", "
    If InStr(wb.Sheets(s).Name, "MPS") > 0 Or wb.Sheets(s).Name = "MPS" Then
        mpsIndex = s
    End If
Next
out = out & vbCrLf

If mpsIndex > 0 Then
    Set ws = wb.Sheets(mpsIndex)
    out = out & "Found MPS Sheet: " & ws.Name & vbCrLf
    
    Dim totalSum
    totalSum = 0
    Dim r, c, val, q
    Dim lastRow
    lastRow = ws.UsedRange.Rows.Count
    out = out & "Last Row: " & lastRow & vbCrLf
    
    Dim cols
    cols = Array(9, 13, 18, 23, 29, 35)
    
    For Each c In cols
        out = out & "Col " & c & " headers: " & ws.Cells(3, c).Value & " / " & ws.Cells(4, c).Value & vbCrLf
    Next
    
    For r = 7 To lastRow
        If Trim(ws.Cells(r, 3).Value) <> "" Then
            For Each c In cols
                val = ws.Cells(r, c).Value
                If IsNumeric(val) Then
                    If val > 0 Then
                        totalSum = totalSum + CDbl(val)
                    End If
                End If
            Next
        End If
    Next
    
    out = out & "Total Sum on MPS tab (Cols 9,13,18,23,29,35): " & totalSum & vbCrLf
End If

Set ts = fso.CreateTextFile("mps_vbs_out.txt", True, True) ' Unicode
ts.Write out
ts.Close

wb.Close False
objExcel.Quit

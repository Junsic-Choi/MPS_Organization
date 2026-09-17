On Error Resume Next
Dim xl, wb, toClose
Set xl = GetObject(, "Excel.Application")
If Err.Number = 0 And Not (xl Is Nothing) Then
    For Each wb In xl.Workbooks
        If InStr(LCase(wb.Name), "export") > 0 Or InStr(LCase(wb.Name), "mappe") > 0 Then
            wb.Close False
        End If
    Next
    If xl.Workbooks.Count = 0 Then
        xl.Quit
    End If
End If
On Error GoTo 0
WScript.Quit 0

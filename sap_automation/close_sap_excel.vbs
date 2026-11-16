On Error Resume Next
Dim xl, wb, retry, closedAny
closedAny = False

For retry = 1 To 10
    Err.Clear
    Set xl = GetObject(, "Excel.Application")
    If Err.Number = 0 And Not (xl Is Nothing) Then
        xl.DisplayAlerts = False
        For Each wb In xl.Workbooks
            Err.Clear
            Dim wbName
            wbName = LCase(wb.Name)
            If InStr(wbName, "export") > 0 Or InStr(wbName, "mappe") > 0 Then
                wb.Saved = True
                wb.Close False
                If Err.Number = 0 Then
                    closedAny = True
                End If
            End If
        Next
        If xl.Workbooks.Count = 0 Then
            xl.DisplayAlerts = False
            xl.Quit
            closedAny = True
        End If
        If closedAny Then Exit For
    End If
    WScript.Sleep 500
Next

On Error GoTo 0
WScript.Quit 0

Dim xlApp, wb, ws, fso, out, i, r, c, val, sVal
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("mapping_audit.txt", True)

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim filePath
filePath = fso.GetAbsolutePathName("MPS2603-1.xlsx")

Set wb = xlApp.Workbooks.Open(filePath, False, True)

out.WriteLine "Audit for 'Hutek' (휴텍) and 'XG80' in " & filePath

' 1. Check 배포용 sheet (First sheet)
Set ws = wb.Sheets(1)
out.WriteLine "--- Sheet: " & ws.Name & " (Master) ---"
For r = 1 To 2000
    val = ws.Cells(r, 3).Value ' Model column
    If Not IsEmpty(val) Then
        sVal = CStr(val)
        If InStr(sVal, "XG8") > 0 Or InStr(sVal, "휴텍") > 0 Then
            out.WriteLine "Row " & r & ": [" & ws.Cells(r, 1).Value & "] [" & ws.Cells(r, 2).Value & "] [" & sVal & "]"
        End If
    End If
    ' Also check column 1 and 2 for "휴텍"
    val = ws.Cells(r, 1).Value
    If Not IsEmpty(val) Then
        If InStr(CStr(val), "휴텍") > 0 Then
             out.WriteLine "Row " & r & " (Site): [" & CStr(val) & "] Model=[" & ws.Cells(r, 3).Value & "]"
        End If
    End If
Next

' 2. Check MPS sheet
Dim mpsWs
Set mpsWs = Nothing
On Error Resume Next
Set mpsWs = wb.Sheets("MPS")
If mpsWs Is Nothing Then Set mpsWs = wb.Sheets("mps")
If mpsWs Is Nothing Then Set mpsWs = wb.Sheets("Sheet2")
On Error GoTo 0

If Not mpsWs Is Nothing Then
    out.WriteLine "--- Sheet: " & mpsWs.Name & " (Demand) ---"
    For r = 1 To 2000
        val = mpsWs.Cells(r, 5).Value ' Product column in MPS sheet is usually 5 (E)
        If Not IsEmpty(val) Then
            sVal = CStr(val)
            If InStr(sVal, "XG8") > 0 Then
                out.WriteLine "Row " & r & ": SiteCode=[" & mpsWs.Cells(r, 7).Value & "] GroupCode=[" & mpsWs.Cells(r, 3).Value & "] Prod=[" & sVal & "]"
            End If
        End If
    Next
End If

wb.Close False
xlApp.Quit
out.Close
WScript.Echo "Audit Complete"

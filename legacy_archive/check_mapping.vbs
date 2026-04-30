Dim xlApp, wb, ws, fso, out, i, r, c, val, sVal
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("mapping_check.txt", True)

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim filePath
filePath = fso.GetAbsolutePathName("MPS2603-1.xlsx")

Set wb = xlApp.Workbooks.Open(filePath, False, True)

out.WriteLine "CHECKING FOR HUTTEK (휴텍) AND XG80"

' 1. Check first sheet (Distribution/Master)
Set ws = wb.Sheets(1)
out.WriteLine "Sheet 1: " & ws.Name
For r = 1 To 1500
    val = ws.Cells(r, 1).Value
    If InStr(CStr(val), "휴텍") > 0 Or InStr(CStr(ws.Cells(r, 2).Value), "휴텍") > 0 Or InStr(CStr(ws.Cells(r, 3).Value), "XG8") > 0 Then
        out.WriteLine "Row " & r & ": [" & ws.Cells(r,1).Value & "] [" & ws.Cells(r,2).Value & "] [" & ws.Cells(r,3).Value & "] [" & ws.Cells(r,4).Value & "]"
    End If
Next

' 2. Check MPS sheet
Set ws = Nothing
On Error Resume Next
Set ws = wb.Sheets("MPS")
If ws Is Nothing Then Set ws = wb.Sheets("mps")
On Error GoTo 0

If Not ws Is Nothing Then
    out.WriteLine "Sheet: " & ws.Name
    For r = 1 To 1500
        val = ws.Cells(r, 5).Value ' Product column
        If InStr(CStr(val), "XG8") > 0 Then
            out.WriteLine "Row " & r & ": Col3=[" & ws.Cells(r,3).Value & "] Col5=[" & ws.Cells(r,5).Value & "] Col7=[" & ws.Cells(r,7).Value & "]"
        End If
    Next
End If

wb.Close False
xlApp.Quit
out.Close
WScript.Echo "Done"

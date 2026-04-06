On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\wonjin_check_v15.txt", True)
Set xl = GetObject(, "Excel.Application")
Set wb = xl.Workbooks.Item(1)
Set ws = wb.Sheets.Item(2)

out.WriteLine "Sheet: " & ws.Name
lastSite = ""
wonjinTotal = 0

For r = 6 To 3000
    site = Trim(ws.Cells(r, 1).Text)
    If site <> "" Then lastSite = site
    
    If InStr(lastSite, "원진") > 0 Or InStr(lastSite, "06") > 0 Then
        rowSum = 0
        For c = 1 To 100
            v = ws.Cells(r, c).Value2
            If IsNumeric(v) And v > 0 Then
                rowSum = rowSum + v
            End If
        Next
        If rowSum > 0 Then wonjinTotal = wonjinTotal + rowSum
    End If
Next

out.WriteLine "Wonjin Total Units: " & wonjinTotal
out.Close

On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\cols.txt", True)
Set xl = GetObject(, "Excel.Application")
Set ws = xl.Workbooks.Item(1).Sheets.Item(2)
For c = 1 To 100
    h = ws.Cells(5, c).Text
    If InStr(h, "생산") > 0 Then
        out.WriteLine c & ":" & h
    End If
Next
out.Close

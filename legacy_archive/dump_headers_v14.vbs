On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\headers_v14.txt", True)
Set xl = GetObject(, "Excel.Application")
Set wb = xl.Workbooks.Item(1)
Set ws = wb.Sheets.Item(2)
out.WriteLine "Sheet: " & ws.Name
For c = 1 To 150
    h = ws.Cells(5, c).Text
    If h <> "" Then
        out.WriteLine c & "|[" & h & "]"
    End If
Next
out.Close

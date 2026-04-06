' dump_mps_v3.vbs
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
Set wb = xl.Workbooks.Item(1)
Set ws = wb.Sheets.Item(4)
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\mps_dump_v3.txt", True)

For r = 1 To 100
    line = "R" & r & ":"
    For c = 1 To 30
        val = ws.Cells(r, c).Text
        If val <> "" Then line = line & " [" & c & "]=" & val
    Next
    out.WriteLine line
Next
out.Close

On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\mps_scan.txt", True)
Set xl = GetObject(, "Excel.Application")
Set wb = xl.Workbooks.Item(1)

' First show all sheet names
out.WriteLine "=== Sheet Names ==="
For i = 1 To wb.Sheets.Count
    out.WriteLine i & ": " & wb.Sheets(i).Name
Next

' Scan Sheet 4 headers
out.WriteLine ""
out.WriteLine "=== Sheet 4 Headers (Row 1-5, Col 1-20) ==="
Set ws4 = wb.Sheets.Item(4)
For r = 1 To 5
    rowStr = "Row " & r & ": "
    For c = 1 To 20
        rowStr = rowStr & "[" & ws4.Cells(r, c).Text & "]"
    Next
    out.WriteLine rowStr
Next

' Show first 20 data rows
out.WriteLine ""
out.WriteLine "=== Sheet 4 Data (Rows 6-25) ==="
For r = 6 To 25
    rowStr = "Row " & r & ": "
    For c = 1 To 10
        rowStr = rowStr & "[" & ws4.Cells(r, c).Text & "]"
    Next
    out.WriteLine rowStr
Next

out.Close

On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\mps_scan2.txt", True, True)
Set xl = GetObject(, "Excel.Application")
Set wb = xl.Workbooks.Item(1)

' Sheet names
out.WriteLine "=== All Sheets ==="
For i = 1 To wb.Sheets.Count
    out.WriteLine i & ": [" & wb.Sheets(i).Name & "]"
Next

' Sheet 4 headers rows 1-6
out.WriteLine ""
out.WriteLine "=== Sheet 4 (MPS) Headers ==="
Set ws4 = wb.Sheets.Item(4)
For r = 1 To 6
    rowStr = "R" & r
    For c = 1 To 15
        rowStr = rowStr & " | C" & c & ":" & ws4.Cells(r, c).Text
    Next
    out.WriteLine rowStr
Next

' First 30 data rows
out.WriteLine ""
out.WriteLine "=== Sheet 4 Data Rows ==="
For r = 7 To 36
    rowStr = "R" & r
    For c = 1 To 10
        rowStr = rowStr & " | " & ws4.Cells(r, c).Text
    Next
    out.WriteLine rowStr
Next

out.Close
MsgBox "Done: mps_scan2.txt"

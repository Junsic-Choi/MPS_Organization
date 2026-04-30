' audit_mps_cols_v102.vbs
On Error Resume Next
Set xl = CreateObject("Excel.Application")
Set wb = xl.Workbooks.Open("c:\Users\i0215099\Desktop\MPS_UPDATE\prod_data.xlsx", 0, True, 5, "dnpc1234")
If wb Is Nothing Then WScript.Quit 1

Set ws = wb.Sheets(4)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\mps_col_audit.txt", True)

For r = 1 To 10
    line = "R" & r & ": "
    For c = 1 To 30
        line = line & c & "[" & ws.Cells(r, c).Value & "] "
    Next
    f.WriteLine line
Next
f.Close
wb.Close False
xl.Quit

' optical_audit_v103.vbs
On Error Resume Next
Set xl = CreateObject("Excel.Application")
Set wb = xl.Workbooks.Open("c:\Users\i0215099\Desktop\MPS_UPDATE\prod_data.xlsx", 0, True, 5, "dnpc1234")
If wb Is Nothing Then 
    WScript.Echo "FAIL: Workbook not opened"
    WScript.Quit 1
End If

WScript.Echo "--- SHEET 4 (MPS) ---"
Set ws = wb.Sheets(4)
For r = 1 To 5
    line = "R" & r & ": "
    For c = 1 To 10
        line = line & c & "[" & ws.Cells(r, c).Value & "] "
    Next
    WScript.Echo line
Next

WScript.Echo "--- SHEET 2 (PROD) ---"
Set ws2 = wb.Sheets(2)
For r = 1 To 7
    line = "R" & r & ": "
    For c = 1 To 15
        line = line & c & "[" & ws2.Cells(r, c).Value & "] "
    Next
    WScript.Echo line
Next

wb.Close False
xl.Quit

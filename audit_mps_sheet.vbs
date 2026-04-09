' audit_mps_sheet.vbs
On Error Resume Next
Set xl = CreateObject("Excel.Application")
Set wb = xl.Workbooks.Open("C:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용).xlsx", 0, True, 5, "dnpc1234")
If wb Is Nothing Then 
    Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile("C:\Users\i0215099\Desktop\MPS_UPDATE\mps_sheet_audit.txt", True)
    f.WriteLine "ERROR: FAILED TO OPEN WORKBOOK"
    f.Close
    WScript.Quit 1
End If

Set ws = wb.Sheets(4) ' MPS Sheet
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile("C:\Users\i0215099\Desktop\MPS_UPDATE\mps_sheet_audit.txt", True)

lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row ' xlUp
If lastRow > 1500 Then lastRow = 1500 ' Safety

arr = ws.Range("A1:E" & lastRow).Value
For r = 1 To lastRow
    line = r & "|"
    For c = 1 To 5
        line = line & arr(r, c) & "|"
    Next
    f.WriteLine line
Next

f.Close: wb.Close False: xl.Quit

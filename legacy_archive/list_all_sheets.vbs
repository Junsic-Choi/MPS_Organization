' list_all_sheets.vbs
On Error Resume Next
Set xl = CreateObject("Excel.Application")
Set wb = xl.Workbooks.Open("C:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용).xlsx", 0, True, 5, "dnpc1234")
If wb Is Nothing Then 
    Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile("C:\Users\i0215099\Desktop\MPS_UPDATE\sheet_list.txt", True)
    f.WriteLine "FAILED TO OPEN"
    f.Close
    WScript.Quit 1
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile("C:\Users\i0215099\Desktop\MPS_UPDATE\sheet_list.txt", True)

For i = 1 To wb.Sheets.Count
    Set ws = wb.Sheets(i)
    f.WriteLine i & ":" & ws.Name & "|A1:" & ws.Cells(1,1).Value & "|B1:" & ws.Cells(1,2).Value & "|C1:" & ws.Cells(1,3).Value & "|D1:" & ws.Cells(1,4).Value
Next

f.Close: wb.Close False: xl.Quit

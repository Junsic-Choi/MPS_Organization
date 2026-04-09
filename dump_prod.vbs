' dump_prod.vbs
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
Set wb = Nothing
For Each w In xl.Workbooks
    If InStr(w.Name, "MPS2603-1") > 0 Then
        Set wb = w
        Exit For
    End If
Next

If wb Is Nothing Then WScript.Quit 1

Set ws = wb.Sheets(2)
arr = ws.UsedRange.Value

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\raw_prod_vbs.txt", True)

For r = 1 To UBound(arr, 1)
    line = ""
    For c = 1 To 15 ' Only need first 15 columns
        line = line & arr(r, c) & "||"
    Next
    f.WriteLine line
Next
f.Close

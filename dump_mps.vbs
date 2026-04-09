' dump_mps.vbs
Set xl = GetObject(, "Excel.Application")
Set wb = Nothing
For Each w In xl.Workbooks
    If InStr(w.Name, "MPS2603-1") > 0 Then
        Set wb = w
        Exit For
    End If
Next

If Not wb Is Nothing Then
    Set ws = wb.Sheets(4)
    arr = ws.Range("A1:AD1500").Value
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\mps_raw_vbs.txt", True)
    For r = 1 To UBound(arr, 1)
        line = ""
        For c = 1 To 10 ' First 10 columns
            line = line & arr(r, c) & "||"
        Next
        f.WriteLine line
    Next
    f.Close
End If

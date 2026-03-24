On Error Resume Next
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetAbsolutePathName(".") & "\data_working.xlsx"
WScript.Echo "Opening: " & strPath
Set wb = excel.Workbooks.Open(strPath)
If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
Else
    WScript.Echo "Success! Sheets:"
    For Each sh In wb.Sheets
        WScript.Echo " - " & sh.Name
    Next
    wb.Close False
End If
excel.Quit

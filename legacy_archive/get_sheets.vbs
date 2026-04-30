Dim fso, xlApp, wb, ws, fn, outFile, txt
Dim i

fn = "MPS2603-1.xlsx"
outFile = "sheet_names_result.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim fullPath
fullPath = fso.GetAbsolutePathName(fn)

If Not fso.FileExists(fullPath) Then
    txt = "FILE_NOT_FOUND: " & fullPath
Else
    On Error Resume Next
    Set wb = xlApp.Workbooks.Open(fullPath, False, True)
    If Err.Number <> 0 Then
        txt = "OPEN_ERROR: " & Err.Description
    Else
        txt = "SHEETS (" & wb.Sheets.Count & "):" & vbCrLf
        For i = 1 To wb.Sheets.Count
            txt = txt & "  [" & i & "] " & wb.Sheets(i).Name & vbCrLf
        Next
        wb.Close False
    End If
End If

xlApp.Quit

Dim f
Set f = fso.CreateTextFile(outFile, True)
f.Write txt
f.Close

WScript.Echo "Done: " & outFile

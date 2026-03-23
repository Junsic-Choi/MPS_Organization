On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
currDir = fso.GetAbsolutePathName(".")
Set logFile = fso.CreateTextFile(currDir & "\vbs_audit_debug.txt", True)
logFile.WriteLine "VBS Audit Started at " & Now

Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    logFile.WriteLine "Error creating Excel Object: " & Err.Description
    WScript.Quit 1
End If

objExcel.Visible = False
objExcel.DisplayAlerts = False

filePath = currDir & "\일반비_MPS2603-1(생산배포용).xlsx"
logFile.WriteLine "Opening File: " & filePath

Set objWorkbook = objExcel.Workbooks.Open(filePath, 0, True)
If Err.Number <> 0 Then
    logFile.WriteLine "Error opening Workbook: " & Err.Description
    objExcel.Quit
    WScript.Quit 1
End If

Set outFile = fso.CreateTextFile(currDir & "\all_sheets_vbs_audit_v2.txt", True)

For Each objSheet In objWorkbook.Sheets
    outFile.WriteLine "--- Sheet: " & objSheet.Name & " ---"
    For r = 3 To 4
        For c = 1 To 100
            val = objSheet.Cells(r, c).Value
            If val <> "" Then
                outFile.WriteLine "R" & r & " C" & c & " : [" & val & "]"
            End If
        Next
    Next
    outFile.WriteLine ""
Next

outFile.Close
objWorkbook.Close False
objExcel.Quit
logFile.WriteLine "VBS Audit Finished at " & Now
logFile.Close

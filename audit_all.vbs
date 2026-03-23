Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set fso = CreateObject("Scripting.FileSystemObject")
currDir = fso.GetAbsolutePathName(".")
filePath = currDir & "\일반비_MPS2603-1(생산배포용).xlsx"

Set objWorkbook = objExcel.Workbooks.Open(filePath, 0, True)
Set outFile = fso.CreateTextFile(currDir & "\all_sheets_vbs_audit.txt", True, True)

For Each objSheet In objWorkbook.Sheets
    outFile.WriteLine "--- Sheet: " & objSheet.Name & " ---"
    For r = 3 To 4
        For c = 1 To 50
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

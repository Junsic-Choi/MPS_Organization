Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set FSO = CreateObject("Scripting.FileSystemObject")
Set logFile = FSO.CreateTextFile("diag_headers.txt", True)

On Error Resume Next

Dim currentDir
currentDir = FSO.GetAbsolutePathName(".")

Dim matchingFile
matchingFile = ""
For Each file In FSO.GetFolder(currentDir).Files
    If InStr(file.Name, "생산배포용") > 0 And InStr(file.Name, ".xlsx") > 0 Then
        matchingFile = file.Path
        Exit For
    End If
Next

If matchingFile = "" Then
    logFile.WriteLine "ERROR: Excel file not found"
Else
    logFile.WriteLine "Opening: " & matchingFile
    Set objWorkbook = objExcel.Workbooks.Open(matchingFile, 0, True)
    
    Set objSheet = Nothing
    For Each sh In objWorkbook.Sheets
        If sh.Name = "생산배포용" Then
            Set objSheet = sh
            Exit For
        End If
    Next

    If Not objSheet Is Nothing Then
        logFile.WriteLine "Found sheet: 생산배포용"
        logFile.WriteLine "--- Row 7 Headers ---"
        For c = 1 To 30
            txt = objSheet.Cells(7, c).Value
            logFile.WriteLine "Col " & c & ": [" & txt & "]"
        Next

        logFile.WriteLine "--- Row 8 Values ---"
        For c = 1 To 30
            txt = objSheet.Cells(8, c).Value
            logFile.WriteLine "Row 8 Col " & c & ": [" & txt & "]"
        Next
    Else
        logFile.WriteLine "ERROR: sheet not found"
    End If
    
    objWorkbook.Close False
End If

objExcel.Quit
logFile.Close

If Err.Number <> 0 Then
    Set f = FSO.OpenTextFile("diag_headers.txt", 8)
    f.WriteLine "EXCEPTION: " & Err.Description
    f.Close
End If

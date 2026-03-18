On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set logFile = FSO.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\diag_headers_final.txt", True)

logFile.WriteLine "Start Diag: " & Now

Set folder = FSO.GetFolder("c:\Users\i0215099\Desktop\MPS_UPDATE")
For Each file In folder.Files
    If InStr(file.Name, "생산배포용") > 0 And InStr(file.Name, ".xlsx") > 0 Then
        logFile.WriteLine "Opening: " & file.Name
        Set wb = objExcel.Workbooks.Open(file.Path, 0, True)
        Set sh = Nothing
        For Each s In wb.Sheets
            If s.Name = "생산배포용" Then Set sh = s
        Next
        
        If Not sh Is Nothing Then
            logFile.WriteLine "Headers in Row 7:"
            For c = 1 To 26
                logFile.WriteLine "Col " & c & ": [" & sh.Cells(7, c).Text & "]"
            Next
        Else
            logFile.WriteLine "Sheet NOT FOUND"
        End If
        wb.Close False
        Exit For
    End If
Next

objExcel.Quit
logFile.Close

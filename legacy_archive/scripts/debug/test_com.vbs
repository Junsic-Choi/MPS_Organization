On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set log = fso.CreateTextFile("vbs_com_test.txt", True)
log.WriteLine "--- Starting COM Test ---"

Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    log.WriteLine "FAILED to create Excel Object: " & Err.Description
    excel.Quit
    log.Close
    WScript.Quit
End If
log.WriteLine "SUCCESS: Excel Object Created."

strPath = fso.GetAbsolutePathName(".") & "\data_working.xlsx"
log.WriteLine "Opening: " & strPath
Set workbook = excel.Workbooks.Open(strPath)
If Err.Number <> 0 Then
    log.WriteLine "FAILED to open workbook: " & Err.Description
Else
    log.WriteLine "SUCCESS: Workbook opened. Sheet count: " & workbook.Sheets.Count
    workbook.Close False
End If

excel.Quit
log.WriteLine "--- Test Finished ---"
log.Close

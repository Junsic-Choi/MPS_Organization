Set objExcel = CreateObject("Excel.Application")
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\i0215099\Desktop\MPS_UPDATE\MPS2603-1.xlsx", 0, True)
If Err.Number <> 0 Then
    WScript.Echo "Error opening workbook: " & Err.Description
    objExcel.Quit
    WScript.Quit
End If
Set objWorksheet = objWorkbook.Sheets("MPS")
rowCount = objWorksheet.UsedRange.Rows.Count
WScript.Echo "RowCount:" & rowCount
objWorkbook.Close False
objExcel.Quit

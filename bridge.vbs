On Error Resume Next
Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "Usage: bridge.vbs <input_xlsx> <output_csv>"
    WScript.Quit 1
End If

inputFile = args(0)
outputFile = args(1)

Set objExcel = CreateObject("Excel.Application")
If objExcel Is Nothing Then
    WScript.Echo "Error: Could not start Excel"
    WScript.Quit 1
End If

objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(inputFile, 0, True)
If objWorkbook Is Nothing Then
    WScript.Echo "Error: Could not open workbook. DRM might be blocking."
    objExcel.Quit
    WScript.Quit 1
End If

' Save "생산배포용" sheet as CSV
' Find sheet by name
targetSheet = Null
For Each sh In objWorkbook.Sheets
    If InStr(sh.Name, "생산배포용") > 0 Then
        Set targetSheet = sh
        Exit For
    End If
Next

If targetSheet Is Nothing Then
    ' Fallback to sheet 2
    Set targetSheet = objWorkbook.Sheets(2)
End If

targetSheet.Activate
objWorkbook.SaveAs outputFile, 6 ' 6 = xlCSV

objWorkbook.Close False
objExcel.Quit

If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
    WScript.Quit 1
Else
    WScript.Echo "Success"
    WScript.Quit 0
End If

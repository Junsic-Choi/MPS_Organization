Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set fso = CreateObject("Scripting.FileSystemObject")
currDir = fso.GetAbsolutePathName(".")
filePath = currDir & "\일반비_MPS2603-1(생산배포용).xlsx"

Set objWorkbook = objExcel.Workbooks.Open(filePath, 0, True)
Set objSheet = objWorkbook.Sheets(2)

Set outFile = fso.CreateTextFile(currDir & "\header_map_vbs.txt", True, True) ' Unicode

For r = 1 To 10
    line = "Row " & r & " : "
    For c = 1 To 50
        val = objSheet.Cells(r, c).Value
        line = line & "[" & val & "] "
    Next
    outFile.WriteLine line
Next

outFile.Close
objWorkbook.Close False
objExcel.Quit

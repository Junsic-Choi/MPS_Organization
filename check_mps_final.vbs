Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set fso = CreateObject("Scripting.FileSystemObject")
currDir = fso.GetAbsolutePathName(".")
filePath = currDir & "\일반비_MPS2603-1(생산배포용).xlsx"
Set wb = objExcel.Workbooks.Open(filePath, 0, True)
Set ws = wb.Sheets.Item(4)
count = 0
r = 1
Do While r < 5000
    For c = 1 To 50
        val = ws.Cells(r, c).Value
        If IsNumeric(val) And val > 0 Then
            count = count + val
        End If
    Next
    r = r + 1
Loop
Set outFile = fso.CreateTextFile(currDir & "\mps_total_qty.txt", True)
outFile.WriteLine "Total Qty on Sheet 4: " & count
outFile.Close
wb.Close False
objExcel.Quit

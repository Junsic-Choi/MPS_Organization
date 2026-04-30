Dim xlApp, wb, ws, fso, out, i, r, c
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("search_results.txt", True)

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim filePath
filePath = fso.GetAbsolutePathName("MPS2603-1.xlsx")

Set wb = xlApp.Workbooks.Open(filePath, False, True)

out.WriteLine "Searching for 'XG80' and '휴텍' in " & filePath

For Each ws In wb.Sheets
    out.WriteLine "--- Sheet: " & ws.Name & " ---"
    For r = 1 To 1000 ' Check first 1000 rows
        Dim rowText
        rowText = ""
        Dim found
        found = False
        For c = 1 To 20 ' Check first 20 columns
            Dim val
            val = ws.Cells(r, c).Value
            If Not IsEmpty(val) Then
                Dim sVal
                sVal = CStr(val)
                If InStr(sVal, "XG80") > 0 Or InStr(sVal, "휴텍") > 0 Then
                    found = True
                End If
                rowText = rowText & "[" & c & "]" & sVal & " | "
            End If
        Next
        If found Then
            out.WriteLine "Row " & r & ": " & rowText
        End If
    Next
Next

wb.Close False
xlApp.Quit
out.Close
WScript.Echo "Search Complete"

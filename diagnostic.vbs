' diagnostic_v105.vbs
Set xl = CreateObject("Excel.Application")
WScript.Echo "Excel Object Created"
xl.Visible = False
Set wb = xl.Workbooks.Open("c:\Users\i0215099\Desktop\MPS_UPDATE\prod_data.xlsx", 0, True, 5, "dnpc1234")
WScript.Echo "Workbook Object: " & TypeName(wb)
If Not wb Is Nothing Then
    WScript.Echo "Success: Opened " & wb.Name
    wb.Close False
End If
xl.Quit

' dump_to_csv.vbs
On Error Resume Next
Set xl = CreateObject("Excel.Application")
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open("C:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용).xlsx", 0, True, 5, "dnpc1234")
If Not wb Is Nothing Then
    ' Sheet 4 (MPS)
    wb.Sheets(4).SaveAs "C:\Users\i0215099\Desktop\MPS_UPDATE\mps_ref.csv", 6 ' 6 = xlCSV
    ' Sheet 2 (Production)
    wb.Sheets(2).SaveAs "C:\Users\i0215099\Desktop\MPS_UPDATE\prod_raw.csv", 6
    wb.Close False
End If
xl.Quit

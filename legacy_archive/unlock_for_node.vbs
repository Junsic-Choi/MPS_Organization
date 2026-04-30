' unlock_for_node.vbs
On Error Resume Next
Set xl = CreateObject("Excel.Application")
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open("C:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용).xlsx", 0, True, 5, "dnpc1234")
If Not wb Is Nothing Then
    wb.SaveAs "C:\Users\i0215099\Desktop\MPS_UPDATE\temp_mps_unlocked.xlsx", 51
    wb.Close False
End If
xl.Quit

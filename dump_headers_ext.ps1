$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(2)
$res = ""
for ($c = 1; $c -le 50; $c++) {
    $v3 = $ws.Cells.Item(3, $c).Text
    $v4 = $ws.Cells.Item(4, $c).Text
    $res += "Col $c : R3=[$v3] R4=[$v4]`n"
}
$res | Out-File -FilePath "$dir\sheet2_extended_headers.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()

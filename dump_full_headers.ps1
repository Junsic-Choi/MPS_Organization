$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(2)
$res = "Row 3 & 4 Head (1-100):`n"
for ($c = 1; $c -le 100; $c++) {
    $v3 = "$($ws.Cells.Item(3, $c).Value2)"
    $v4 = "$($ws.Cells.Item(4, $c).Value2)"
    if ($v3 -ne "" -or $v4 -ne "") {
        $res += "Col $c : R3=[$v3] R4=[$v4]`n"
    }
}
$res | Out-File -FilePath "$dir\full_header_map.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()

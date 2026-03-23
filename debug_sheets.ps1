$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$res = "Sheets:`n"
foreach ($s in $wb.Sheets) {
    $res += "- " + $s.Name + "`n"
}

$ws = $wb.Sheets.Item(2)
$res += "`nSheet 2 (Index 2) Name: " + $ws.Name + "`n"
$res += "Row 4 Scan:`n"
for ($c = 1; $c -le 100; $c++) {
    $v4 = "$($ws.Cells.Item(4, $c).Value2)"
    if ($v4 -ne "") {
        $res += "Col $c : Value2=[$v4]`n"
    }
}
$res | Out-File -FilePath "$dir\sheet_debug.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$temp = "$dir\temp_data_dump.xlsx"
Copy-Item $path $temp -Force
$wb = $excel.Workbooks.Open($temp, 0, $true)
$ws = $wb.Sheets.Item(2)
$res = ""
for ($r = 1; $r -le 50; $r++) {
    $line = "Row $r : "
    for ($c = 1; $c -le 15; $c++) {
        $v = "$($ws.Cells.Item($r, $c).Value2)"
        $line += "[$v] "
    }
    $res += $line + "`n"
}
$res | Out-File -FilePath "$dir\sheet2_data_dump.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()
Remove-Item $temp -ErrorAction SilentlyContinue

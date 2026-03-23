$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
$temp = "$dir\temp_header_map.xlsx"
Copy-Item $path $temp -Force
$wb = $excel.Workbooks.Open($temp, 0, $true)
$ws = $wb.Sheets.Item(2)
$csv = ""
for ($r = 1; $r -le 10; $r++) {
    $rowArr = @()
    for ($c = 1; $c -le 50; $c++) {
        $v = "$($ws.Cells.Item($r, $c).Value2)"
        $rowArr += "`"$v`""
    }
    $csv += ($rowArr -join ",") + "`n"
}
$csv | Out-File -FilePath "$dir\header_map.csv" -Encoding UTF8
$wb.Close($false)
$excel.Quit()
Remove-Item $temp -ErrorAction SilentlyContinue

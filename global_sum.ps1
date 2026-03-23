$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(2)
$totalSum = 0
$lastRow = $ws.UsedRange.Rows.Count
$lastCol = $ws.UsedRange.Columns.Count

for ($r = 7; $r -le $lastRow; $r++) {
    for ($c = 5; $c -le $lastCol; $c++) {
        $val = $ws.Cells.Item($r, $c).Value2
        if ($null -ne $val -and [double]$val -gt 0) {
            $totalSum += [double]$val
        }
    }
}

"Global Sum (Sheet 2, Col 5+): $totalSum" | Out-File -FilePath "$dir\global_sum.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()

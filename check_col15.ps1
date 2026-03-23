$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(2)
$sum = 0
for ($r = 7; $r -le $ws.UsedRange.Rows.Count; $r++) {
    $val = $ws.Cells.Item($r, 15).Value2
    if ($null -ne $val -and [double]$val -gt 0) {
        $sum += [double]$val
    }
}
"Sum of Col 15 (Sheet 2): $sum" | Out-File -FilePath "$dir\col15_sum.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()

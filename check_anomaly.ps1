$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(2)
$max = 0
$rowsWithLargeVal = 0
for ($r = 7; $r -le $ws.UsedRange.Rows.Count; $r++) {
    $val = $ws.Cells.Item($r, 6).Value2
    if ($null -ne $val -and [double]$val -gt $max) {
        $max = [double]$val
    }
    if ($null -ne $val -and [double]$val -gt 1000) {
        $rowsWithLargeVal++
    }
}
"Max Value in Col 6: $max`nRows with Value > 1000: $rowsWithLargeVal" | Out-File -FilePath "$dir\col6_anomaly.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()

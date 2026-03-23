$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(4)
$total = 0
$targetCols = @(9, 13, 18, 23, 29, 35)
for ($r = 7; $r -le $ws.UsedRange.Rows.Count; $r++) {
    foreach ($c in $targetCols) {
        $val = $ws.Cells.Item($r, $c).Value2
        if ($null -ne $val -and [double]$val -gt 0) {
            $total += [double]$val
        }
    }
}
"Sheet 4 Target Total: $total" | Out-File -FilePath "$dir\s4_final_total.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()

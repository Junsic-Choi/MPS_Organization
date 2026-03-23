$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(2)
$totalSum = 0
$colReport = ""

for ($c = 1; $c -le 200; $c++) {
    $v4 = "$($ws.Cells.Item(4, $c).Value2)"
    if ($v4 -match "생산") {
        $colSum = 0
        for ($r = 7; $r -le $ws.UsedRange.Rows.Count; $r++) {
            $val = $ws.Cells.Item($r, $c).Value2
            if ($null -ne $val -and [double]$val -gt 0) {
                $colSum += [double]$val
            }
        }
        $totalSum += $colSum
        $v3 = "$($ws.Cells.Item(3, $c).Value2)"
        $colReport += "Col $c ($v3): $colSum`n"
    }
}

$res = "Total Sum of Production: $totalSum`n`nBreakdown by column:`n$colReport"
$res | Out-File -FilePath "$dir\quantity_audit.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()

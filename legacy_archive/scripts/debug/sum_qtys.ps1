$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = Get-Location
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    $cols = @(5, 8, 9, 10, 11, 13, 15)
    $sums = @{}
    foreach ($c in $cols) { $sums[$c] = 0 }
    
    for ($r = 7; $r -le 1000; $r++) {
        foreach ($c in $cols) {
            $val = $ws.Cells.Item($r, $c).Value2
            if ($null -ne $val -and $val -gt 0) {
                $sums[$c] += [int]$val
            }
        }
    }
    
    $total = 0
    $res = ""
    foreach ($c in $cols) {
        $res += "Col $($c) : $($sums[$c]) `r`n"
        $total += $sums[$c]
    }
    $res += "Total: $total"
    $res | Out-File "$dir\sums_debug.txt" -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}

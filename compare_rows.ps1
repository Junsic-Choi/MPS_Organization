$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$log = "$dir\row_comparison.txt"
"Comparing Row 7 and 8 (and others)..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    $targetCols = @(5, 8, 9, 10, 11, 13)
    
    for ($r = 7; $r -le 20; $r++) {
        $m = "$($ws.Cells.Item($r, 3).Text)"
        $rowStr = "Row $r | Model: [$m] | "
        foreach ($c in $targetCols) {
            $v = $ws.Cells.Item($r, $c).Value2
            $rowStr += "Col $c: [$v] "
        }
        $rowStr | Out-File $log -Append -Encoding UTF8
    }
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}

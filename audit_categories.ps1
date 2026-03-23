$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\category_audit_log.txt"
if (Test-Path $log) { Remove-Item $log }
Start-Transcript -Path $log -Force

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    $ws = $workbook.Sheets.Item(2)
    
    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = 50 # Check first 50 columns
    
    "Sheet 2 Audit (Row 7 to $lastRow):" | Write-Output
    
    for ($c = 5; $c -le $lastCol; $c++) {
        $v3 = "$($ws.Cells.Item(3, $c).Value2)"
        $v4 = "$($ws.Cells.Item(4, $c).Value2)"
        
        $sum = 0
        for ($r = 7; $r -le $lastRow; $r++) {
            $val = $ws.Cells.Item($r, $c).Value2
            if ($null -ne $val -and [double]$val -gt 0) {
                $sum += [double]$val
            }
        }
        
        if ($sum -gt 0) {
            Write-Output "Col $c : R3=[$v3] R4=[$v4] Sum=$sum"
        }
    }
}
catch {
    Write-Output "ERROR: $_"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    Stop-Transcript
}

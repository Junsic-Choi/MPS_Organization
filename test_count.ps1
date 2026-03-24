$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $path = "c:\Users\i0215099\Desktop\MPS_UPDATE\data_working.xlsx"
    if (!(Test-Path $path)) {
        Write-Output "File not found!"
        exit
    }
    $wb = $excel.Workbooks.Open($path, 0, $true)
    $ws = $wb.Sheets.Item(2)
    
    $cols = @(9, 13, 18, 23, 29, 35) # I, M, R, W, AC, AI
    $total = 0
    foreach ($c in $cols) {
        $sum = 0
        for ($r = 7; $r -le 2000; $r++) {
            $model = $ws.Cells.Item($r, 3).Value2
            $val = $ws.Cells.Item($r, $c).Value2
            if ($null -ne $model -and "$model".Trim() -ne "") {
                if ($null -ne $val -and [double]$val -gt 0) {
                    $sum += [math]::Floor([double]$val)
                }
            }
        }
        $v3 = $ws.Cells.Item(3, $c).Value2
        $v4 = $ws.Cells.Item(4, $c).Value2
        Write-Output "Col $c ($v3 / $v4): $sum"
        $total += $sum
    }
    Write-Output "Total: $total"
}
catch {
    Write-Output $_.Exception.Message
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    $excel.Quit()
}

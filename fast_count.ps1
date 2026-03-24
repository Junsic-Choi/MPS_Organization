$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $path = "c:\Users\i0215099\Desktop\MPS_UPDATE\data_working.xlsx"
    $wb = $excel.Workbooks.Open($path, 0, $true)
    $ws = $wb.Sheets.Item(2)
    
    # Fast read
    $data = $ws.UsedRange.Value2
    $rows = $data.GetLength(0)
    $cols = $data.GetLength(1)
    
    $colsToCheck = @(9, 13, 18, 23, 29, 35) # I, M, R, W, AC, AI
    $total = 0
    foreach ($c in $colsToCheck) {
        $sum = 0
        if ($c -le $cols) {
            for ($r = 7; $r -le $rows; $r++) {
                $val = $data[$r, $c]
                $model = $data[$r, 3]
                if ($null -ne $model -and "$model".Trim() -ne "") {
                    if ($null -ne $val -and [double]$val -gt 0) {
                        $sum += [math]::Floor([double]$val)
                    }
                }
            }
            $v3 = $data[3, $c]
            $v4 = $data[4, $c]
            Write-Output "Col $c ($v3 / $v4): $sum"
        }
        else {
            Write-Output "Col $c is out of bounds!"
        }
        $total += $sum
    }
    Write-Output "Total for these columns: $total"
    
}
catch {
    Write-Output "Error: $_"
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    $excel.Quit()
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = $PSScriptRoot
    $file = Get-ChildItem -Path $dir -Filter "*MPS*(생산배포용)*.xlsx" | Select-Object -First 1
    if (-not $file) { throw "MPS file not found" }
    Write-Host "Opening: $($file.Name)"
    $wb = $excel.Workbooks.Open($file.FullName, 0, $true)
    $sh = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq '생산배포용') { $sh = $s; break } }
    if ($sh) {
        Write-Host "--- Row 7 Headers (Col 1 to 25) ---"
        for ($c = 1; $c -le 25; $c++) {
            $val = $sh.Cells.Item(7, $c).Text
            Write-Host "Col $c: [$val]"
        }
    }
    $wb.Close($false)
}
finally {
    $excel.Quit()
}

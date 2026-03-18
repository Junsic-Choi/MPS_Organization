$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = $PSScriptRoot
    # Use generic filter and check for MPS keywords inside
    $file = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { $_.Name -like "*MPS*" } | Select-Object -First 1
    if (-not $file) { throw "MPS file not found in $dir" }
    Write-Host "Opening: $($file.Name)"
    $wb = $excel.Workbooks.Open($file.FullName, 0, $true)
    $sh = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq 'MPS') { $sh = $s; break } }
    if ($sh) {
        $range = $sh.UsedRange.Value2
        $rows = $range.GetUpperBound(0)
        Write-Host "--- SMX/VCF Search in MPS ---"
        for ($r = 6; $r -le $rows; $r++) {
            $name = [string]$range[$r, 5]
            $code = [string]$range[$r, 4]
            if ($name -match "SMX|VCF|VF|2100|2600|3100|850") {
                Write-Host ("Row {0}: Code=[{1}] Name=[{2}]" -f $r, $code, $name)
            }
        }
    }
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = $PSScriptRoot
    $file = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { $_.Name -like "*MPS*" } | Select-Object -First 1
    if (-not $file) { throw "MPS file not found" }
    $wb = $excel.Workbooks.Open($file.FullName, 0, $true)
    $sh = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq 'MPS') { $sh = $s; break } }
    if ($sh) {
        $range = $sh.UsedRange.Value2
        $rows = $range.GetUpperBound(0)
        Write-Host "--- Specific Search in MPS (PUMA ST, DNM, LYNX) ---"
        for ($r = 6; $r -le $rows; $r++) {
            $name = [string]$range[$r, 5]
            $code = [string]$range[$r, 4]
            # Search for PUMA ST, DNM, LYNX
            if ($name -match "PST|10GS|DNM|LYNX|LYN|MYNX|MYN|DNT|DNX") {
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

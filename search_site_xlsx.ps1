$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $path = "c:\Users\i0215099\Desktop\MPS_UPDATE\site.xlsx"
    if (-not (Test-Path $path)) { throw "site.xlsx not found" }
    $wb = $excel.Workbooks.Open($path, 0, $true)
    $sh = $wb.Sheets.Item(1)
    
    $targets = @("NHM5000", "PUMA", "NHM6300", "LYNX")
    Write-Host "--- Searching site.xlsx for $targets ---"
    
    $rows = $sh.UsedRange.Rows.Count
    for ($r = 1; $r -le $rows; $r++) {
        $plant = "$($sh.Cells.Item($r, 3).Value2)"
        $code = "$($sh.Cells.Item($r, 4).Value2)"
        $desc = "$($sh.Cells.Item($r, 5).Value2)"
        
        foreach ($t in $targets) {
            if ($code -match $t -or $desc -match $t) {
                Write-Host "Found in site.xlsx Row $r : Plant=[$plant] Code=[$code] Desc=[$desc]"
            }
        }
        if ($r -gt 10000) { break }
    }
    $wb.Close($false)
}
catch {
    Write-Host "!! Exception: $_"
}
finally { $excel.Quit() }

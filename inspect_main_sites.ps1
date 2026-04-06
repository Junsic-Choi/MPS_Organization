$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
    $f = Get-ChildItem -Path $dir -Filter "*MPS2603-1*" | Select-Object -First 1
    $wb = $excel.Workbooks.Open($f.FullName, 0, $true)
    $ws = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -match "생산" -or $s.Name -match "Production") { $ws = $s; break } }
    if (-not $ws) { $ws = $wb.Sheets.Item(2) }

    $last = $ws.UsedRange.Rows.Count
    if ($last -gt 2000) { $last = 2000 }
    
    $map = @{}
    for ($r=1; $r -le $last; $r++) {
        $siteName = "$($ws.Cells.Item($r, 1).Text)".Trim()
        $siteCode = "$($ws.Cells.Item($r, 7).Text)".Trim() # Looking at Col G for code
        if ($siteName -and $siteCode -and $siteName -ne "생산처") {
            if (-not $map.ContainsKey($siteCode)) {
                $map[$siteCode] = $siteName
            }
        }
    }
    
    $res = @()
    foreach ($k in $map.Keys) { $res += "$k | $($map[$k])" }
    $res | Out-File "$dir\main_site_map_debug.txt" -Encoding UTF8
    $wb.Close($false)
} finally {
    $excel.Quit()
}

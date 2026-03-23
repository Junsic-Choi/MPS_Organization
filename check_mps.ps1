$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $dir = Get-Location
    $files = Get-ChildItem -Path $dir -Filter "일반비_*.xlsx"
    if ($files.Count -eq 0) {
        "No matching file found." | Out-File "$dir\ps_out.txt"
        exit
    }
    $path = $files[0].FullName
    "Opening file: $path" | Out-File "$dir\ps_out.txt"

    $wb = $excel.Workbooks.Open($path, 0, $true)
    
    $res = "Sheets:`n"
    foreach ($s in $wb.Sheets) {
        $res += "- " + $s.Name + "`n"
    }
    
    $res | Out-File -Append "$dir\ps_out.txt" -Encoding UTF8

    foreach ($s in $wb.Sheets) {
        if ($s.Name -match "MPS" -or $s.Name -match "생산배포용") {
            $sheetRes = "`nFound sheet: " + $s.Name + "`n"
            $sheetRes += "Row 3 & 4 (Columns 1 to 50):`n"
            for ($c = 1; $c -le 50; $c++) {
                $v3 = "$($s.Cells.Item(3, $c).Value2)"
                $v4 = "$($s.Cells.Item(4, $c).Value2)"
                if ($v3 -ne "" -or $v4 -ne "") {
                    $sheetRes += "Col $c : V3=[$v3] V4=[$v4]`n"
                }
            }
            $sheetRes | Out-File -Append "$dir\ps_out.txt" -Encoding UTF8
        }
    }

    $wb.Close($false)
    $excel.Quit()
}
catch {
    "Error: $_" | Out-File -Append "$dir\ps_out.txt" -Encoding UTF8
    if ($null -ne $wb) { $wb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}

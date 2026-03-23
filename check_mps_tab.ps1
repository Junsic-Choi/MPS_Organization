$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $dir = Get-Location
    $path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
    $wb = $excel.Workbooks.Open($path, 0, $true)
    
    $mpsSheet = $null
    foreach ($s in $wb.Sheets) {
        if ($s.Name -match "MPS" -and $s.Name -notmatch "생산배포용") {
            $mpsSheet = $s
            break
        }
    }
    
    if ($null -eq $mpsSheet) {
        # Fallback if there is a combined name
        foreach ($s in $wb.Sheets) {
            if ($s.Name -match "MPS") {
                $mpsSheet = $s
                break
            }
        }
    }

    if ($null -ne $mpsSheet) {
        $res = "MPS Sheet Name: $($mpsSheet.Name)`n"
        $lastRow = $mpsSheet.UsedRange.Rows.Count
        $res += "Used Rows: $lastRow`n"
        
        $res += "Headers (Row 3 & 4) for specific columns (I=9, M=13, R=18, W=23, AC=29, AI=35, AO=41):`n"
        foreach ($c in @(9, 13, 18, 23, 29, 35, 41)) {
            $v3 = "$($mpsSheet.Cells.Item(3, $c).Text)"
            $v4 = "$($mpsSheet.Cells.Item(4, $c).Text)"
            $res += "Col $c : V3=[$v3] V4=[$v4]`n"
        }
        
        # Test how many rows have a value in Column A or B to see if there are 4650 rows
        $dataRows = 0
        $sumMap = @{}
        for ($r = 5; $r -le $lastRow; $r++) {
            $val = $mpsSheet.Cells.Item($r, 3).Text
            if ($val -ne "") {
                $dataRows++
                foreach ($c in @(9, 13, 18, 23, 29, 35)) {
                    $q = $mpsSheet.Cells.Item($r, $c).Value2
                    if ($null -ne $q -and [double]$q -gt 0) {
                        $sumMap[$c] += [double]$q
                    }
                }
            }
        }
        $res += "Data Rows (Col 3 not empty): $dataRows`n"
        foreach ($k in $sumMap.Keys | Sort-Object) {
            $res += "Sum of Col $k: $($sumMap[$k])`n"
        }
        $res | Out-File -FilePath "$dir\mps_info.txt" -Encoding UTF8
    }
    else {
        "MPS Sheet NOT found." | Out-File -FilePath "$dir\mps_info.txt" -Encoding UTF8
    }

    $wb.Close($false)
    $excel.Quit()
}
catch {
    "Error: $_" | Out-File -FilePath "$dir\mps_info.txt" -Encoding UTF8
    if ($null -ne $wb) { $wb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}

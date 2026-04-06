# robust_map_v3.ps1
$logFile = "c:\Users\i0215099\Desktop\MPS_UPDATE\mapping_v3_log.txt"
$csvFile = "c:\Users\i0215099\Desktop\MPS_UPDATE\Final_Result_4650.csv"

Function Write-Log($msg) {
    "[$((Get-Date).ToString('HH:mm:ss'))] $msg" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

if (Test-Path $logFile) { Remove-Item $logFile }
if (Test-Path $csvFile) { Remove-Item $csvFile }

Write-Log "Starting Robust Map V3 (Total 4650 Target)"

try {
    # 1. Connect to Excel
    $xl = $null
    try {
        $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $true
    }
    
    # 2. Get Workbook
    $wb = $null
    foreach($w in $xl.Workbooks) {
        if ($w.Name -like "*MPS2603*") { $wb = $w; break }
    }
    if ($null -eq $wb) {
        $file = Get-ChildItem -Path "c:\Users\i0215099\Desktop\MPS_UPDATE" -Filter "*MPS2603*.xlsx" | Select-Object -First 1 -ExpandProperty FullName
        $wb = $xl.Workbooks.Open($file, 0, $true)
    }

    $wsP = $wb.Worksheets.Item(2) # 생산배포용 탭
    $wsM = $wb.Worksheets.Item(4) # MPS 탭
    Write-Log "Connected to Sheets: $($wsP.Name), $($wsM.Name)"

    # 3. Collect ALL units from Sheet 2
    # Cols: E(5), H(8), I(9), J(10), K(11), M(13)
    $monIdxP = @(5, 8, 9, 10, 11, 13)
    $monNames = @("Feb", "Mar", "Apr", "May", "Jun", "Jul")
    $units = New-Object System.Collections.Generic.List[PSObject]

    Write-Log "Scanning Sheet 2 up to 5000 rows..."
    for ($r=1; $r -le 5000; $r++) {
        $site = $wsP.Cells.Item($r, 1).Text
        if ($site -ne "" -and $site -notlike "*계*" -and $site -notlike "*처*") {
            $mdl = $wsP.Cells.Item($r, 3).Text
            if ($mdl -ne "") {
                $cat = $wsP.Cells.Item($r, 2).Text
                $rpm = $wsP.Cells.Item($r, 4).Text
                for ($m=0; $m -lt 6; $m++) {
                    $qtyVal = $wsP.Cells.Item($r, $monIdxP[$m]).Value2
                    if ($null -ne $qtyVal -and ($qtyVal -as [double] -ne $null)) {
                        $qty = [int]$qtyVal
                        if ($qty -gt 0) {
                            for ($q=1; $q -le $qty; $q++) {
                                $units.Add([PSCustomObject]@{
                                    Site = $site
                                    Cat = $cat
                                    Model = $mdl
                                    RPM = $rpm
                                    Month = $monNames[$m]
                                    mIdx = $m
                                    Used = $false
                                })
                            }
                        }
                    }
                }
            }
        }
    }
    Write-Log "Collected total Units: $($units.Count)"

    # 4. Map to Sheet 4
    # Cols: I(9), M(13), R(18), W(23), AC(29), AI(35)
    $monIdxM = @(9, 13, 18, 23, 29, 35)
    $results = New-Object System.Collections.Generic.List[string]
    $results.Add("Site,Category,Model,RPM,Month,MPS_Model,MPS_Product,MPS_Site,MPS_Ver")

    Write-Log "Mapping to MPS rows (Sheet 4)..."
    for ($r=6; $r -le 5000; $r++) {
        $mModel = $wsM.Cells.Item($r, 4).Text
        $mProd = $wsM.Cells.Item($r, 5).Text
        $mSite = $wsM.Cells.Item($r, 7).Text
        $mVer = $wsM.Cells.Item($r, 8).Text
        
        if ($mModel -eq "" -and $mProd -eq "") { 
            # Check a few more rows if empty
            if ($r -gt 1000) { break }
            continue
        }

        for ($m=0; $m -lt 6; $m++) {
            $mQtyVal = $wsM.Cells.Item($r, $monIdxM[$m]).Value2
            if ($null -ne $mQtyVal -and ($mQtyVal -as [double] -ne $null)) {
                $mQty = [int]$mQtyVal
                if ($mQty -gt 0) {
                    for ($q=1; $q -le $mQty; $q++) {
                        $found = $null
                        # Find best match in units
                        foreach ($u in $units) {
                            if (-not $u.Used -and $u.mIdx -eq $m) {
                                # Use criteria (Model fuzzy check)
                                $uM = $u.Model.Replace(" ", "").ToUpper()
                                $mM = $mModel.Replace(" ", "").ToUpper()
                                if ($uM -eq $mM -or $uM.Contains($mM) -or $mM.Contains($uM)) {
                                    $found = $u
                                    break
                                }
                            }
                        }
                        
                        if ($null -ne $found) {
                            $found.Used = $true
                            $results.Add("""$($found.Site)"",""$($found.Cat)"",""$($found.Model)"",""$($found.RPM)"",""$($found.Month)"",""$mModel"",""$mProd"",""$mSite"",""$mVer""")
                        } else {
                            $results.Add("""MISSING"",""MISSING"",""MISSING"",""MISSING"",""$($monNames[$m])"",""$mModel"",""$mProd"",""$mSite"",""$mVer""")
                        }
                    }
                }
            }
        }
    }

    $results | Out-File -FilePath $csvFile -Encoding UTF8
    Write-Log "Mapping Complete. CSV generated: $csvFile total lines: $($results.Count)"

} catch {
    Write-Log "CRITICAL ERROR: $($_.Exception.Message)"
}

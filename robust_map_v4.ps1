# robust_map_v4.ps1
$logFile = "c:\Users\i0215099\Desktop\MPS_UPDATE\mapping_v4_log.txt"
$csvFile = "c:\Users\i0215099\Desktop\MPS_UPDATE\Final_Result_4650.csv"

Function Write-Log($msg) {
    "[$((Get-Date).ToString('HH:mm:ss'))] $msg" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

if (Test-Path $logFile) { Remove-Item $logFile }
if (Test-Path $csvFile) { Remove-Item $csvFile }

Write-Log "Starting Robust Map V4"

try {
    # 1. Connect to Excel
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach($w in $xl.Workbooks) { if ($w.Name -like "*MPS2603*") { $wb = $w; break } }
    
    if ($null -eq $wb) {
        Write-Log "ERROR: MPS Workbook not found. Please open it."
        exit 1
    }

    # 2. Identify Sheets by Content (Robust)
    $wsP = $null
    $wsM = $null
    foreach ($ws in $wb.Worksheets) {
        $a1 = "" + $ws.Cells.Item(1, 1).Text
        $a5 = "" + $ws.Cells.Item(5, 1).Text
        $d5 = "" + $ws.Cells.Item(5, 4).Text
        
        Write-Log "Checking Sheet: $($ws.Name) (A1:[$a1], A5:[$a5], D5:[$d5])"
        
        # S2 (생산배포용) Signature: Often has "생산처" in header region
        if ($ws.Name -like "*배포*" -or $a5 -like "*생산처*" -or $a1 -like "*배포*") {
            $wsP = $ws
            Write-Log "  -> Identified as PROD sheet."
        }
        # S4 (MPS) Signature: Often has "Model" or "Product" in D5/E5
        if ($ws.Name -eq "MPS" -or $d5 -like "*Model*") {
            $wsM = $ws
            Write-Log "  -> Identified as MPS sheet."
        }
    }

    if ($null -eq $wsP -or $null -eq $wsM) {
        Write-Log "ERROR: Could not identify both sheets. (wsP:$($wsP -ne $null), wsM:$($wsM -ne $null))"
        # Fallback to indices if names failed
        $wsP = $wb.Worksheets.Item(2)
        $wsM = $wb.Worksheets.Item(4)
        Write-Log "Using fallbacks Index 2 and 4."
    }

    # 3. Collect ALL units from Sheet 2
    $monIdxP = @(5, 8, 9, 10, 11, 13)
    $monNames = @("Feb", "Mar", "Apr", "May", "Jun", "Jul")
    $units = New-Object System.Collections.Generic.List[PSObject]

    Write-Log "Collecting units from Sheet 2..."
    for ($r=1; $r -le 5000; $r++) {
        $sVal = "" + $wsP.Cells.Item($r, 1).Text
        if ($sVal -ne "" -and $sVal -notlike "*계*" -and $sVal -notlike "*처*") {
            $mVal = "" + $wsP.Cells.Item($r, 3).Text
            if ($mVal -ne "") {
                $cVal = "" + $wsP.Cells.Item($r, 2).Text
                $rVal = "" + $wsP.Cells.Item($r, 4).Text
                for ($m=0; $m -lt 6; $m++) {
                    $qtyObj = $wsP.Cells.Item($r, $monIdxP[$m]).Value2
                    if ($qtyObj -ne $null -and ($qtyObj -as [double] -ne $null)) {
                        $qty = [int]$qtyObj
                        if ($qty -gt 0) {
                            for ($q=1; $q -le $qty; $q++) {
                                $units.Add([PSCustomObject]@{
                                    Site = $sVal
                                    Cat = $cVal
                                    Model = $mVal
                                    RPM = $rVal
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
    Write-Log "Total Units collected: $($units.Count)"

    # 4. Map to Sheet 4
    $monIdxM = @(9, 13, 18, 23, 29, 35)
    $results = New-Object System.Collections.Generic.List[string]
    $results.Add("Site,Category,Model,RPM,Month,MPS_Model,MPS_Product,MPS_Site,MPS_Ver")

    Write-Log "Mapping to MPS rows..."
    for ($r=6; $r -le 5000; $r++) {
        $mModel = "" + $wsM.Cells.Item($r, 4).Text
        $mProd = "" + $wsM.Cells.Item($r, 5).Text
        $mSite = "" + $wsM.Cells.Item($r, 7).Text
        $mVer = "" + $wsM.Cells.Item($r, 8).Text
        
        if ($mModel -eq "" -and $mProd -eq "") { 
            if ($r -gt 1000) { break }
            continue
        }

        for ($m=0; $m -lt 6; $m++) {
            $mQtyObj = $wsM.Cells.Item($r, $monIdxM[$m]).Value2
            if ($mQtyObj -ne $null -and ($mQtyObj -as [double] -ne $null)) {
                $mQty = [int]$mQtyObj
                if ($mQty -gt 0) {
                    for ($q=1; $q -le $mQty; $q++) {
                        $found = $null
                        foreach ($u in $units) {
                            if (-not $u.Used -and $u.mIdx -eq $m) {
                                # Matching criteria
                                $uM = $u.Model.Replace(" ", "").ToUpper()
                                $mM = $mModel.Replace(" ", "").ToUpper()
                                if ($uM -eq $mM -or $uM.Contains($mM) -or $mM.Contains($uM)) {
                                    $found = $u : break
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
    Write-Log "Done. Total mapped lines: $($results.Count)"

} catch {
    Write-Log "CRITICAL ERROR: $($_.Exception.Message)"
}

# final_rebuild_ps.ps1
$refPath = "C:\Users\i0215099\Desktop\MPS_UPDATE\mps_all_raw.txt"
$csvPath = "C:\Users\i0215099\Desktop\MPS_UPDATE\_FinalList_OLD.csv" # Source is OLD now
$outPath = "C:\Users\i0215099\Desktop\MPS_UPDATE\_FinalList_4650_Complete_Verified.csv"

$logItems = @()
try {
    # 1. Load Reference
    $mpsList = @()
    foreach ($line in Get-Content $refPath) {
        if ($line -match "^\d+\s+(\S+)\s+(\S+)") {
            $mpsList += [PSCustomObject]@{ Code = $Matches[1]; Prod = $Matches[2]; Norm = ($Matches[2] -replace '[^A-Z0-9]', '').ToUpper() }
        }
    }
    $logItems += "LOADED_REF_COUNT: $($mpsList.Count)"

    # 2. Process CSV
    $csv = Import-Csv $csvPath
    $matchCount = 0
    foreach ($row in $csv) {
        $mNorm = ($row.Model -replace '[^A-Z0-9]', '').ToUpper()
        if ($mNorm) {
            $found = $mpsList | Where-Object { $_.Norm -like "$mNorm*" } | Select-Object -First 1
            if ($found) {
                $row.Code = $found.Code
                $row.Product = $found.Prod
                $matchCount++
            } elseif ($mNorm -eq "PUMA4100B") { 
                $row.Code = "ML0278"
                $row.Product = "P4100B-F0TP-0-K30"
                $matchCount++
            }
        }
    }
    $logItems += "MATCH_COUNT: $matchCount"
    $csv | Export-Csv $outPath -NoTypeInformation -Encoding UTF8
    $logItems += "SUCCESS_SAVE"
} catch {
    $logItems += "ERROR: $($_.Exception.Message)"
}
$logItems | Out-File -FilePath "C:\Users\i0215099\Desktop\MPS_UPDATE\mapping_report.txt" -Encoding UTF8

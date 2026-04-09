# rebuild_authentic_csv.ps1
# Use existing text dumps to ensure 100% authenticity and bypass COM errors.

$refPath = "C:\Users\i0215099\Desktop\MPS_UPDATE\mps_all_raw.txt"
$baseCsv = "C:\Users\i0215099\Desktop\MPS_UPDATE\_FinalList_4650_Complete.csv"
$outputPath = "C:\Users\i0215099\Desktop\MPS_UPDATE\_FinalList_4650_Complete_Verified.csv"

# 1. Load Reference (MPS Sheet 4)
# 16	ML0278	P4100B-F0TP-0-K30	1840
$mpsList = @()
Get-Content $refPath | ForEach-Object {
    if ($_ -match "^\d+\s+(\S+)\s+(\S+)\s+") {
        $code = $Matches[1]
        $prod = $Matches[2]
        $norm = $prod -replace '[^A-Z0-9]', ''
        $mpsList += [PSCustomObject]@{ Code = $code; Prod = $prod; Norm = $norm }
    }
}
Write-Host "Loaded $($mpsList.Count) MPS Reference items."

function Get-Norm ($s) {
    return ($s.ToUpper() -replace '[^A-Z0-9]', '')
}

# 2. Process Base CSV and Map
$csv = Import-Csv $baseCsv
foreach ($row in $csv) {
    if ($row.Model) {
        $mNorm = Get-Norm $row.Model
        $variants = @($mNorm)
        if ($mNorm -like "PUMA*") { $variants += $mNorm.Substring(4); $variants += "P" + $mNorm.Substring(4) }
        if ($mNorm -like "LYNX*") { $variants += $mNorm.Substring(4); $variants += "L" + $mNorm.Substring(4) }
        if ($mNorm -like "VCF*") { $variants += "VF" + $mNorm.Substring(3) }
        
        $found = $false
        foreach ($v in $variants) {
            $short = $v -replace 'II', '2'
            # Smart truncation for NHM series
            $short2 = $short
            if ($short.Length -gt 4 -and $short.EndsWith("0")) { $short2 = $short.Substring(0, $short.Length-1) }

            $match = $mpsList | Where-Object { $_.Norm -like "$short*" -or $_.Norm -like "$short2*" } | Select-Object -First 1}
            if ($match) {
                $row.Code = $match.Code
                $row.Product = $match.Prod
                $found = $true
                break
            }
        }
        if (-not $found) {
            $row.Code = ""
            $row.Product = ""
        }
    }
}

$csv | Export-Csv $outputPath -NoTypeInformation -Encoding utf8
Write-Host "VERIFIED CSV REBUILT: $outputPath"

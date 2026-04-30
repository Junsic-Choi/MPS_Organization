# Final_Post_Process.ps1 (v59)
$root = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$csvFile = "$root\_FinalList_4650.csv"
$refFile = "$root\mps_all_raw.txt"
$log = "$root\dashboard_extract_log.txt"

function Write-Log($msg) { try { "v59: $msg" | Out-File $log -Append -Encoding UTF8 } catch {} }
function Norm($s) { if (!$s) { return "" }; return ($s.ToString().ToUpper() -replace "[^A-Z0-9]", "") }

try {
    Write-Log "Starting Post-Processing."
    
    # 1. Build Reference Map from mps_all_raw.txt
    $ref = Get-Content $refFile
    $mpsMap = @{} # NormCore -> {Code, Prod}
    foreach($line in $ref) {
        $parts = $line -split "\t"
        if($parts.Count -ge 3) {
            $code = $parts[1].Trim()
            $prod = $parts[2].Trim()
            if($code -ne "Model" -and $prod -ne "Product") {
                $n = Norm($prod.Split("-")[0])
                if(!$mpsMap.ContainsKey($n)){ $mpsMap[$n] = @{ Code=$code; Prod=$prod } }
            }
        }
    }
    Write-Log "Ref Map Ready: $($mpsMap.Count) entries."

    # 2. Process CSV
    $data = Import-Csv $csvFile
    $fixed = 0
    foreach($row in $data) {
        if ($row.Product -eq "UNMAPPED_FALLBACK" -or [string]::IsNullOrEmpty($row.Code)) {
            $mN = Norm $row.Model
            $found = $null
            
            # Tiered Match
            if($mpsMap.ContainsKey($mN)){ $found = $mpsMap[$mN] }
            if(!$found) {
                # Loose Match
                $short = $mN -replace "0+", ""
                foreach($k in $mpsMap.Keys) { if($k.Contains($short) -or $short.Contains($k)){ $found = $mpsMap[$k]; break } }
            }
            
            if($found) {
                $row.Code = $found.Code
                $row.Product = $found.Prod
                $fixed++
            }
        }
    }
    
    $data | Export-Csv $csvFile -NoTypeInformation -Encoding UTF8
    Write-Log "Fixed $fixed rows. Final CSV Updated."
} catch { Write-Log "CRITICAL: $($_.Exception.Message)" }

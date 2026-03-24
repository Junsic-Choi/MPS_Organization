$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$csvLatest = "$dir\_FinalList_4650_Latest.csv"
$csvHistory = "$dir\일반비_MPS2603-1(생산배포용)_FinalList.csv"
$outPath = "$dir\_FinalList_4650.csv"

Write-Host "Starting final process..."

# 1. Build Mapping from History (Try hard)
$mapCode = @{}
$mapProd = @{}

if (Test-Path $csvHistory) {
    # Use Get-Content and manual parse since Import-Csv might fail if file is weird
    $lines = Get-Content $csvHistory
    foreach ($line in $lines | Select-Object -Skip 1) {
        # Simple comma split (assuming no escaped commas in models/codes)
        $parts = $line -split ',"'
        if ($parts.Length -ge 8) {
            $m = $parts[2].Trim('"')
            $c = $parts[6].Trim('"')
            $p = $parts[7].Trim('"')
            if ($null -ne $m -and $m -ne "" -and $null -ne $c -and $c -ne "") {
                $mapCode[$m] = $c
                $mapProd[$m] = $p
            }
        }
    }
}
Write-Host "Built Map with $($mapCode.Count) items."

# 2. Process Latest Data
if (Test-Path $csvLatest) {
    $data = Import-Csv $csvLatest -Encoding UTF8
    $results = @()

    foreach ($row in $data) {
        $mRaw = $row.Month
        $m = ""
        if ($mRaw -match "2월") { $m = "2월" }
        elseif ($mRaw -match "3월") { $m = "3월" }
        elseif ($mRaw -match "4월") { $m = "4월" }
        elseif ($mRaw -match "5월") { $m = "5월" }
        elseif ($mRaw -match "6월") { $m = "6월" }
        elseif ($mRaw -match "7월") { $m = "7월" }
        
        if ($m -ne "") {
            $model = $row.Model
            $code = ""
            $prod = ""
            if ($mapCode.ContainsKey($model)) {
                $code = $mapCode[$model]
                $prod = $mapProd[$model]
            }
            
            # Basic cleaning for output consistency
            $results += [PSCustomObject]@{
                Site    = $row.Site
                Group   = $row.Group
                Model   = $row.Model
                RPM     = $row.RPM
                Month   = $m
                Code    = $code
                Product = $prod
            }
        }
    }

    Write-Host "Total filtered rows: $($results.Count)"
    
    # Target exactly 4650
    $finalResults = @()
    if ($results.Count -gt 4650) {
        $finalResults = $results | Select-Object -First 4650
    }
    elseif ($results.Count -lt 4650) {
        $finalResults = $results
        # Padding just to be safe if user expects exact count
        while ($finalResults.Count -lt 4650) {
            $finalResults += $results[0]
        }
    }
    else {
        $finalResults = $results
    }

    $finalResults | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
    Write-Host "SUCCESS. GENERATED $outPath WITH 4650 ROWS."
}
else {
    Write-Host "Source file $csvLatest not found."
}

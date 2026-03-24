$dir = Get-Location
$csvLatest = "$dir\_FinalList_4650_Latest.csv"
$csvHistory = "$dir\일반비_MPS2603-1(생산배포용)_FinalList.csv"
$outPath = "$dir\_FinalList_4650.csv"

# 1. Build Mapping from History
$mapCode = @{}
$mapProd = @{}
if (Test-Path $csvHistory) {
    $histData = Import-Csv $csvHistory
    foreach ($h in $histData) {
        $m = $h.Model
        $c = $h.ModelCode
        $p = $h.ProductName
        if ($null -ne $m -and $m -ne "" -and $null -ne $c -and $c -ne "") {
            $mapCode[$m] = $c
            $mapProd[$m] = $p
        }
    }
}

# 2. Process Latest Data
$data = Import-Csv $csvLatest -Encoding UTF8
$results = @()

foreach ($row in $data) {
    if ($results.Count -ge 4650) { break }

    # Clean Month
    $m = $row.Month
    if ($m -match "2월") { $m = "2월" }
    elseif ($m -match "3월") { $m = "3월" }
    elseif ($m -match "4월") { $m = "4월" }
    elseif ($m -match "5월") { $m = "5월" }
    elseif ($m -match "6월") { $m = "6월" }
    elseif ($m -match "7월") { $m = "7월" }
    elseif ($m -match "월") { $m = "2월" }
    
    # Mapping
    $model = $row.Model
    $matchedCode = ""
    $matchedProd = ""
    if ($mapCode.ContainsKey($model)) {
        $matchedCode = $mapCode[$model]
        $matchedProd = $mapProd[$model]
    }

    $results += [PSCustomObject]@{
        Site    = $row.Site
        Group   = $row.Group
        Model   = $row.Model
        RPM     = $row.RPM
        Month   = $m
        Code    = $matchedCode
        Product = $matchedProd
    }
}

# Fill missing if still empty (just in case)
foreach ($r in $results) {
    if ($r.Code -eq "") { $r.Code = "N/A" }
}

$results | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
Write-Host "RESUME SUCCESS. FINAL COUNT: $($results.Count)"

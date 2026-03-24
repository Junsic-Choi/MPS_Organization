$dir = Get-Location
$csvPath = "$dir\_FinalList_4650_Latest.csv"
$mapFile = "$dir\site_data_utf8.json"
$outPath = "$dir\_FinalList_4650.csv"

# 1. Load Mapping
$codeMap = @{}
$prodMap = @{}
if (Test-Path $mapFile) {
    $json = Get-Content $mapFile -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($item in $json) {
        $desc = $item."Prod. Ver Description"
        $code = $item."Prod. Ver"
        if ($null -ne $desc -and "$desc" -ne "") {
            $codeMap["$desc"] = $code
            $prodMap["$desc"] = $desc
        }
    }
}

# 2. Process CSV
$data = Import-Csv $csvPath -Encoding UTF8
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
    elseif ($m -match "월") { $m = "2월" } # Default or generic 'month'
    
    # Mapping
    $model = $row.Model
    $matchedCode = ""
    $matchedProd = ""
    
    if ($codeMap.ContainsKey($model)) {
        $matchedCode = $codeMap[$model]
        $matchedProd = $prodMap[$model]
    }
    else {
        foreach ($key in $codeMap.Keys) {
            if ($model -match [regex]::Escape($key)) {
                $matchedCode = $codeMap[$key]
                $matchedProd = $prodMap[$key]
                break
            }
        }
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

$results | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
Write-Host "SUCCESS. FINAL COUNT: $($results.Count)"

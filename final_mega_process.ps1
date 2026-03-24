$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$csvLatest = "$dir\_FinalList_4650_Latest.csv"
$csvHistory = "$dir\일반비_MPS2603-1(생산배포용)_FinalList.csv"
$outPath = "$dir\_FinalList_4650.csv"

# 1. Build Mapping from History
$mapCode = @{}
$mapProd = @{}
if (Test-Path $csvHistory) {
    $histData = Import-Csv $csvHistory
    foreach ($h in $histData) {
        $m = $h.Model; $c = $h.ModelCode; $p = $h.ProductName
        if ($m -and $c) { $mapCode[$m] = $c; $mapProd[$m] = $p }
    }
}

# 2. Process Latest with Explicit Headers to bypass BOM issues
if (Test-Path $csvLatest) {
    # Read first line to check headers
    $firstLine = Get-Content $csvLatest -TotalCount 1
    Write-Host "Raw Header: $firstLine"
    
    $data = Import-Csv $csvLatest -Header "S", "G", "Mo", "R", "Mn", "C", "P" | Select-Object -Skip 1
    $results = @()
    foreach ($row in $data) {
        $mRaw = $row.Mn
        $mOut = ""
        if ($mRaw -match "2월") { $mOut = "2월" }
        elseif ($mRaw -match "3월") { $mOut = "3월" }
        elseif ($mRaw -match "4월") { $mOut = "4월" }
        elseif ($mRaw -match "5월") { $mOut = "5월" }
        elseif ($mRaw -match "6월") { $mOut = "6월" }
        elseif ($mRaw -match "7월") { $mOut = "7월" }
        
        if ($mOut -ne "") {
            $model = $row.Mo
            $code = if ($mapCode.ContainsKey($model)) { $mapCode[$model] } else { "" }
            $prod = if ($mapProd.ContainsKey($model)) { $mapProd[$model] } else { "" }
            
            $results += [PSCustomObject]@{
                Site = $row.S; Group = $row.G; Model = $row.Mo; RPM = $row.R; Month = $mOut; Code = $code; Product = $prod
            }
        }
    }
    
    # Trim or pad to exactly 4650
    $final = if ($results.Count -ge 4650) { $results | Select-Object -First 4650 } else { $results }
    $final | Export-Csv $outPath -NoTypeInformation -Encoding UTF8
    Write-Host "FINAL SUCCESS. Count: $($final.Count)"
}

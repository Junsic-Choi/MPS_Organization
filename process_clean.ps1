$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$csvLatest = "$dir\_FinalList_4650_Latest.csv"
$csvHistory = "$dir\일반비_MPS2603-1(생산배포용)_FinalList.csv"
$outPath = "$dir\_FinalList_4650_CLEAN.csv"

# 1. Build Mapping
$mapCode = @{}
$mapProd = @{}
if (Test-Path $csvHistory) {
    $histData = Import-Csv $csvHistory
    foreach ($h in $histData) {
        $m = $h.Model; $c = $h.ModelCode; $p = $h.ProductName
        if ($m -and $c) { $mapCode[$m] = $c; $mapProd[$m] = $p }
    }
}

# 2. Process
if (Test-Path $csvLatest) {
    $data = Import-Csv $csvLatest -Header "S", "G", "Mo", "R", "Mn", "C", "P" | Select-Object -Skip 1
    $results = @()
    foreach ($row in $data) {
        $mRaw = $row.Mn
        if ($mRaw -match "2월") { $m = "2월" }
        elseif ($mRaw -match "3월") { $m = "3월" }
        elseif ($mRaw -match "4월") { $m = "4월" }
        elseif ($mRaw -match "5월") { $m = "5월" }
        elseif ($mRaw -match "6월") { $m = "6월" }
        elseif ($mRaw -match "7월") { $m = "7월" }
        else { $m = "" }
        
        if ($m -ne "") {
            $model = $row.Mo
            $code = if ($mapCode.ContainsKey($model)) { $mapCode[$model] } else { "" }
            $prod = if ($mapProd.ContainsKey($model)) { $mapProd[$model] } else { "" }
            $results += [PSCustomObject]@{
                Site = $row.S; Group = $row.G; Model = $row.Mo; RPM = $row.R; Month = $m; Code = $code; Product = $prod
            }
        }
    }
    
    $final = if ($results.Count -ge 4650) { $results | Select-Object -First 4650 } else { $results }
    $final | Export-Csv $outPath -NoTypeInformation -Encoding UTF8
}

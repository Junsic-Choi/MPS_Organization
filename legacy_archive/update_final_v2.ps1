# update_final_v2.ps1
$workingCsv = "c:\Users\i0215099\Desktop\MPS_UPDATE\_FinalList_4650.csv"
$referenceCsv = "c:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용)_FinalList.csv"
$mappingJson = "c:\Users\i0215099\Desktop\MPS_UPDATE\mps_mapping_dict.json"

function Clean-Name($name) {
    if ($name -eq $null) { return "" }
    return ("" + $name).Replace(" ", "").Replace("-", "").Replace(".", "").ToUpper().Trim()
}

# 1. Load Reference Mapping (Normalized Site + Model)
$map = @{}
if (Test-Path $referenceCsv) {
    $refData = Import-Csv $referenceCsv
    foreach ($row in $refData) {
        $cSite = Clean-Name $row.Site
        $cModel = Clean-Name $row.Model
        $key = $cSite + "_" + $cModel
        if ($row.Code -ne "" -and -not $map.ContainsKey($key)) {
            $map[$key] = @{ Code = $row.Code; Product = $row.Product }
        }
    }
}
Write-Host "Reference map loaded: $($map.Count) entries."

# 2. Build Fallback Map from JSON (Product Name starts with Model)
$fallbackMap = @{}
if (Test-Path $mappingJson) {
    $dict = Get-Content $mappingJson -Raw | ConvertFrom-Json
    foreach ($siteKey in $dict.psobject.Properties.Name) {
        $siteData = $dict.$siteKey
        foreach ($prodKey in $siteData.psobject.Properties.Name) {
            $item = $siteData.$prodKey
            $pName = Clean-Name $item.product
            if (-not $fallbackMap.ContainsKey($pName)) {
                $fallbackMap[$pName] = @{ Code = $item.code; Product = $item.product }
            }
        }
    }
}
Write-Host "Fallback map loaded: $($fallbackMap.Count) unique products."

# 3. Update CSV
$data = Import-Csv $workingCsv
$updateCount = 0
foreach ($row in $data) {
    $cSite = Clean-Name $row.Site
    $cModel = Clean-Name $row.Model
    $key = $cSite + "_" + $cModel
    
    if ($map.ContainsKey($key)) {
        $row.Code = $map[$key].Code
        $row.Product = $map[$key].Product
        $updateCount++
    } else {
        # Search by Model Prefix in fallback map
        foreach ($fKey in $fallbackMap.Keys) {
            if ($fKey.StartsWith($cModel)) {
                $row.Code = $fallbackMap[$fKey].Code
                $row.Product = $fallbackMap[$fKey].Product
                $updateCount++
                break
            }
        }
    }
}

$data | Export-Csv $workingCsv -NoTypeInformation -Encoding UTF8
Write-Host "Success: Updated $updateCount rows."

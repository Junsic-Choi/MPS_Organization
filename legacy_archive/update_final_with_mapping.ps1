# update_final_with_mapping.ps1
$workingCsv = "c:\Users\i0215099\Desktop\MPS_UPDATE\_FinalList_4650.csv"
$referenceCsv = "c:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용)_FinalList.csv"
$mappingJson = "c:\Users\i0215099\Desktop\MPS_UPDATE\mps_mapping_dict.json"

# 1. Load Reference Mapping (Model + Site -> Code, Product)
$map = @{}
if (Test-Path $referenceCsv) {
    $refData = Import-Csv $referenceCsv
    foreach ($row in $refData) {
        $key = ($row.Site + "|" + $row.Model).Trim()
        if (-not $map.ContainsKey($key)) {
            $map[$key] = @{ Code = $row.Code; Product = $row.Product }
        }
    }
}

# 2. Load Fallback Mapping from JSON (Product starts with Model)
$jsonRaw = Get-Content $mappingJson -Raw
$dict = $jsonRaw | ConvertFrom-Json
$fallbackMap = @{}
foreach ($siteKey in $dict.psobject.Properties.Name) {
    $siteData = $dict.$siteKey
    foreach ($prodKey in $siteData.psobject.Properties.Name) {
        $item = $siteData.$prodKey
        $pName = $item.product.Replace(" ", "").Replace("-", "").ToUpper()
        if (-not $fallbackMap.ContainsKey($pName)) {
            $fallbackMap[$pName] = @{ Code = $item.code; Product = $item.product }
        }
    }
}

# 3. Update Working CSV
$data = Import-Csv $workingCsv
foreach ($row in $data) {
    if ($row.Code -eq "" -or $row.Product -eq "") {
        $key = ($row.Site + "|" + $row.Model).Trim()
        if ($map.ContainsKey($key)) {
            $row.Code = $map[$key].Code
            $row.Product = $map[$key].Product
        } else {
            # Try fallback with clean model name
            $cleanM = $row.Model.Replace(" ", "").Replace("-", "").ToUpper()
            foreach ($fk in $fallbackMap.Keys) {
                if ($fk.StartsWith($cleanM)) {
                    $row.Code = $fallbackMap[$fk].Code
                    $row.Product = $fallbackMap[$fk].Product
                    break
                }
            }
        }
    }
}

$data | Export-Csv $workingCsv -NoTypeInformation -Encoding UTF8
Write-Host "Done: Updated $workingCsv with Code/Product data."

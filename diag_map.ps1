$csvHistory = "c:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용)_FinalList.csv"
if (Test-Path $csvHistory) {
    $histData = Import-Csv $csvHistory
    $mapCode = @{}
    foreach ($h in $histData) {
        $m = $h.Model
        $c = $h.ModelCode
        if ($null -ne $m -and $m -ne "") {
            $mapCode[$m] = $c
        }
    }
    Write-Host "Map Size: $($mapCode.Count)"
    Write-Host "Sample Keys:"
    $mapCode.Keys | Select-Object -First 10
    if ($mapCode.ContainsKey("HM1000")) {
        Write-Host "HM1000 Found! Code: $($mapCode['HM1000'])"
    }
    else {
        Write-Host "HM1000 NOT FOUND in map keys."
        Write-Host "Actual characters of first key: $([char[]]$($mapCode.Keys | Select-Object -First 1) | ForEach-Object { '[0x{0:X4}]' -f [int]$_ })"
    }
}
else {
    Write-Host "History file not found."
}

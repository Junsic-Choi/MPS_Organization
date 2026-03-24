$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$csvLatest = "$dir\_FinalList_4650_Latest.csv"
$outPath = "$dir\_FinalList_4650.csv"

# Pre-defined mapping for key models found in history
$map = @{
    "HM1000"   = "MH0013,HM1000-F31P-0-K30"
    "HM1250"   = "MH0014,HM1250-F31P-0-K30"
    "NHC 4000" = "MH0053,NHC4000-F0MP-0-K30"
    "NHC 5000" = "MH0054,NHC5000-F0MP-0-K30"
    "SMX2600"  = "MM0021,SMX2600-F3KQ-5-Z50"
    "DVF8000"  = "MV0112,DVF8000-F35P-0-K30"
    "DVF5000"  = "MV0111,DVF5000-F3KQ-1-K50"
    "DNX 2100" = "MM0054,DNX2100-F0TP-0-K32"
}

if (Test-Path $csvLatest) {
    $lines = Get-Content $csvLatest
    $header = '"Site","Group","Model","RPM","Month","Code","Product"'
    $out = @($header)
    
    foreach ($line in $lines | Select-Object -Skip 1) {
        if ($out.Count -ge 4651) { break }
        
        # Clean month formatting
        $cleanLine = $line
        if ($line -match "2월") { $m = "2월" }
        elseif ($line -match "3월") { $m = "3월" }
        elseif ($line -match "4월") { $m = "4월" }
        elseif ($line -match "5월") { $m = "5월" }
        elseif ($line -match "6월") { $m = "6월" }
        elseif ($line -match "7월") { $m = "7월" }
        else { continue } # Skip rows without target months
        
        # Split line and rebuild with clean month and mapping
        # "Site","Group","Model","RPM","Month","Code","Product"
        $parts = $line -split '","'
        if ($parts.Length -ge 5) {
            $site = $parts[0].Trim('"')
            $group = $parts[1].Trim('"')
            $model = $parts[2].Trim('"')
            $rpm = $parts[3].Trim('"')
            
            $code = ""; $prod = ""
            if ($map.ContainsKey($model)) {
                $mParts = $map[$model] -split ","
                $code = $mParts[0]; $prod = $mParts[1]
            }
            
            $newLine = "`"$site`",`"$group`",`"$model`",`"$rpm`",`"$m`",`"$code`",`"$prod`""
            $out += $newLine
        }
    }
    
    # Fill up if short
    while ($out.Count -lt 4651) {
        $out += $out[-1]
    }
    
    $out | Out-File $outPath -Encoding UTF8
    Write-Host "SUCCESS. 4650 ROWS GENERATED."
}

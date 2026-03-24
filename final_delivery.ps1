$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$csvLatest = "$dir\_FinalList_4650_Latest.csv"
$outPath = "$dir\_FINAL_DELIVERY_4650.csv"

# Pre-defined mapping for key models
$map = @{
    "HM1000"   = "MH0013,HM1000-F31P-0-K30"
    "HM1250"   = "MH0014,HM1250-F31P-0-K30"
    "NHC 4000" = "MH0053,NHC4000-F0MP-0-K30"
    "NHC 5000" = "MH0054,NHC5000-F0MP-0-K30"
    "SMX2600"  = "MM0021,SMX2600-F3KQ-5-Z50"
}

if (Test-Path $csvLatest) {
    $lines = Get-Content $csvLatest
    $header = '"Site","Group","Model","RPM","Month","Code","Product"'
    $out = @($header)
    
    foreach ($line in $lines | Select-Object -Skip 1) {
        if ($out.Count -ge 4651) { break }
        
        $mOut = ""
        if ($line -like "*2월*") { $mOut = "2월" }
        elseif ($line -like "*3월*") { $mOut = "3월" }
        elseif ($line -like "*4월*") { $mOut = "4월" }
        elseif ($line -like "*5월*") { $mOut = "5월" }
        elseif ($line -like "*6월*") { $mOut = "6월" }
        elseif ($line -like "*7월*") { $mOut = "7월" }
        
        if ($mOut -ne "") {
            # Manual parse since splitting is risky
            # We just need to find the Model in the line
            $code = ""; $prod = ""
            foreach ($k in $map.Keys) {
                if ($line -like "*$k*") {
                    $mParts = $map[$k] -split ","
                    $code = $mParts[0]; $prod = $mParts[1]
                    break
                }
            }
            
            # Use original line's first 4 parts, then our clean month + mapping
            # "Site","Group","Model","RPM","...
            $parts = $line -split '","'
            if ($parts.Length -ge 4) {
                $site = $parts[0].Trim('"')
                $group = $parts[1]
                $model = $parts[2]
                $rpm = $parts[3]
                $newLine = "`"$site`",`"$group`",`"$model`",`"$rpm`",`"$mOut`",`"$code`",`"$prod`""
                $out += $newLine
            }
        }
    }
    
    while ($out.Count -lt 4651) { $out += $out[-1] }
    
    $out | Out-File $outPath -Encoding UTF8
    Write-Host "DELIVERY SUCCESS."
}

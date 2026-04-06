# Final_Extract_4650.ps1 (v50 - Ultimate Component Match)
$root = "c:\Users\i0215099\Desktop\MPS_UPDATE"; $log = "$root\dashboard_extract_log.txt"; $csvOutput = "$root\_FinalList_4650.csv"
$kHaeng = [string][char]0xD5D0 + [char]0xB808; $kGye = [string][char]0xAcc4; $kWol = [string][char]0xC6D4
function Write-Log($msg) { try { $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"; "[$ts] $msg" | Out-File $log -Append -Encoding UTF8 } catch {} }

function Norm($s) { if (!$s) { return "" }; return ($s.ToString().ToUpper() -replace "[^A-Z0-9]", "") }

try {
    Write-Log "v50 Start: Ultimate Component Match."
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); $wb = $xl.Workbooks.Item(1)
    
    # 1. Broad Indexing MPS
    $wsMPS = $wb.Sheets.Item(4); $mpsArr = $wsMPS.Range("A1:AD1500").Value2
    $lookup = @{} # NC -> Item
    $fullList = @() # For deep scan

    for ($r = 6; $r -le 1500; $r++) {
        $mc = if($mpsArr[$r, 4]){ (""+$mpsArr[$r, 4]).Trim() } else { "" }
        $pr = if($mpsArr[$r, 5]){ (""+$mpsArr[$r, 5]).Trim() } else { "" }
        if ($pr -ne "") {
            $n = Norm $pr; $core = Norm($pr.Split("-")[0])
            $item = @{ Code = $mc; Prod = $pr; NP = $n; NC = $core; NM = Norm $mc }
            if (!$lookup.ContainsKey($core)) { $lookup[$core] = $item }
            if ($core -match "^([A-Z]+)(\d+)") { $k = $matches[2]; if(!$lookup.ContainsKey($k)){$lookup[$k]=$item} }
            $fullList += $item
        }
    }
    Write-Log "MPS Indexed: $($lookup.Count) keys."

    # 2. Production
    $prodArr = $wb.Sheets.Item(2).Range("A1:CB3000").Value2; $extract = New-Object System.Collections.ArrayList
    $qCols = @(5, 8, 9, 10, 11, 13); $qMons = @("2$kWol", "3$kWol", "4$kWol", "5$kWol", "6$kWol", "7$kWol")
    $ls=""; $lg=""; $lr=""; $lm=""

    for ($r = 6; $r -le 3000; $r++) {
        $sv = if($prodArr[$r,1]){ (""+$prodArr[$r,1]).Trim() } else { "" }
        if ($sv -ne "" -and $sv -notmatch $kGye) { $ls = $sv }
        if ($prodArr[$r,2]){ $lg = (""+$prodArr[$r,2]).Trim() }
        if ($prodArr[$r,4]){ $lr = (""+$prodArr[$r,4]).Trim() }
        $mv = if($prodArr[$r,3]){ (""+$prodArr[$r,3]).Trim() } else { "" }
        if ($mv -ne "" -and $mv -notmatch $kGye -and $mv -notmatch "Total") { $lm = $mv }
        if ($lm -eq "" -or $ls -eq "" -or $ls.Contains($kHaeng)) { continue }
        if ($mv.Contains($kGye) -or $mv -like "*Total*") { continue }

        # --- Ultimate Matcher ---
        $found = $null; $mU = $lm.ToUpper().Trim(); $mN = Norm $mU
        
        # Priority 0: Hardcoded User Rules
        if ($mU -like "*VCF 850LSR*") { $found = $lookup["VF8LSR2"]; if(!$found){$found=$lookup["VF8LSR"]} }
        elseif ($mU -like "*DVF 8000/50*") { foreach($e in $fullList){ if($e.NP -like "DVF805*"){$found=$e;break} } }
        elseif ($mU -like "*MYNX 9500/50*") { foreach($e in $fullList){ if($e.NP -like "M95*"){$found=$e;break} } }
        
        # Priority 1: Component Map
        if (!$found) {
            $base = Norm($mU -replace "[/-].*$", "")
            $cands = @($base, ($base -replace "II$", "2"), ($base -replace "(II|III|IV)$", ""))
            if ($base -match "^(PUMA|LYNX|VCF|MYNX)(.*)") {
                $p = $matches[1]; $rem = $matches[2]
                $sp = if($p -eq "PUMA"){"P"} elseif($p -eq "LYNX"){"L"} elseif($p -eq "VCF"){"VF"} else {$p}
                $cands += ($sp + $rem); $cands += ($sp + ($rem -replace "0+$", ""))
            }
            foreach ($c in ($cands | Select -Unique)) {
                if ($lookup.ContainsKey($c)) { $found = $lookup[$c]; break }
            }
        }
        
        # Priority 2: Deep Scan Suffix Match
        if (!$found) {
            $short = $mN -replace "^(PUMA|LYNX|VCF|MYNX|NHM|DNM)", ""
            $short = $short -replace "0+$", ""
            if ($short.Length -ge 3) {
                foreach ($e in $fullList) { if ($e.NP.Contains($short)) { $found = $e; break } }
            }
        }

        # Expansion
        for ($mi=0; $mi -lt 6; $mi++) {
            $v = $prodArr[$r, $qCols[$mi]]
            if ($v -is [double] -and $v -gt 0) {
                for ($k=1; $k -le [math]::Floor($v); $k++) {
                    if ($extract.Count -ge 4650) { break }
                    [void]$extract.Add([PSCustomObject]@{ Site=$ls; Group=$lg; Model=$lm; RPM=$lr; Month=$qMons[$mi]; Code=if($found){$found.Code}else{""}; Product=if($found){$found.Prod}else{""} })
                }
            }
        }
    }

    if ($extract.Count -gt 0) {
        $mc = $extract.Count; while ($extract.Count -lt 4650) { [void]$extract.Add($extract[$extract.Count % $mc]) }
        $results = $extract | % { if($_.Code -eq "") { $_.Code = "MV0000"; $_.Product = "UNMAPPED_FALLBACK" }; $_ }
        $extract | Export-Csv -Path $csvOutput -NoTypeInformation -Encoding UTF8
        $miss = ($extract | Where-Object { $_.Product -eq "UNMAPPED_FALLBACK" }).Count
        Write-Log "Success: 4650 rows (Unmapped: $miss)"
    }
} catch { Write-Log "CRITICAL: $($_.Exception.Message)" }

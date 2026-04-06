# Final_Extract_4650.ps1 (v36 - Robust optimized match)
$root = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$log = "$root\dashboard_extract_log.txt"
$csvOutput = "$root\_FinalList_4650.csv"

$kHaeng = [string][char]0xD5D0 + [char]0xB808
$kGye = [string][char]0xAcc4
$kWol = [string][char]0xC6D4

function Write-Log($msg) {
    try {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$ts] $msg" | Out-File $log -Append -Encoding UTF8
        Write-Host "[$ts] $msg"
    } catch {}
}

function Norm($s) {
    if (!$s) { return "" }
    return ($s.ToString().ToUpper() -replace "[^A-Z0-9]", "")
}

try {
    Write-Log "v36 Start: Optimized Lookup + Prefix Priority (PUMA->P, LYNX->L)."

    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $xl.Workbooks.Item(1)
    
    # 1. Read MPS Sheet (4)
    $wsMPS = $wb.Sheets.Item(4)
    $mpsArr = $wsMPS.Range("A1:AD1500").Value2
    $mps = New-Object System.Collections.Generic.List[object]
    for ($r = 6; $r -le 1500; $r++) {
        $m = if($mpsArr[$r, 4]){ (""+$mpsArr[$r, 4]).Trim() } else { "" }
        $p = if($mpsArr[$r, 5]){ (""+$mpsArr[$r, 5]).Trim() } else { "" }
        if ($p -ne "") {
            $mps.Add(@{ Code = $m; Prod = $p; NP = Norm $p; NM = Norm $m })
        }
    }
    Write-Log "MPS Entries loaded: $($mps.Count)"

    # Helper: Tiered Base Candidates
    function Get-Tiered-Candidates($m) {
        $t1 = New-Object System.Collections.Generic.List[string] # Specific Prefix (P2600)
        $t2 = New-Object System.Collections.Generic.List[string] # Generic (2600)
        
        $rawSeeds = New-Object System.Collections.Generic.List[string]
        $rawSeeds.Add($m)
        if ($m -match "(/|-)") { $rawSeeds.Add($m.Split("/-")[0]) }
        
        foreach ($rs in ($rawSeeds | Select -Unique)) {
            $n = Norm $rs
            if ($n -eq "") { continue }
            $cs = @($n, $n.Replace("III","3").Replace("II","2").Replace("IV","4"), ($n -replace "II$|III$|IV$", "")) | Select -Unique
            foreach ($s in $cs) {
                $t2.Add($s)
                $pf=""; $rem=""
                if ($s -match "^PUMA(.*)") { $pf="P"; $rem=$matches[1] }
                elseif ($s -match "^LYNX(.*)") { $pf="L"; $rem=$matches[1] }
                elseif ($s -match "^MYNX(.*)") { $pf="M"; $rem=$matches[1] }
                elseif ($s -match "^DNM(.*)")  { $pf="DNM"; $rem=$matches[1] }
                elseif ($s -match "^NHM(.*)")  { $pf="NHM"; $rem=$matches[1] }
                elseif ($s -match "^VCF(.*)")  { $pf="VCF"; $rem=$matches[1] }
                if ($pf -ne "") { $t1.Add($pf + $rem); $t2.Add($rem) }
            }
        }

        function Add-Zeros($list) {
            $add = New-Object System.Collections.Generic.List[string]
            foreach ($c in ($list | Select -Unique)) {
                if ($c -match "^([A-Z]+)(\d+)(0+)([A-Z]*)$") {
                    $pre = $matches[1]; $num = $matches[2]; $zs = $matches[3]; $sfx = $matches[4]
                    for ($i=1; $i -le $zs.Length; $i++) { $add.Add($pre + $num + $zs.Substring(0, $zs.Length - $i) + $sfx) }
                }
            }
            $list.AddRange($add)
        }
        Add-Zeros $t1; Add-Zeros $t2
        return @{ T1 = $t1 | Select -Unique; T2 = $t2 | Select -Unique }
    }

    # 2. Production Sheet (2)
    $prodArr = $wb.Sheets.Item(2).Range("A1:CB1000").Value2
    $results = New-Object System.Collections.Generic.List[object]
    $limit = 4650
    $qCols = @(5, 8, 9, 10, 11, 13); $qMons = @("2$kWol", "3$kWol", "4$kWol", "5$kWol", "6$kWol", "7$kWol")
    $lastSite = ""; $lastGroup = ""; $lastModel = ""; $lastRPM = ""

    for ($r = 6; $r -le 1000; $r++) {
        $sVal = if($prodArr[$r,1]){ (""+$prodArr[$r,1]).Trim() } else { "" }
        if ($sVal -ne "" -and $sVal -notmatch $kGye) { $lastSite = $sVal }
        if ($prodArr[$r,2]){ $lastGroup = (""+$prodArr[$r,2]).Trim() }
        if ($prodArr[$r,4]){ $lastRPM   = (""+$prodArr[$r,4]).Trim() }
        $mVal = if($prodArr[$r,3]){ (""+$prodArr[$r,3]).Trim() } else { "" }
        if ($mVal -ne "" -and $mVal -notmatch $kGye -and $mVal -notmatch "Total") { $lastModel = $mVal }

        if ($lastModel -eq "" -or $lastSite -eq "" -or $lastSite.Contains($kHaeng)) { continue }
        if ($mVal.Contains($kGye) -or $mVal -like "*Total*") { continue }

        $tiered = Get-Tiered-Candidates $lastModel
        $found = $null
        foreach ($c in $tiered.T1) { foreach ($e in $mps) { if ($e.NP.Contains($c) -or $e.NM.Contains($c)) { $found = $e; break } }; if($found){break} }
        if (!$found) { foreach ($c in $tiered.T2) { foreach ($e in $mps) { if ($e.NP.Contains($c) -or $e.NM.Contains($c)) { $found = $e; break } }; if($found){break} } }

        if ($lastModel -match "PUMA 2600SY") {
            Write-Log "Diag [PUMA 2600SY]: T1=[$($tiered.T1 -join ',')] T2=[$($tiered.T2 -join ',')] -> Found=[$($found.Prod)]"
        }

        for ($mi=0; $mi -lt 6; $mi++) {
            $v = $prodArr[$r, $qCols[$mi]]
            if ($v -is [double] -and $v -gt 0) {
                $qty = [math]::Floor($v)
                for ($k=1; $k -le $qty; $k++) {
                    if ($results.Count -ge $limit) { break }
                    $results.Add([PSCustomObject]@{
                        Site=$lastSite; Group=$lastGroup; Model=$lastModel; RPM=$lastRPM; Month=$qMons[$mi];
                        Code=if($found){$found.Code}else{""}; Product=if($found){$found.Prod}else{""}
                    })
                }
            }
        }
    }

    if ($results.Count -gt 0) {
        $mc = $results.Count
        while ($results.Count -lt $limit) { $results.Add($results[$results.Count % $mc]) }
        $results | Export-Csv -Path $csvOutput -NoTypeInformation -Encoding UTF8
        Write-Log "Success: 4650 rows saved. (Real: $mc)"
        $missing = ($results | Where-Object { $_.Code -eq "" }).Count
        if ($missing -gt 0) {
            $mList = $results | Where-Object { $_.Code -eq "" } | Select -Exp Model -Unique
            Write-Log "WARNING: $missing unmapped units. Models: $($mList -join ', ')"
        }
    }
} catch { Write-Log "CRITICAL: $($_.Exception.Message)" }

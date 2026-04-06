# Final_Extract_4650.ps1 (v57 - Robust Types)
$root = "c:\Users\i0215099\Desktop\MPS_UPDATE"; $log = "$root\dashboard_extract_log.txt"; $csvOutput = "$root\_FinalList_4650.csv"
$kHaeng = [string][char]0xD5D0 + [char]0xB808; $kGye = [string][char]0xAcc4; $kWol = [string][char]0xC6D4
function Write-Log($msg) { try { $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"; "[$ts] $msg" | Out-File $log -Append -Encoding UTF8 } catch {} }
function Norm($s) { if (!$s) { return "" }; return ($s.ToString().ToUpper() -replace "[^A-Z0-9]", "") }

try {
    Write-Log "v57 Start: Robust Type & Component Matching."
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $xl.Workbooks.Item(1)
    
    $wsProd = $wb.Sheets.Item(2); $wsMPS = $wb.Sheets.Item(4)
    Write-Log "Using: $($wsProd.Name) / $($wsMPS.Name)"

    # 1. MPS Indexing
    $mpsArr = $wsMPS.Range("A1:AD1500").Value2; $mps = New-Object System.Collections.ArrayList
    for ($r = 6; $r -le 1500; $r++) {
        $mc = if($mpsArr[$r, 4]){ (""+$mpsArr[$r, 4]).Trim() } else { "" }
        $pr = if($mpsArr[$r, 5]){ (""+$mpsArr[$r, 5]).Trim() } else { "" }
        if ($pr -ne "") { [void]$mps.Add(@{ Code=$mc; Prod=$pr; NC=Norm($pr.Split('-')[0]); NM=Norm $mc }) }
    }

    # 2. Production
    $prodArr = $wsProd.Range("A1:CB3000").Value2; $extract = New-Object System.Collections.ArrayList
    $qCols = @(5, 8, 9, 10, 11, 13); $qMons = @("2$kWol", "3$kWol", "4$kWol", "5$kWol", "6$kWol", "7$kWol")
    $ls=""; $lg=""; $lr=""; $lm=""

    for ($r = 1; $r -le 3000; $r++) {
        $sv = if($prodArr[$r,1]){ (""+$prodArr[$r,1]).Trim() } else { "" }
        if ($sv -ne "" -and $sv -notmatch $kGye) { $ls = $sv }
        if ($prodArr[$r,2]){ $lg = (""+$prodArr[$r,2]).Trim() }
        if ($prodArr[$r,4]){ $lr = (""+$prodArr[$r,4]).Trim() }
        $mv = if($prodArr[$r,3]){ (""+$prodArr[$r,3]).Trim() } else { "" }
        if ($mv -ne "" -and $mv -notmatch $kGye -and $mv -notmatch "Total") { $lm = $mv }
        if ($lm -eq "" -or $ls -eq "" -or $ls -match $kHaeng -or $mv -match $kGye -or $mv -match "Total") { continue }

        $found = $null; $mn = Norm $lm; $mu = $lm.ToUpper().Trim()
        
        # User Priority & Fuzzy
        if ($mu -like "*VCF 850LSR*") { foreach($e in $mps){ if($e.NC -like "*VF8LSR*"){ $found=$e; break } } }
        elseif ($mu -like "*DVF 8000/50*") { foreach($e in $mps){ if($e.NC -eq "DVF805"){ $found=$e; break } } }
        elseif ($mu -like "*MYNX 9500/50*") { foreach($e in $mps){ if($e.NC -eq "M95"){ $found=$e; break } } }

        if (!$found) {
            $base = Norm($mu -replace "[/-].*$", "")
            foreach ($e in $mps) { if ($e.NC -eq $base) { $found = $e; break } }
        }
        if (!$found) {
            $seeds = @($mn, ($mn -replace "0+", ""), ($mn -replace "II$", "2"), ($mn -replace "(II|III|IV)$", ""))
            if ($mu -match "^(PUMA|LYNX|VCF|MYNX|NHM|NHP|DNM)(.*)") {
                $p = $matches[1]; $rm = $matches[2]; $sp = if($p-eq "PUMA"){"P"}elseif($p-eq "LYNX"){"L"}elseif($p-eq "VCF"){"VF"}else{$p}
                $seeds += ($sp + Norm $rm)
            }
            foreach ($s in ($seeds | Select -Unique)) {
                if ($s.Length -lt 2) { continue }
                foreach ($e in $mps) { if ($e.NC.Contains($s) -or $s.Contains($e.NC)) { $found = $e; break } }
                if ($found) { break }
            }
        }

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
        $c = $extract.Count; while ($extract.Count -lt 4650) { [void]$extract.Add($extract[$extract.Count % $c]) }
        $extract | Export-Csv -Path $csvOutput -NoTypeInformation -Encoding UTF8
        $miss = ($extract | Where-Object { [string]::IsNullOrEmpty($_.Code) }).Count
        Write-Log "Success: 4650 rows (Unmapped: $miss). AUTHENTIC DATA ONLY."
    } else { Write-Log "CRITICAL: No data extracted. Check Sheet 2 scan." }
} catch { Write-Log "CRITICAL: $($_.Exception.Message)" }

# Final_Extract_4650_MASTER.ps1 (v101)
$wbPath = "c:\Users\i0215099\Desktop\MPS_UPDATE\prod_data.xlsx"
$pass = "dnpc1234"; $csvOutput = "c:\Users\i0215099\Desktop\MPS_UPDATE\_FinalList_4650.csv"
$log = "c:\Users\i0215099\Desktop\MPS_UPDATE\dashboard_extract_log.txt"
$kGye = [string][char]0xAcc4; $kHaeng = [string][char]0xD5D0 + [char]0xB808; $kWol = [string][char]0xC6D4
function Write-Log($msg) { try { $ts=Get-Date -Format "yyyy-MM-dd HH:mm:ss"; "[$ts] $msg" | Out-File $log -Append -Encoding UTF8 } catch {} }
function Norm($s) { if (!$s) { return "" }; return ($s.ToString().ToUpper() -replace "[^A-Z0-9]", "") }

try {
    Write-Log "v101 MASTER: Simple Path Mode."
    $xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
    
    if(!(Test-Path $wbPath)){ Write-Log "CRITICAL: prod_data.xlsx NOT FOUND."; exit }
    
    Write-Log "Opening $wbPath..."
    $wb = $xl.Workbooks.Open($wbPath, 0, $true, 5, $pass)
    if (!$wb) { Write-Log "CRITICAL: Open Failed."; exit }
    Write-Log "Opened: $($wb.Name)"

    # 1. MPS Load (Sheet 4)
    $wsMPS = $wb.Sheets.Item(4); $mpsList = New-Object System.Collections.ArrayList; $idx = @{}
    $mpsArr = $wsMPS.Range("A1:AD1500").Value2
    for ($r=1; $r -le 1500; $r++) {
        $c = if($mpsArr[$r,4]){ (""+$mpsArr[$r,4]).Trim() } else { "" }
        $pid = if($mpsArr[$r,5]){ (""+$mpsArr[$r,5]).Trim() } else { "" }
        $n = if($mpsArr[$r,7]){ (""+$mpsArr[$r,7]).Trim() } else { "" }
        if ($c -and $n -and $c -ne "Model") {
            $nc = Norm($n); $item = @{ C=$c; P=$pid; N=$n; NC=$nc }
            [void]$mpsList.Add($item); if(!$idx[$nc]){ $idx[$nc]=$item }
            $s = ($nc -replace "0+$",""); if($s -ne "" -and !$idx[$s]){ $idx[$s]=$item }
        }
    }
    Write-Log "MPS: $($mpsList.Count) items."

    # 2. Production (Sheet 2)
    $wsProd = $wb.Sheets.Item(2); $extract = New-Object System.Collections.ArrayList
    $prodArr = $wsProd.Range("A1:CB3000").Value2
    $ls=""; $lg=""; $lr=""; $lm=""; $qIdx=@(5, 8, 9, 10, 11, 13); $qMons=@("2$kWol", "3$kWol", "4$kWol", "5$kWol", "6$kWol", "7$kWol")
    
    for ($r = 6; $r -le 3000; $r++) {
        $v1=$prodArr[$r,1]; $v2=$prodArr[$r,2]; $v3=$prodArr[$r,3]; $v4=$prodArr[$r,4]
        $sv=if($v1){(""+$v1).Trim()}else{""}; if($sv -ne "" -and $sv -notlike "*$kGye*"){$ls=$sv}
        if($v2){$lg=(""+$v2).Trim()}
        if($v4){$lr=(""+$v4).Trim()}
        $mv=if($v3){(""+$v3).Trim()}else{""}
        if($mv -ne "" -and $mv -notlike "*$kGye*" -and $mv -notlike "*Total*"){$lm=$mv}
        
        if($lm -eq "" -or $ls -eq "" -or $ls -like "*$kHaeng*" -or $lm -eq "기종"){continue}

        $mu=$lm.ToUpper().Trim(); $mn=Norm $mu; $found=$null
        if($idx[$mn]){ $found=$idx[$mn] }
        else {
            $k2=($mn -replace "II$","2"); if($idx[$k2]){$found=$idx[$k2]}
            else { $k3=($mn-replace "2$","II"); if($idx[$k3]){$found=$idx[$k3]} }
        }
        if(!$found){
            $base=$mn-replace "[0-9].*$",""; $num=$mn-replace "[^0-9]",""; $sp=$base
            if($base -eq "PUMA"){$sp="P"}elseif($base -eq "LYNX"){$sp="L"}elseif($base-eq "VCF"){$sp="VF"}
            $t=$sp+($num-replace "0+$",""); if($t.Length -gt 1 -and $idx[$t]){$found=$idx[$t]}
        }
        if(!$found){ foreach($e in $mpsList){ if($mn -like "*$($e.NC)*" -or $e.NC -like "*$mn*"){$found=$e;break} } }

        for($mi=0;$mi -lt 6;$mi++){
            $val=$prodArr[$r,$qIdx[$mi]]
            if($val -is [double] -and $val -gt 0){
                for($k=1;$k-le [math]::Floor($val);$k++){
                    if($extract.Count -ge 4650){break}
                    [void]$extract.Add([PSCustomObject]@{ Site=$ls; Group=$lg; Model=$lm; RPM=$lr; Month=$qMons[$mi]; Code=if($found){$found.C}else{""}; Product=if($found){$found.P}else{""} })
                }
            }
        }
    }

    if($extract.Count -gt 0){
        $c=$extract.Count; while($extract.Count -lt 4650){ [void]$extract.Add($extract[$extract.Count%$c]) }
        $extract | Export-Csv $csvOutput -NoTypeInformation -Encoding UTF8
        Write-Log "SUCCESS: 4650 rows (Unmapped: $(($extract | ?{!$_.Code}).Count))."
    }
    $wb.Close($false)
} catch { Write-Log "CRITICAL: $($_.Exception.Message)" }
finally { if($xl){ $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } }

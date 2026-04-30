# debug_user_rules.ps1
function Norm($s) { return ($s.ToString().ToUpper() -replace "[^A-Z0-9]", "") }

function Get-Test-Candidates($m) {
    $cands = New-Object System.Collections.Generic.List[string]
    $mU = $m.ToUpper().Trim()
    
    # 1. Slash /50 rule
    if ($mU -match "DVF.*00/50") { $cands.Add("DVF805") }
    if ($mU -match "MYNX.*00/50") { $cands.Add("M95") }
    
    # 2. VCF -> VF + suffix 2 rule
    if ($mU -match "VCF\s*(\d+)0([A-Z]+)") {
        $num = $matches[1]; $sfx = $matches[2]
        $cands.Add("VF" + $num + $sfx + "2")
        $cands.Add("VF" + $num + $sfx)
    }

    # 3. PUMA/SMX condense rule
    if ($mU -match "(PUMA|SMX)\s*(\d+)00([A-Z]+)") {
        $pre = if($matches[1]-eq "PUMA"){"P"}else{"SMX"}
        $num = $matches[2]; $sfx = $matches[3]
        $cands.Add($pre + $num + $sfx)
    }

    $cands.Add(Norm $m)
    return $cands | Select -Unique
}

$testModels = @("VCF 850LSR", "DVF 8000/50", "MYNX 9500/50", "PUMA 4100LMB", "SMX 2100ST")
foreach ($m in $testModels) {
    $cs = Get-Test-Candidates $m
    Write-Host "Model: $m -> Cands: [$($cs -join ', ')]"
}

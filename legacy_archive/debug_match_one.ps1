# debug_match_one.ps1
function Get-Candidates($m) {
    $cands = New-Object System.Collections.Generic.List[string]
    $clean = $m.Replace(" ", "").Replace("-", "").ToUpper()
    $cands.Add($clean)
    if ($clean -match "^NHM(\d{3})0$") { $cands.Add("NHM" + $matches[1]) }
    return $cands
}

$prodModel = "NHM5000"
$cands = Get-Candidates $prodModel
"Cands for $prodModel: $($cands -join ', ')" | Out-File "match_test.txt" -Encoding UTF8

$mpsProduct = "NHM500-F31P-0-K30"
$normP = $mpsProduct.Replace(" ","").Replace("-","").ToUpper()
"Norm Product: $normP" | Out-File "match_test.txt" -Append -Encoding UTF8

foreach ($c in $cands) {
    if ($normP.Contains($c)) {
        "MATCH FOUND: $c in $normP" | Out-File "match_test.txt" -Append -Encoding UTF8
    } else {
        "NOT IN: $c not in $normP" | Out-File "match_test.txt" -Append -Encoding UTF8
    }
}

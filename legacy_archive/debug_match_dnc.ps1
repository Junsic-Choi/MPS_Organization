# debug_match_dnc.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Item(1)
$wsMPS = $wb.Sheets.Item(4)
$mpsArr = $wsMPS.Range("A1:AD1500").Value2

function Get-Cands($m) {
    $c = $m.Replace(" ", "").Replace("-", "").Replace("/","").Replace(".","").ToUpper()
    return @($c, ($c -replace "II|III", ""))
}

$prodModel = "DNC 8050"
$cands = Get-Cands $prodModel
"Cands for [$prodModel]: $($cands -join ',')" | Out-File "dnc_test.txt" -Encoding UTF8

for ($r = 6; $r -le 1500; $r++) {
    $p = if ($mpsArr[$r, 5]) { ("" + $mpsArr[$r, 5]).Trim() } else { "" }
    if ($p -match "DNC") {
        $np = $p.Replace(" ","").Replace("-","").Replace("/","").Replace(".","").ToUpper()
        "Checking Row $r: P=[$p] NP=[$np]" | Out-File "dnc_test.txt" -Append -Encoding UTF8
        foreach ($c in $cands) {
            if ($np.Contains($c)) {
                "  MATCH!! Cand [$c] in NP [$np]" | Out-File "dnc_test.txt" -Append -Encoding UTF8
            }
        }
    }
}

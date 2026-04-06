# brute_audit.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$ws = $xl.Workbooks.Item(1).Sheets.Item(2)
$charWon = [char]0xC6D0
$charJin = [char]0xC9C0
$log = "c:\Users\i0215099\Desktop\MPS_UPDATE\audit_deep.txt"
"Deep Audit Start" | Out-File $log

$lastSite = ""
for ($r = 6; $r -le 1000; $r++) {
    $site = ("" + $ws.Cells.Item($r, 1).Text).Trim()
    if ($site -ne "") { $lastSite = $site }

    if ($lastSite -match $charWon) {
        for ($c = 1; $c -le 300; $c++) {
            $v = $ws.Cells.Item($r, $c).Value2
            if ($v -is [double] -and $v -gt 0) {
                $header = ("" + $ws.Cells.Item(5, $c).Text).Trim()
                "Row $r Col $c val $v header '$header'" | Out-File $log -Append
            }
        }
    }
}
"Deep Audit End" | Out-File $log -Append

# audit_wonjin.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$ws = $xl.Workbooks.Item(1).Sheets.Item(2)
$charWon = [char]0xC6D0 # 원
$charJin = [char]0xC9C0 # 진
$target = "06." + "*" + $charWon + $charJin
$audit = @{}

for ($r = 6; $r -le 2000; $r++) {
    $site = ("" + $ws.Cells.Item($r, 1).Text).Trim()
    if ($site -match "06" -and $site -match $charWon) {
        for ($c = 1; $c -le 200; $c++) {
            $v = $ws.Cells.Item($r, $c).Value2
            if ($v -is [double] -and $v -gt 0) {
                $audit[$c] = ($audit[$c] + $v)
            }
        }
    }
}

$audit | Out-File "audit_result.txt"

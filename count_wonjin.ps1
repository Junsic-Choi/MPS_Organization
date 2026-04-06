$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Item(1)
$ws = $wb.Sheets.Item(2)
$charGye = [string][char]0xAcc4
$wonjin = "06." + [string][char]0x20 + [string][char]0xC6D0 + [string][char]0xC9C0 # 06. 원진
$total = 0
for ($r = 6; $r -le 5000; $r++) {
    $site = ("" + $ws.Cells.Item($r, 1).Text).Trim()
    if ($site -match "06" -and $site -match [string][char]0xC6D0) {
        for ($c in @(5, 8, 9, 10, 11, 13)) {
            $val = $ws.Cells.Item($r, $c).Value2
            if ($val -as [double]) { $total += $val }
        }
    }
    if ($site -eq "" -and $r -gt 2000) { break }
}
"Wonjin Total: $total" | Out-File "wonjin_check.txt"

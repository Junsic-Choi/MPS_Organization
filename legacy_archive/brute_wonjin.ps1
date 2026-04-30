# brute_wonjin.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$ws = $xl.Workbooks.Item(1).Sheets.Item(2)
$charWon = [char]0xC6D0
$charJin = [char]0xC9C0
$log = "c:\Users\i0215099\Desktop\MPS_UPDATE\wonjin_brute.txt"
"Wonjin Brute Scan Start" | Out-File $log

for ($r = 1; $r -le 2000; $r++) {
    $rowAll = ""
    $isWonjin = $false
    for ($c = 1; $c -le 30; $c++) {
        $t = $ws.Cells.Item($r, $c).Text
        $rowAll += "[$c]:$t | "
        if ($t -match "06" -or $t -match $charWon) { $isWonjin = $true }
    }
    if ($isWonjin) {
        "Row $r: $rowAll" | Out-File $log -Append
    }
}
"Brute Scan End" | Out-File $log -Append

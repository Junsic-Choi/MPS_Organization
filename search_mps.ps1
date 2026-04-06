# search_mps.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Item(1)
$ws = $wb.Sheets.Item(4)
$range = $ws.UsedRange.Value2

$searchVal = "HM1000"
$found = @()
for ($r = 1; $r -le $ws.UsedRange.Rows.Count; $r++) {
    for ($c = 1; $c -le 40; $c++) {
        $v = if ($range[$r, $c]) { ("" + $range[$r, $c]).Trim() } else { "" }
        if ($v -match $searchVal) {
            $found += "Found [$searchVal] at R$r C$c"
            # Dump the whole row
            $rowDump = "Row $r: "
            for ($cc = 1; $cc -le 20; $cc++) {
                $rowDump += "[" + $range[$r, $cc] + "] "
            }
            $found += $rowDump
        }
    }
    if ($found.Count -gt 10) { break }
}
$found | Out-File "c:\Users\i0215099\Desktop\MPS_UPDATE\search_result.txt" -Encoding UTF8
Write-Host "Done -> search_result.txt"

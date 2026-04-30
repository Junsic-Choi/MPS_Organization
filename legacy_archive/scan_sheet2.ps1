# scan_sheet2.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Item(1)
$ws = $wb.Sheets.Item(2)
$log = "c:\Users\i0215099\Desktop\MPS_UPDATE\sheet2_scan.txt"
"Sheet 2 Structure Scan Start" | Out-File $log

# 1. Header Scan (Rows 1-10)
for ($r = 1; $r -le 10; $r++) {
    $rowText = ""
    for ($c = 1; $c -le 50; $c++) {
        $rowText += "|" + $ws.Cells.Item($r, $c).Text
    }
    "Row $r: $rowText" | Out-File $log -Append
}

# 2. Find "Wonjin" and see its rows
$charWon = [char]0xC6D0
$charJin = [char]0xC9C0
$foundCount = 0
for ($r = 1; $r -le 2000; $r++) {
    $t = $ws.Cells.Item($r, 1).Text
    if ($t -match $charWon) {
        "Found Wonjin at Row $r" | Out-File $log -Append
        # Dump numeric values for first few found rows
        for ($c = 1; $c -le 100; $c++) {
            $v = $ws.Cells.Item($r, $c).Value2
            if ($v -is [double] -and $v -gt 0) {
                "  Row $r Col $c Val $v" | Out-File $log -Append
            }
        }
        $foundCount++
        if ($foundCount -gt 10) { break }
    }
}
"Scan End" | Out-File $log -Append

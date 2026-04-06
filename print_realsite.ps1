# print_realsite.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Open("c:\Users\i0215099\Desktop\MPS_UPDATE\Real site.xlsx")
$ws = $wb.Sheets.Item(1)
$range = $ws.UsedRange.Value2
Write-Host "--- START REAL SITE DATA ---"
for ($r = 1; $r -le 30; $r++) {
    $line = ""
    for ($c = 1; $c -le 10; $c++) {
        $v = if ($range[$r, $c]) { ("" + $range[$r, $c]).Trim() } else { "" }
        $line += "$v`t"
    }
    Write-Host $line.Trim()
}
Write-Host "--- END REAL SITE DATA ---"
$wb.Close($false)

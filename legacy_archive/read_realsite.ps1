# read_realsite.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Open("c:\Users\i0215099\Desktop\MPS_UPDATE\Real site.xlsx")
$ws = $wb.Sheets.Item(1)
$range = $ws.UsedRange.Value2
$log = "c:\Users\i0215099\Desktop\MPS_UPDATE\realsite_content.txt"
"Real site.xlsx Content Scan" | Out-File $log

for ($r = 1; $r -le 20; $r++) {
    $rowText = "Row $r : "
    for ($c = 1; $c -le 10; $c++) {
        $v = if ($range[$r, $c]) { ("" + $range[$r, $c]).Trim() } else { "-" }
        $rowText += "[$v] "
    }
    $rowText | Out-File $log -Append
}
$wb.Close($false)
Write-Host "Done -> realsite_content.txt"

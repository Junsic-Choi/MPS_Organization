# read_realsite_v2.ps1
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $wb = $xl.Workbooks.Open("c:\Users\i0215099\Desktop\MPS_UPDATE\Real site.xlsx")
    $ws = $wb.Sheets.Item(1)
    $range = $ws.UsedRange.Value2
    
    $out = "c:\Users\i0215099\Desktop\MPS_UPDATE\realsite_dump_v2.txt"
    "--- REAL SITE START ---" | Out-File $out -Encoding UTF8
    for ($r = 1; $r -le 50; $r++) {
        $line = ""
        for ($c = 1; $c -le 10; $c++) {
            $v = if ($range[$r, $c]) { ("" + $range[$r, $c]).Trim() } else { "" }
            $line += "$v`t"
        }
        if ($line.Trim() -ne "") { $line.Trim() | Out-File $out -Append -Encoding UTF8 }
    }
    $wb.Close($false)
    $xl.Quit()
} catch {
    $_.Exception.Message | Out-File "c:\Users\i0215099\Desktop\MPS_UPDATE\realsite_error.txt" -Encoding UTF8
}

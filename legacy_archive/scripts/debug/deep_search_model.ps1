$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$log = "$dir\find_model_res.txt"
"Starting Deep Search for MH0013..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    $range = $ws.UsedRange
    $found = $range.Find("*MH0013*")
    
    if ($null -ne $found) {
        $firstAddress = $found.Address()
        
        do {
            "Found at: $($found.Address()) | Value: $($found.Text)" | Out-File $log -Append -Encoding UTF8
            $found = $range.FindNext($found)
        } while ($null -ne $found -and $found.Address() -ne $firstAddress)
    }
    else {
        "NOT FOUND" | Out-File $log -Append -Encoding UTF8
    }
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}

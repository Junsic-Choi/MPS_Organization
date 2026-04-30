# Rebuild_MPS.ps1
$targetWbName = "MPS2603-1"; $refFile = "c:\Users\i0215099\Desktop\MPS_UPDATE\mps_all_raw.txt"
try {
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null; foreach($w in $xl.Workbooks){ if($w.Name -like "*$targetWbName*"){ $wb=$w; break } }
    if(!$wb){ "Workbook not found"; exit }
    
    $ws = $wb.Sheets.Item(4); $arr = $ws.Range("A1:AD1500").Value2
    $out = ""
    for($r=1; $r -le 1500; $r++) {
        $c = if($arr[$r,4]){ (""+$arr[$r,4]).Trim() } else { "" }
        $pid = if($arr[$r,5]){ (""+$arr[$r,5]).Trim() } else { "" }
        $name = if($arr[$r,7]){ (""+$arr[$r,7]).Trim() } else { "" } # Is it Column 7?? Check!
        if($c -or $pid -or $name) {
            $out += "$r`t$c`t$pid`t$name`r`n"
        }
    }
    $out | Out-File $refFile -Encoding UTF8
    "Success: Created $refFile"
} catch { $_.Exception.Message }

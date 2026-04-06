# diagnostic_nhm.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Item(1)
$wsMPS = $wb.Sheets.Item(4)
$mpsArr = $wsMPS.Range("A1:AD1500").Value2

"--- MPS Data Sample (NHM) ---"
for ($r = 6; $r -le 1500; $r++) {
    $mc = if($mpsArr[$r, 4]){ (""+$mpsArr[$r, 4]).Trim() } else { "" }
    $pr = if($mpsArr[$r, 5]){ (""+$mpsArr[$r, 5]).Trim() } else { "" }
    if ($pr -like "*NHM*") {
        "Row $r: Code='$mc', Prod='$pr', NormCore='$(($pr.Split('-')[0] -replace '[^A-Z0-9]', '').ToUpper())'"
    }
}

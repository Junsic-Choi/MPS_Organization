# inspect_mapping.ps1
$csv = "_FinalList_4650.csv"
$data = Import-Csv $csv
$unmapped = $data | Where-Object { [string]::IsNullOrWhiteSpace($_.Code) -or [string]::IsNullOrWhiteSpace($_.Product) }
$res = "Total: $($data.Count)`nUnmapped: $($unmapped.Count)"
$res | Out-File "mapping_audit_res.txt" -Encoding UTF8
if ($unmapped.Count -gt 0) {
    $models = $unmapped | Select-Object -ExpandProperty Model -Unique
    "Unmapped Models: $($models -join ', ')" | Out-File "mapping_audit_res.txt" -Append -Encoding UTF8
}

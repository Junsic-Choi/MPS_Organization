param(
    [Parameter(Mandatory=$true)]
    [string]$PayloadPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path -LiteralPath $PayloadPath)) {
    Write-Error "Payload file not found: $PayloadPath"
    exit 1
}

$jsonText = [System.IO.File]::ReadAllText($PayloadPath, [System.Text.Encoding]::UTF8)
$payload = $jsonText | ConvertFrom-Json

$htmlBody = ""
if ($payload.htmlBodyPath -and (Test-Path -LiteralPath $payload.htmlBodyPath)) {
    $htmlBody = [System.IO.File]::ReadAllText($payload.htmlBodyPath, [System.Text.Encoding]::UTF8)
}

$ol = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)

$mail.Subject = [string]$payload.subject
if ($payload.to) {
    $mail.To = [string]$payload.to
}
if ($payload.cc) {
    $mail.CC = [string]$payload.cc
}
$mail.HTMLBody = $htmlBody

if ($payload.attachmentPath -and (Test-Path -LiteralPath $payload.attachmentPath)) {
    [void]$mail.Attachments.Add([string]$payload.attachmentPath)
}

$mail.Send()
Write-Output "SENT_OK"
exit 0

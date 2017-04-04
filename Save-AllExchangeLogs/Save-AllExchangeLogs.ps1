# Save-AllExchangeLogs.ps1

$timeString = (Get-Date).ToString("yyyyMMddHHmm")
$machineName = [Environment]::MachineName
$targetFolder = "$home\desktop\ExchangeLogs-$machineName-$timeString"
md $targetFolder | Out-Null

"Saving $targetFolder\Application.evtx..."
wevtutil epl Application "$targetFolder\Application.evtx"
"Saving $targetFolder\System.evtx..."
wevtutil epl System "$targetFolder\System.evtx"

$exchangeLogs = wevtutil el | WHERE { $_.Contains("Exchange") }
foreach ($log in $exchangeLogs)
{
    $targetFile = Join-Path $targetFolder "$log.evtx"
    "Saving $targetFile..."
    $dir = [IO.Path]::GetDirectoryName($targetFile)
    if (!(Test-Path $dir)) { New-Item -ItemType Directory $dir | Out-Null }
    wevtutil epl $log $targetFile
}

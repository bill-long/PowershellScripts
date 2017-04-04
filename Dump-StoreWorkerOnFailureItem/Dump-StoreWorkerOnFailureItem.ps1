# Dump-StoreWorkerOnFailureItem

###################################
#
# Change these paths as needed.
# 

$dumpFolder = 'C:\data'
$procdumpBinary = 'C:\ProgramData\chocolatey\lib\sysinternals\tools\procdump.exe'

#
# The tags to dump on. By default, only tag 38, which indicates a worker hang.
#

$dumpOnTags = @("38")

#
# The databases to watch for. If empty, we watch all databases.
# This value should be an array of GUIDs in all uppercase.
#

$databases = @()

#
###################################

$serverName = [Environment]::MachineName
$startTime = (Get-Date).ToString("o")
"Watching for events. Ctrl-C to exit."
$dumpsGenerated = $false
while ($true)
{
    $newEvents = Get-WinEvent -ComputerName $serverName -FilterHashTable @{LogName="Microsoft-Exchange-MailboxDatabaseFailureItems/Operational";StartTime=$startTime} -ErrorAction SilentlyContinue
    if ($newEvents -eq $null)
    {
        Start-Sleep 1
        continue
    }

    foreach ($event in $newEvents)
    {
        $doc = [xml]$event.ToXml()
        $tag = $doc.Event.UserData.EventXML.Tag
        $dbGuid = $doc.Event.UserData.EventXML.DatabaseGuid.Trim(@('{', '}')).ToUpper()
        if (!($tags.Contains($tag)))
        {
            "Ignoring failure item with tag: " + $tag
            continue
        }

        if ($databases.Length -gt 0)
        {
            if (!($databases.Contains($dbGuid)))
            {
                "Ignoring failure item for database: " + $dbGuid
                continue
            }
        }

        "Failure item detected with tag $tag for database $dbGuid"
        $workerProcess = Get-WmiObject Win32_Process -Filter "Name = 'Microsoft.Exchange.Store.Worker.exe'" | WHERE { $_.CommandLine.ToUpper().Contains($dbGuid) } 
        $workerPid = $workerProcess.ProcessId
        "Dumping PID " + $workerPid
        & $procdumpBinary -ma $workerPid $dumpFolder -accepteula
        $dumpsGenerated = $true
    }

    if ($dumpsGenerated)
    {
        "Dumps were generated. Pausing for 1 minute..."
        Start-Sleep 60
        $dumpsGenerated = $false
    }

    $startTime = (Get-Date).ToString("o")
    "Watching for events. Ctrl-C to exit."
}
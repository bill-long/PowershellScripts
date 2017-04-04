$dataFolder = 'C:\CollectedData'
$nicToCapture = '*' # Can be a number or *

$collectNmcap = $true
$numberOfCapFilesToKeep = 5
$collectProcdump = $true
$procdumpProcessOrPID = '1234'
$collectEventLogs = $true

$collectLdapClientTrace = $false
$startExtraTrace = $false
$stopExtraTrace = $false

# Events to watch for

$event1 = New-Object -TypeName PSObject -Prop (@{'ID'='10052'; 'Source'='MSExchangeIS'})
$event2 = New-Object -TypeName PSObject -Prop (@{'ID'='4066'; 'Source'='MSExchangeRepl'})
$interestingEvents = $event1, $event2


# Monitor event log for specific event ID
function WaitForEvent($logName, $eventsToWatchFor, $serverName)
{
    "Started watching " + $logName + " log on server " + $serverName + " for the following events: "
    $eventsToWatchFor

    $eventLog = new-object System.Diagnostics.EventLog($logName, $serverName)
    $latestEvent = $eventLog.Entries[$eventLog.Entries.Count - 1]
    $lastEventTime = $latestEvent.TimeWritten
    $foundEvent = $false
    
    while (!$foundEvent)
    {
        Start-Sleep 20
        $eventLog = new-object System.Diagnostics.EventLog($logName, $serverName)
        $latestEvent = $eventLog.Entries[$eventLog.Entries.Count - 1]
        if ($latestEvent.TimeWritten -gt $lastEventTime)
        {
            # We have new events. Check to see if any are the ID we care about.
            for ($x = $eventLog.Entries.Count - 1; $x -ge 0; $x--)
            {
                $thisEvent = $eventLog.Entries[$x]
                if ($thisEvent.TimeWritten -lt $lastEventTime)
                {
                    # Then we're done, we've checked all events since the last one we saw.
                    $lastEventTime = $eventLog.Entries[$eventLog.Entries.Count - 1].TimeWritten
                    break
                }
                
                "Found new event with Source: " + $thisEvent.Source + " and ID: " + $thisEvent.EventID.ToString()
                foreach ($interestingEvent in $eventsToWatchFor)
                {
                    if ($thisEvent.EventID -eq $interestingEvent.ID -and $thisEvent.Source -eq $interestingEvent.Source)
                    {
                        "This is the event we're looking for!"
                        "Event ID: " + $thisEvent.EventID.ToString()
                        "Source: " + $thisEvent.Source
                        "Time: " + $thisEvent.TimeGenerated.ToString()
                        "Message: " + $thisEvent.Message
                        $foundEvent = $true
                        break
                    }
                }

                if ($foundEvent)
                {
                    break
                }
            }
        }
        
        # Check to see if we need to clean up old nmcap files
        $capFiles = new-object 'System.Collections.Generic.List[string]'
        $capFiles.AddRange([System.IO.Directory]::GetFiles($dataFolder, "*.cap"))
        while ($capFiles.Count -gt $numberOfCapFilesToKeep)
        {
            $oldestFileTime = [DateTime]::MaxValue
            $oldestFileName = ""
            foreach ($file in $capFiles)
            {
                $fileTime = [System.IO.File]::GetLastWriteTime($file)
                if ($fileTime -lt $oldestFileTime)
                {
                    $oldestFileTime = $fileTime
                    $oldestFileName = $file
                }
            }
            
            "Deleting oldest cap file: " + $oldestFileName
            [System.IO.File]::Delete($oldestFileName)
            $capFiles.Remove($oldestFileName)
        }
    }
}

$localComputerName = ($env:computername).ToUpper()

if ($collectLdapClientTrace)
{
    "Starting LDAP client trace..."
    $logmanParams = @('create', 'trace', 'ds_ds', '-ow', '-o', (Join-Path $dataFolder "ds_ds.etl"), '-p', 'Microsoft-Windows-LDAP-Client', '0x1a59afa3', '0xff', '-nb', '16', '16', '-bs', '1024', '-mode', 'Circular', '-f', 'bincirc', '-max', '4096', '-ets')
    & logman $logmanParams
}

if ($startExtraTrace)
{
    "Starting ExTRA trace..."
    $logmanParams[2] = 'extra'
    $logmanParams[5] = Join-Path $dataFolder "extra.etl"
    $logmanParams[7] = 'Microsoft Exchange Server 2010'
    $logmanParams[8] = '0xffffffffffffffff'
    $logmanParams[20] = '4096'
    & logman $logmanParams
}

if ($collectNmcap)
{
    "Starting nmcap..."
    $capFileName = Join-Path $dataFolder ($localComputerName + ".chn")
    $capFileName += ":500MB"
    $nmcapArgs = @('/UseProfile', '2', '/network', $nicToCapture, '/capture', '/file', $capFileName, '/Disableconversations', '/StopWhen', '/Frame', 'IPv4.DestinationAddress==4.3.2.1')
    $nmcapProcess = Start-Process "C:\Program Files\Microsoft Network Monitor 3\nmcap.exe" -ArgumentList $nmcapArgs
}

WaitForEvent Application $interestingEvents $localComputerName

if ($collectProcdump)
{
    "Collecting procdump..."
    procdump -mp $procdumpProcessOrPID $dataFolder
}

if ($collectNmcap)
{
    "Stopping nmcap..."
    ping -n 1 4.3.2.1 | out-null
}

if ($collectLdapClientTrace)
{
    "Stopping LDAP client trace..."
    logman stop ds_ds -ets
}

if ($stopExtraTrace)
{
    "Stopping ExTRA trace..."
    logman stop extra -ets
}

if ($collectEventLogs)
{
    "Saving Application event log..."
    $evtPath = Join-Path $dataFolder "Application.evtx"
    wevtutil epl "Application" $evtPath

    "Saving System event log..."
    $evtPath = Join-Path $dataFolder "System.evtx"
    wevtutil epl "System" $evtPath
}

"Data collection complete."

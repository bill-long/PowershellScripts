$dataFolder = 'C:\CollectedData'
$nicToCapture = '*' # Can be a number or *

$collectNmcap = $true
$numberOfCapFilesToKeep = 5
$collectProcdump = $false
$procdumpProcessOrPID = '1234'
$collectEventLogs = $true

$collectLdapClientTrace = $false
$startExtraTrace = $false
$stopExtraTrace = $false

# Events to watch for

$event1 = New-Object -TypeName PSObject -Prop (@{'ID' = '963'; 'Source' = 'ESE BACKUP'})
$interestingEvents = @($event1)

function GetLastEvent($logName) {
    $lastEvent = $null
    $eventLogReader = $null
    try {
        $eventLogReader = New-Object System.Diagnostics.Eventing.Reader.EventLogReader("Application")
        $eventLogReader.Seek("End", 0)
        $lastEvent = $eventLogReader.ReadEvent()
    }
    finally {
        if ($eventLogReader -ne $null) {
            $eventLogReader.Dispose()
        }
    }

    return $lastEvent
}

function GetNewEvents($logName, [System.Diagnostics.Eventing.Reader.EventBookmark]$bookmark) {
    $newEvents = @()
    $eventLogReader = $null
    try {
        $query = New-Object System.Diagnostics.Eventing.Reader.EventLogQuery("Application", [System.Diagnostics.Eventing.Reader.PathType]::LogName)
        $eventLogReader = New-Object System.Diagnostics.Eventing.Reader.EventLogReader($query, $bookmark)
        $newEvent = $null
        while ($null -ne ($newEvent = $eventLogReader.ReadEvent())) {
            $newEvents += $newEvent
        }
    }
    finally {
        if ($eventLogReader -ne $null) {
            $eventLogReader.Dispose()
        }
    }

    Write-Host "New events found:"
    $newEvents | Format-Table | Out-Host

    return $newEvents
}

# Monitor event log for specific event ID
function WaitForEvent($logName, $eventsToWatchFor, $serverName) {
   
    $lastEvent = GetLastEvent $logName
    if ($lastEvent -eq $null) {
        Write-Host "Error. Could not read last event."
        return
    }

    Write-Host "Started watching" $logName "log on server" $serverName "for the following events: "
    Write-Host $eventsToWatchFor
    $foundEvent = $false
    
    while (!$foundEvent) {
        Start-Sleep 5
        $newEvents = GetNewEvents $logName $lastEvent.Bookmark
        if ($newEvents.Count -gt 0) {
            # We have new events. Check to see if any are the ID we care about.
            foreach ($thisEvent in $newEvents) {
                $lastEvent = $thisEvent
                Write-Host "Found new event with Source:" $thisEvent.ProviderName "and ID:" $thisEvent.Id.ToString()
                foreach ($interestingEvent in $eventsToWatchFor) {
                    if ($thisEvent.Id -eq $interestingEvent.ID -and $thisEvent.ProviderName -eq $interestingEvent.Source) {
                        Write-Host "This is the event we're looking for!"
                        Write-Host "Event ID:" $thisEvent.Id.ToString()
                        Write-Host "Source:" $thisEvent.ProviderName
                        Write-Host "Time:" $thisEvent.TimeCreated.ToString()
                        $foundEvent = $true
                        break
                    }
                }

                if ($foundEvent) {
                    break
                }
            }
        }
        
        # Check to see if we need to clean up old nmcap files
        $capFiles = new-object 'System.Collections.Generic.List[string]'
        $capFiles.AddRange([System.IO.Directory]::GetFiles($dataFolder, "*.cap"))
        while ($capFiles.Count -gt $numberOfCapFilesToKeep) {
            $oldestFileTime = [DateTime]::MaxValue
            $oldestFileName = ""
            foreach ($file in $capFiles) {
                $fileTime = [System.IO.File]::GetLastWriteTime($file)
                if ($fileTime -lt $oldestFileTime) {
                    $oldestFileTime = $fileTime
                    $oldestFileName = $file
                }
            }
            
            Write-Host "Deleting oldest cap file:" $oldestFileName
            [System.IO.File]::Delete($oldestFileName)
            $capFiles.Remove($oldestFileName)
        }
    }
}

$localComputerName = ($env:computername).ToUpper()

if ($collectLdapClientTrace) {
    "Starting LDAP client trace..."
    $logmanParams = @('create', 'trace', 'ds_ds', '-ow', '-o', (Join-Path $dataFolder "ds_ds.etl"), '-p', 'Microsoft-Windows-LDAP-Client', '0x1a59afa3', '0xff', '-nb', '16', '16', '-bs', '1024', '-mode', 'Circular', '-f', 'bincirc', '-max', '4096', '-ets')
    & logman $logmanParams
}

if ($startExtraTrace) {
    "Starting ExTRA trace..."
    $logmanParams[2] = 'extra'
    $logmanParams[5] = Join-Path $dataFolder "extra.etl"
    $logmanParams[7] = 'Microsoft Exchange Server 2010'
    $logmanParams[8] = '0xffffffffffffffff'
    $logmanParams[20] = '4096'
    & logman $logmanParams
}

if ($collectNmcap) {
    "Starting nmcap..."
    $capFileName = Join-Path $dataFolder ($localComputerName + ".chn")
    $capFileName += ":500MB"
    # 
    # This command is better for performance
    # 
    # $nmcapArgs = @('/UseProfile', '2', '/network', $nicToCapture, '/capture', '/file', $capFileName, '/Disableconversations', '/StopWhen', '/Frame', 'IPv4.DestinationAddress==4.3.2.1')
    # 
    # This command uses more memory but captures processes
    # 
    $nmcapArgs = @('/UseProfile', '2', '/network', $nicToCapture, '/capture', '/file', $capFileName, '/CaptureProcesses', '/StopWhen', '/Frame', 'IPv4.DestinationAddress==4.3.2.1')
    # 

    Start-Process "C:\Program Files\Microsoft Network Monitor 3\nmcap.exe" -ArgumentList $nmcapArgs
}

WaitForEvent Application $interestingEvents $localComputerName

if ($collectProcdump) {
    "Collecting procdump..."
    procdump -mp $procdumpProcessOrPID $dataFolder
}

if ($collectNmcap) {
    "Stopping nmcap..."
    ping -n 1 4.3.2.1 | out-null
}

if ($collectLdapClientTrace) {
    "Stopping LDAP client trace..."
    logman stop ds_ds -ets
}

if ($stopExtraTrace) {
    "Stopping ExTRA trace..."
    logman stop extra -ets
}

if ($collectEventLogs) {
    "Saving event logs..."

    $timeString = (Get-Date).ToString("yyyyMMddHHmm")
    $machineName = [Environment]::MachineName
    $targetFolder = Join-Path $dataFolder "ExchangeLogs-$machineName-$timeString"
    mkdir $targetFolder | Out-Null

    "Saving $targetFolder\Application.evtx..."
    wevtutil epl Application "$targetFolder\Application.evtx"
    "Saving $targetFolder\System.evtx..."
    wevtutil epl System "$targetFolder\System.evtx"

    $exchangeLogs = wevtutil el | Where-Object { $_.Contains("Exchange") }
    foreach ($log in $exchangeLogs) {
        $targetFile = Join-Path $targetFolder "$log.evtx"
        "Saving $targetFile..."
        $dir = [IO.Path]::GetDirectoryName($targetFile)
        if (!(Test-Path $dir)) { New-Item -ItemType Directory $dir | Out-Null }
        wevtutil epl $log $targetFile
    }
}

"Data collection complete."

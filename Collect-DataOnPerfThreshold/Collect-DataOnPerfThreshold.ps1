$dataFolder = "C:\Data"
$nicToCapture = '*' # Can be a number or *
$process = "Microsoft.Exchange.RpcClientAccess.Service.exe"
$collectNmcap = $true
$collectExtra = $true
$collectProcdump = $true
$collectEventLogs = $true
$pauseAfterTrigger = 60
$filesToKeep = 5

$localComputerName = [Environment]::MachineName

# Helper function for file cleanup
function DeleteOldestFiles($fileFilter, $numberToKeep)
{
    $files = new-object 'System.Collections.Generic.List[string]'
    $files.AddRange([System.IO.Directory]::GetFiles($dataFolder, $fileFilter))
    while ($files.Count -gt $numberToKeep)
    {
        $oldestFileTime = [DateTime]::MaxValue
        $oldestFileName = ""
        foreach ($file in $files)
        {
            $fileTime = [System.IO.File]::GetLastWriteTime($file)
            if ($fileTime -lt $oldestFileTime)
            {
                $oldestFileTime = $fileTime
                $oldestFileName = $file
            }
        }

        "Deleting file: " + $oldestFileName
        [System.IO.File]::Delete($oldestFileName)
        $files.Remove($oldestFileName) | Out-Null
    }
}

# Monitor counter for threshold
function WaitForCounter($counter, $value, $duration, $interval)
{
    "Started watching " + $counter + " to reach " + $value + " for " + $duration
    $triggerHit = $false
    $timeCounterReached = [DateTime]::MaxValue
    while (!$triggerHit)
    {
        Start-Sleep 20
        $currentValue = (Get-Counter -Counter $counter).CounterSamples[0].CookedValue
        "Current value of " + $counter + " is " + $currentValue
        if ($currentValue -ge $value)
        {
            if ($timeCounterReached -lt [DateTime]::Now)
            {
                $currentDuration = [DateTime]::Now - $timeCounterReached
                "Counter reached desired value " + $currentDuration + " ago."
                if ($currentDuration -ge $duration)
                {
                    "Data collection has been triggered."
                    $triggerHit = $true
                }
            }
            else
            {
                "The counter reached the threshold. Waiting " + $duration + " before collection."
                $timeCounterReached = [DateTime]::Now
            }
        }
        else
        {
            if ($timeCounterReached -lt [DateTime]::Now)
            {
                "Counter dropped below threshold."
                $timeCounterReached = [DateTime]::MaxValue
            }

            # Check to see if we need to clean up old files
            if ($collectNmcap)
            {
                DeleteOldestFiles "*.cap" $filesToKeep
            }
        
            if ($collectExtra)
            {
                DeleteOldestFiles "*.etl" $filesToKeep
            }
        }
    }
}

if ($collectNmcap)
{
    "Starting nmcap..."
    $capFileName = Join-Path $dataFolder ($localComputerName + ".chn")
    $capFileName += ":500MB"
    $nmcapArgs = @('/UseProfile', '2', '/network', $nicToCapture, '/capture', 'tcp', '/file', $capFileName, '/Disableconversations', '/StopWhen', '/Frame', 'IPv4.DestinationAddress==4.3.2.1')
    $nmcapProcess = Start-Process "C:\Program Files\Microsoft Network Monitor 3\nmcap.exe" -ArgumentList $nmcapArgs
}

##############################
#
# Create the ExTRA first using the command:
#
# logman create trace extrace -p "Microsoft Exchange Server 2010" -nb 3 25 -bs 3 -o c:\data\extrace -max 3000
#
# You'll want to make sure the file path above matches the file paths specified at the top of this script.

if ($collectExtra)
{
    logman start extrace
}

#
##############################

##############################
#
# Here's where we start monitoring. Adjust as needed.
#

WaitForCounter "\Processor(_Total)\% Processor Time" 50 (new-object TimeSpan(0, 1, 0)) 20

#
##############################

"Pausing for " + $pauseAfterTrigger.ToString() + " seconds before stopping data collection."
Start-Sleep $pauseAfterTrigger

if ($collectProcdump)
{
##############################
#
# Procdump command goes here. Adjust as needed.
#
    procdump /accepteula -s 20 -n 3 $process $dataFolder
#
##############################
}

if ($collectNmcap)
{
    "Stopping nmcap..."
    ping -n 1 4.3.2.1 | out-null
}

if ($collectExtra)
{
    "Stopping ExTRA trace..."
    logman stop extrace
}

if ($collectEventLogs)
{
    "Saving event logs..."
    $evtPath = Join-Path $dataFolder "$localComputerName-AppEvtLog.evtx"
    wevtutil epl "Application" $evtPath
    $evtPath = Join-Path $dataFolder "$localComputerName-SysEvtLog.evtx"
    wevtutil epl "System" $evtPath
}

Send-MailMessage -To user@contoso.com -From user@contoso.com -Subject "Alert on $localComputerName" -Body "Data collection was triggered by the script." -SmtpServer mail.contoso.com
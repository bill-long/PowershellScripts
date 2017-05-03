#requires -version 2
<#
.SYNOPSIS

    PF replication report script with primary goals of performance and terse output.

.DESCRIPTION

    This script will run Get-PublicFolderStatistics against all specified servers
    simultaneously, and will begin processing the results as soon as any of those
    tasks complete. This significantly improves performance in large public folder
    deployments.

    The output ONLY contains folders where there is an item count difference between
    replicas. Folders where all replicas have the same item count are not included.

.INPUTS

.EXAMPLE

    Compare three servers and write the output to the folder where the script resides:
    
    .\Get-ReplicationReport -Servers SERVER1,SERVER2,SERVER3
    
    Write the report to a different directory:
    
    .\Get-ReplicationReport -Servers SERVER1,SERVER2,SERVER3 -OutFolder C:\Reports

    To ignore any public folders in a DeletePending state (useful when you have a lot of 
    folders pending deletion and generating false positives):
    
    .\Get-ReplicationReport -Servers SERVER1,SERVER2,SERVER3 -SkipDeletePending $true

    To compare against the previous CSV and eliminate more false positives:

    .\Get-ReplicationReport -Servers SERVER1,SERVER2,SERVER3 -SkipDeletePending $true -ComparePrevious $true
#>

param([string[]]$Servers, [string]$InFolder, [string]$OutFolder, [bool]$SkipDeletePending, [bool]$ComparePrevious)

if ($Servers.Length -lt 2 -and [string]::IsNullOrEmpty($InFolder))
{
    Write-Host "You must provide at least 2 server names."
    return
}

if ([string]::IsNullOrEmpty($OutFolder))
{
    $OutFolder = split-path -parent $MyInvocation.MyCommand.Definition
}

Get-Job | Remove-Job

$startTime = Get-Date
$jobs = @()

if ($Servers.Length -gt 0)
{
    foreach ($server in $Servers)
    {
        $scriptBlock = {
            param($serverArg)
            $WarningPreference = "SilentlyContinue"
            Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$serverArg/powershell" -Authentication Kerberos) | Out-Null
            Get-PublicFolderStatistics -Server $server -ResultSize Unlimited | Select-Object FolderPath,ItemCount,IsDeletePending,DatabaseName
        }
        
        $jobs += Start-Job $scriptBlock -ArgumentList $server -Name $server
    }
}
else 
{
    foreach ($file in (Get-ChildItem -Path $InFolder -Filter *.csv))
    {
        $scriptBlock = {
            param($fileArg)
            return Import-Csv $fileArg.FullName
        }

        $jobs += Start-Job $scriptBlock -ArgumentList $file -Name $file
    }
}

Write-Host "Waiting for jobs to complete..."
$folders = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.Dictionary[string, int]]'
$dbNames = @()
$startingJobCount = $jobs.Length
$alreadyDoneCount = 0
while (Get-Job)
{
    $completedJobs = Get-Job | Where-Object { $_.State -ne "Running" }
    if ($completedJobs -eq $null)
    {
        Write-Progress -Activity "Waiting for job to finish." -PercentComplete ($alreadyDoneCount / $startingJobCount * 100) -Status ("$alreadyDoneCount / " + $startingJobCount + " jobs processed.")
        Start-Sleep 1
        continue
    }

    $completedJobs
    foreach ($job in $completedJobs)
    {
        Write-Progress -Activity ("Receiving job " + $job.Name) -PercentComplete ($alreadyDoneCount / $startingJobCount * 100) -Status ("$alreadyDoneCount / " + $startingJobCount + " jobs processed.")
        $statistics = Receive-Job $job
        Remove-Job $job
        $total = $statistics.Count
        if ($total -lt 1)
        {
            continue
        }

        $dbNames += $statistics[0].DatabaseName
        $current = 0
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $lastProgressMsec = 0
        foreach ($stat in $statistics)
        {
            ++$current
            if ($timer.ElapsedMilliseconds - $lastProgressMsec -gt 999)
            {
                Write-Progress -Activity ("Processing statistics from " + $stat.DatabaseName) -PercentComplete ($current / $total * 100) -Status "$current / $total"
                $lastProgressMsec = $timer.ElapsedMilliseconds
            }

            if ($SkipDeletePending -eq $true -and $stat.IsDeletePending -eq $true)
            {
                continue
            }
        
            $itemCountForReplicas = $null
            if ($folders.TryGetValue($stat.FolderPath, [ref]$itemCountForReplicas))
            {
                $itemCount = $null
                if ($itemCountForReplicas.TryGetValue($stat.DatabaseName, [ref]$itemCount))
                {
                    Write-Warning ("Duplicate folder path: " + $stat.FolderPath)
                }
                else
                {
                    $itemCountForReplicas.Add($stat.DatabaseName, $stat.ItemCount)
                }
            }
            else
            {
                $itemCountForReplicas = New-Object 'System.Collections.Generic.Dictionary[string, int]'
                $itemCountForReplicas.Add($stat.DatabaseName, $stat.ItemCount)
                $folders.Add($stat.FolderPath, $itemCountForReplicas)
            }
        }

        $timer.Stop()
        $alreadyDoneCount = $startingJobCount - (Get-Job).Length
    }
}

Write-Host
Write-Host "Comparing item counts..."
$previousData = $null
if ($ComparePrevious)
{
    $previousCsvFile = Get-ChildItem *.csv | Sort-Object CreationTime -Descending | Select-Object -First 1
    if ($previousCsvFile -eq $null)
    {
        Write-Warning "ComparePrevious is true, but no previous CSV file was found in $OutFolder"
    }
    else 
    {
        $previousData = Import-Csv $previousCsvFile
    }
}

$dbNames = $dbNames | Sort-Object
$foldersNotInSync = @()
$total = $folders.Keys.Count
$current = 0
$timer = [System.Diagnostics.Stopwatch]::StartNew()
$lastProgressMsec = 0
foreach ($folderPath in $folders.Keys)
{
    ++$current
    if ($timer.ElapsedMilliseconds - $lastProgressMsec -gt 999)
    {
        Write-Progress -Activity "Comparing item counts" -PercentComplete ($current / $total * 100) -Status "$current / $total"
        $lastProgressMsec = $timer.ElapsedMilliseconds
    }

    $folder = $folders[$folderPath]
    $measurement = $folder.Values | Measure-Object -Maximum -Minimum
    if ($measurement.Minimum -ne $measurement.Maximum)
    {
        if ($previousData -ne $null)
        {
            $previousEntry = $previousData | Where-Object { $_.FolderPath -eq $folderPath }
            if ($previousEntry -ne $null)
            {
                $previousMeasurement = $previousEntry.Value | Measure-Object -Minimum
                if ($previousMeasurement.Minimum -ne $measurement.Minimum)
                {
                    continue
                }
            }
        }

        $reportObj = New-Object PSObject
        Add-Member -InputObject $reportObj -MemberType NoteProperty -Name FolderPath -Value $folderPath
        foreach ($dbName in $dbNames)
        {
            $itemCount = $null
            if ($folder.TryGetValue($dbName, [ref]$itemCount))
            {
                Add-Member -InputObject $reportObj -MemberType NoteProperty -Name $dbName -Value $itemCount
            }
            else
            {
                Add-Member -InputObject $reportObj -MemberType NoteProperty -Name $dbName -Value ""
            }
        }
        
        $foldersNotInSync += $reportObj
    }
}

$timer.Stop()
$endTime = Get-Date

$fileName = "ReplicationReport" + [DateTime]::Now.ToString("yyyyMMddHHmm") + ".csv"
$fullName = Join-Path $OutFolder $fileName
$foldersNotInSync | Sort-Object FolderPath | Export-Csv $fullName -NoTypeInformation -Encoding Unicode

Write-Host "Folders found: $total"
Write-Host "Folders not in sync: " $foldersNotInSync.Length
Write-Host "Started: $startTime"
Write-Host "Ended: $endTime"
Write-Host "Duration: " ($endTime - $startTime)
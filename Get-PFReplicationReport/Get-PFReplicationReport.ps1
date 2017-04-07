# Get-PFReplicationReport
# 
# PF replication report script with primary goals of performance and terse output.
# 
# This script will run Get-PublicFolderStatistics against all specified servers
# simultaneously, and will begin processing the results as soon as any of those
# tasks complete.
# 
# The output ONLY contains folders where there is an item count difference between
# replicas. Folders where all replicas have the same item count are not included.
# 
# Syntax examples:
# 
# Display the report in the shell:
# 
# .\Get-ReplicationReport -Servers SERVER1,SERVER2,SERVER3
# 
# Write the report to a CSV file:
# 
# .\Get-ReplicationReport -Servers SERVER1,SERVER2,SERVER3 -Outfile C:\replreport.csv
# 
# Skip any folders in a DeletePending state (useful when you have a lot of folders
# pending deletion and generating false positives in the report):
# 
# .\Get-ReplicationReport -Servers SERVER1,SERVER2,SERVER3 -Outfile C:\replreport.csv -SkipDeletePending $true
# 

param([string[]]$Servers, [string]$Outfile, [bool]$SkipDeletePending)

if ($Servers.Length -lt 2)
{
    Write-Host "You must provide at least 2 server names."
    return
}

Get-Job | Remove-Job

$startTime = Get-Date
$jobs = @()

foreach ($server in $Servers)
{
    $scriptBlock = {
        param($serverArg)
        $WarningPreference = "SilentlyContinue"
        Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$serverArg/powershell" -Authentication Kerberos) | Out-Null
        Get-PublicFolderStatistics -Server $server -ResultSize Unlimited | Select FolderPath,ItemCount,IsDeletePending,DatabaseName
    }
    
    $jobs += Start-Job $scriptBlock -ArgumentList $server -Name $server
}

Write-Host "Waiting for jobs to complete..."
$folders = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.Dictionary[string, int]]'
$dbNames = @()
while (Get-Job)
{

    $completedJobs = Get-Job | WHERE { $_.State -eq "Completed" }
    if ($completedJobs -eq $null)
    {
        $alreadyDoneCount = $Servers.Length - (Get-Job).Length
        Write-Progress -Activity "Waiting for Get-PublicFolderStatistics to finish." -PercentComplete ($alreadyDoneCount / $Servers.Length * 100) -Status ("$alreadyDoneCount / " + $Servers.Length + " servers complete.")
        Start-Sleep 1
        continue
    }

    $completedJobs
    foreach ($job in $completedJobs)
    {
        Write-Host "Receiving job" $job.Name
        $statistics = Receive-Job $job
        Remove-Job $job
        $total = $statistics.Count
        if ($total -lt 1)
        {
            continue
        }

        $dbNames += $statistics[0].DatabaseName
        $current = 0
        foreach ($stat in $statistics)
        {
            ++$current
            Write-Progress -Activity ("Processing statistics from " + $stat.DatabaseName) -PercentComplete ($current / $total * 100) -Status "$current / $total"
            if ($SkipDeletePending -and $stat.IsDeletePending)
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
    }
}

Write-Host
Write-Host "Comparing item counts..."
$dbNames = $dbNames | Sort
$foldersNotInSync = @()
$total = $folders.Keys.Count
$current = 0
foreach ($folderPath in $folders.Keys)
{
    ++$current
    Write-Progress -Activity "Comparing item counts" -PercentComplete ($current / $total * 100) -Status "$current / $total"
    $folder = $folders[$folderPath]
    $measurement = $folder.Values | measure -Maximum -Minimum
    if ($measurement.Minimum -ne $measurement.Maximum)
    {
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

$endTime = Get-Date

if ($Outfile.Length -gt 0)
{
    $foldersNotInSync | Sort FolderPath | Export-Csv $Outfile -NoTypeInformation -Encoding Unicode
}
else
{
    $foldersNotInSync
}

Write-Host "Folders found: $total"
Write-Host "Folders not in sync: " $foldersNotInSync.Length
Write-Host "Started: $startTime"
Write-Host "Ended: $endTime"
Write-Host "Duration: " ($endTime - $startTime)

# Sync-ReplState.ps1
#
# After running a ReplState repair, it is often necessary to change
# the replica list on the affected folders by removing and re-adding
# the database in order to sync the data in the ReplState table with
# the Folders table.
#
# The purpose of this script is to scan the application log for the
# latest ReplState fix, get a list of the affected FIDs, then find
# those FIDs and modify the replica list, thus causing it to sync up.
# 
# Example syntax:
# 
# .\Sync-ReplState.ps1 -Server CONTOSO1
# This syntax will report the folders repaired on CONTOSO1, but 
# will not make any changes.
#
# .\Sync-ReplState.ps1 -Server CONTOSO1 -MakeChanges $true
# This syntax will report each folder, and then prompt asking if
# you want the script to automatically update the replica list.
#
# .\Sync-ReplState.ps1 -Server CONTOSO1 -MakeChanges $true -DoNotPrompt $true
# This syntax will automatically update the replica lists for all
# affected folders without prompting.

param($Server, $MakeChanges, $DoNotPrompt, $DoAllFolders)

$serverName = $Server
if ($serverName.Length -lt 1)
{
	"Server name must be specified."
	return
}

$server = Get-ExchangeServer $serverName
$db = Get-PublicFolderDatabase -Server $serverName
$allOtherDbs = Get-PublicFolderDatabase -IncludePreExchange2010
$randomReplica = $allOtherDbs[0]
if ($randomReplica -eq $db)
{
	$randomReplica = $allOtherDbs[1]
}
	
function DoFolder($folderStatistic)
{
	$folder = Get-PublicFolder $folderStatistic -Server $serverName
	$originalReplicaList = $folder.Replicas
	"    Original replica list:"
	$originalReplicaList | ft Name
	$newReplicaList = $null
	if ($originalReplicaList.Contains($db.Identity))
	{
		if ($folder.Replicas.Count -lt 2)
		{
			# The folder only has a replica on the local server, so
			# add a random replica temporarily.
			$newReplicaList = $originalReplicaList + $randomReplica.Identity
		}
		else
		{
			$newReplicaList = $originalReplicaList - $db.Identity
		}
	}
	else
	{
		$newReplicaList = $originalReplicaList + $db.Identity
	}

	"    Will temporarily change replica list to:"
	$newReplicaList | ft Name

	if ($MakeChanges)
	{
		if ($DoNotPrompt -ne $true)
		{
			$prompt = Read-Host "Would you like the script to do this now (Y/N)? "
			if ($prompt.ToUpper() -ne "Y")
			{
				"Skipping this folder."
				return
			}
		}

		Set-PublicFolder $folder -Replicas $newReplicaList -Server $server.Identity

		"    Replica list was updated. Now changing it back to:"
		$originalReplicaList | ft Name
		Set-PublicFolder $folder -Replicas $originalReplicaList -Server $server.Identity
		"    Replica list update finished."
	}
	else
	{
		"    Script is running in read-only mode. Set -makeChanges $true to automatically update replica lists."
	}
}

function CheckFolder($folder)
{
	$entryIdString = $folder.EntryId.ToString()
	$globcntString = $entryIdString.Substring(0, $entryIdString.Length - 4)
	$globcntString = $globcntString.Substring($globcntString.Length - 12)
	$globcntString = $globcntString.TrimStart(@('0'))
	if ($globcnts.Contains($globcntString))
	{
		"Found repaired folder: " + $folder.FolderPath
		DoFolder $folder
		return
	}
}

if ($DoAllFolders)
{
	"Retrieving the public folders..."

	$localPublicFolders = Get-PublicFolderStatistics -Server $serverName -resultSize unlimited

	"DoAllFolders is $true. Doing all folders..."
	
	foreach ($folder in $localPublicFolders)
	{
		"Doing folder " + $folder.FolderPath
		DoFolder $folder
	}
}
else
{
	$applicationLog = new-object System.Diagnostics.EventLog("Application", $serverName)
	$eventCount = $applicationLog.Entries.Count
	
	"Application log on " + $serverName + " has " + $eventCount.ToString() + " events."
	if ($eventCount -lt 1)
	{
		return
	}
	
	$fids = new-object 'System.Collections.Generic.HashSet[string]'
	
	function AddEventToFidList($event, $fidList)
	{
		$lines = $event.Message.Split(@("`n"))
		foreach ($line in $lines)
		{
			$columns = $line.Split(@(','))
			if ($columns.Length -gt 4)
			{
				$fidList.Add($columns[2].Trim()) | Out-Null
			}
		}
	}
	
	"Finding latest ReplState repair..."
	$repairEventIndex = -1
	for ($x = ($eventCount - 1); $x -ge 0; $x--)
	{
		$thisEvent = $applicationLog.Entries[$x]
	
		if ($thisEvent.EventID -eq 10063 -and $thisEvent.Message.Contains("ReplState"))
		{
			"Found repair event 10063 at " + $thisEvent.TimeGenerated.ToString()
			AddEventToFidList $thisEvent $fids
		}
		elseif ($thisEvent.EventID -eq 10064)
		{
			break
		}
	}
	
	if ($fids.Count -lt 1)
	{
		"Could not find an event with ID 10063 which contained a ReplState repair in the latest repair."
		return
	}
	
	"Found " + $fids.Count.ToString() + " repaired FIDs."
	
	$globcnts = new-object 'System.Collections.Generic.HashSet[string]'
	foreach ($fid in $fids)
	{
		$globcnts.Add($fid.Substring($fid.IndexOf(@('-')) + 1)) | Out-Null
	}
	

	"Retrieving the public folders..."
	
	$localPublicFolders = Get-PublicFolderStatistics -Server $serverName -resultSize unlimited
	
	"Searching for folders that match the reported FIDs..."
	
	$currentCount = 0
	$totalCount = $localPublicFolders.Count
	foreach ($localFolder in $localPublicFolders)
	{
		$percentComplete = [Math]::Floor($currentCount / $totalCount * 100)
		Write-Progress -Activity "Syncing ReplState" -Status "$currentCount / $totalCount" -PercentComplete $percentComplete -CurrentOperation $localFolder.Identity.ToString()
		CheckFolder $localFolder
	}
}

"Done!"
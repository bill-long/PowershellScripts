# Restore-PFPermissions2.ps1
# 
# The purpose of this script is to take a backup of PF permissions
# that was generated on Exchange 2010 using this command:
# 
# Get-PublicFolder -Recurse | Get-PublicFolderClientPermission | Select-Object Identity,User -ExpandProperty AccessRights | Export-CliXML $home\desktop\Legacy_PFPerms.xml
# 
# And to put those permissions back in place on Exchange 2013.
# 
# Note that directly reading the CliXml file is problematic for a few reasons.
# Import-CliXml does not restore all the properties when it is run from 
# Exchange Management Shell. Directly reading the Xml with Get-Content also
# results in certain properties being inaccessible.
# 
# Therefore, the clixml export must be converted before this script is used.
# To convert it, run this command from plan old Powershell (not EMS):
# 
# Import-CliXml C:\Legacy_PFPerms.xml | Select-Object Identity,User,Permission | Export-Csv C:\SanePermissionsExport.csv
# 
# Then pass that file to this script:
# 
# .\Restore-PFPermissions2 C:\SanePermissionsExport.csv

param($pfPermissionFile)

"Loading permissions file..."
$permsToRestore = Import-Csv $pfPermissionFile
if ($permsToRestore -eq $null)
{
    return
}

$sess = Get-PSSession | WHERE { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
if ($sess -eq $null)
{
    . 'C:\Program Files\Microsoft\Exchange Server\v15\Bin\RemoteExchange.ps1'
    Connect-ExchangeServer -Auto
}

$sess = Get-PSSession | WHERE { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
if ($sess -eq $null)
{
    "Could not get a session to Exchange."
    return
}

$currentFolder = ""
$currentFolderPerms = $null
for ($x = 0; $x -lt $PermsToRestore.Count; $x++)
{
    $thisPermissionEntry = $PermsToRestore[$x]

    $folderPath = $thisPermissionEntry.Identity.ToString()
    $user = $thisPermissionEntry.User.ToString()
    $permsThisUserHadBefore = $thisPermissionEntry.Permission.Split(",") | % { $_.Trim() }

    if ($folderPath -ne $currentFolder.ToString() -and $folderPath -ne "\")
    {
        $currentFolder = ""
        $currentFolderPerms = $null
        $usersProcessedForThisFolder = @()
        "Finding folder: " + $folderPath
        $desiredFolder = Get-PublicFolder $folderPath
        if ($desiredFolder -eq $null)
        {
            continue
        }

        $currentFolder = $desiredFolder.Identity.ToString()
        "Found folder: " + $currentFolder
        $currentFolderPerms = Get-PublicFolderClientPermission $currentFolder | Select-Object Identity,User,AccessRights
    }

    "    User: " + $user
    "    Permissions this user had before: "
    $outString = "        "
    $permsThisUserHadBefore | % { $outString += $_ + " " }
    $outString

    # Now make a list of the perms the user has now.
    $permsThisUserHasNow = @()
    for ($y = 0; $y -lt $currentFolderPerms.Length; $y++)
    {
        $match = $false
        if ($currentFolderPerms[$y].User.ToString() -eq $user.ToString())
        {
            $match = $true
        }

        if ($currentFolderPerms[$y].User.ADRecipient -ne $null)
        {
            if ($currentFolderPerms[$y].User.ADRecipient.Identity.ToString() -eq $user.ToString())
            {
                $match = $true
            }
        }

        if ($match)
        {
            $permsThisUserHasNow = $currentFolderPerms[$y].AccessRights | % { $_.ToString() }
            break
        }
    }

    if ($permsThisUserHasNow.Length -gt 0)
    {
        "    Permissions this user currently has:"
        "        " + $permsThisUserHasNow
        "    Removing current permissions..."
        Remove-PublicFolderClientPermission $currentFolder -User $user -Confirm:$false
    }
    else
    {
        "    This user currently has no permissions on the folder."
    }

    if ($permsThisUserHadBefore.Length -gt 0)
    {
        "    Adding these permissions: "
        $outString = "        "
        $permsToAdd | % { $outString += $_ + " " }
        $outString
        Add-PublicFolderClientPermission $currentFolder -User $user -AccessRights $permsThisUserHadBefore
    }
}
#################################################################################
# GenericDoAllPublicFoldersScript.ps1
#
# Generic script to do something to all public folders via EWS. Just change
# DoFolder to do what you want.

param([string]$FolderPath, [string]$HostName, [string]$UserName, [boolean]$Recurse)

#################################################################################
# Update the path below to match the actual path to the EWS managed API DLL.
#

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"

#
#################################################################################

#################################################################################
# Change this function to do whatever you want to do to each public folder
#

function DoFolder($folder, $path)
{
    "Doing folder: " + $path
    # Do whatever
}

#
#################################################################################

$FolderPath = $FolderPath.Trim(@('\'))

if ($HostName -eq "")
{
    $HostName = Read-Host "Hostname for EWS endpoint (leave blank to attempt Autodiscover)"
}

if ($UserName -eq "")
{
    $UserName = Read-Host "User (UPN format)"
}

$password = $host.ui.PromptForCredential("Credentials", "Please enter your password to authenticate to EWS.", $UserName, "").GetNetworkCredential().Password

# If a URL was specified we'll use that; otherwise we'll use Autodiscover 
$exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010) 
$exchService.Credentials = new-object System.Net.NetworkCredential($UserName, $password, "") 
if ($HostName -ne "") 
{ 
    ("Using EWS URL " + "https://" + $HostName + "/EWS/Exchange.asmx") 
    $exchService.Url = new-object System.Uri(("https://" + $HostName + "/EWS/Exchange.asmx")) 
} 
else
{ 
    ("Autodiscovering " + $UserName + "...")
    $exchService.AutoDiscoverUrl($UserName, {$true}) 
}

if ($exchService.Url -eq $null) 
{ 
    return 
}

$pfsRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot) 
$folder = $pfsRoot

if ($FolderPath.Length -gt 0)
{
    $tinyView = new-object Microsoft.Exchange.WebServices.Data.FolderView(2) 
    $displayNameProperty = [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName
    $folderPathSplits = $FolderPath.Split(@('\')) 
    for ($x = 0; $x -lt $folderPathSplits.Length;$x++) 
    { 
        $filter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($displayNameProperty, $folderPathSplits[$x]) 
        $results = $folder.FindFolders($filter, $tinyView) 
        if ($results.TotalCount -gt 1) 
        { 
             ("Ambiguous name: " + $folderPathSplits[$x]) 
             return 
        } 
        elseif ($results.TotalCount -lt 1) 
        { 
             ("Folder not found: " + $folderPathSplits[$x]) 
             return 
        } 
        $folder = $results.Folders[0] 
    }
}

function DoSubfoldersRecursive($folder, $path)
{
    $folderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(2147483647) 
    $subfolders = $folder.FindFolders($folderView)    
    foreach ($subfolder in $subfolders) 
    { 
        try 
        {
            DoFolder $subfolder ($path + "\" + $subfolder.DisplayName)
            DoSubfoldersRecursive $subfolder ($path + "\" + $subfolder.DisplayName) 
        } 
        catch { "Error processing folder: " + $subfolder.DisplayName } 
    }
}

if ($Recurse)
{
    if ($FolderPath.Length -gt 0)
    {
        DoFolder $folder $FolderPath
    }

    DoSubfoldersRecursive $folder $FolderPath
}
else
{
    DoFolder $folder $FolderPath
}

"Done!"
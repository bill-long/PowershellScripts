#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages
#
#################################################################################

# Delete-TNEFProps.ps1
# 
# Delete-TNEFProps.ps1 is a script intended to fix the Exchange Public Folder Replication 
# issue described here: http://bill-long.com/2014/01/16/public-folder-replication-fails-with-tnef-violation-status-0x00008000/ 

param([string]$FolderPath, [string]$HostName, [string]$UserName, [boolean]$Fix, [boolean]$Verbose, [boolean]$Recurse)

#
# Update the path below to match the actual path to the EWS managed API DLL.
#

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"

#
#
#

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

if ($Fix)
{
    "-Fix is TRUE. The TNEF properties will be deleted."
}
else
{
    "-Fix is FALSE. Running in read-only mode."
}

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

function DoFolder($folder, $path)
{
    $tnefProp1 = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1204, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ShortArray)
    $tnefProp2 = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1205, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::ShortArray)
    $subjectProperty = [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject

    $arrayOfPropertiesToLoad = @($subjectProperty, $tnefProp1, $tnefProp2)
    $itemviewPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet($arrayOfPropertiesToLoad)

    $offset = 0; 
    $view = new-object Microsoft.Exchange.WebServices.Data.ItemView(100, $offset) 
    $view.PropertySet = $itemviewPropertySet

    "Checking folder: " + $path

    while (($results = $folder.FindItems($view)).Items.Count -gt 0) 
    {
        foreach ($item in $results) 
        {
            $propValue1 = $null
            $propValue2 = $null

            if ($Verbose)
            {
                "Checking item: " + $item.Subject
            }

            $foundProperty1 = $item.TryGetProperty($tnefProp1, [ref]$propValue1)
            $foundProperty2 = $item.TryGetProperty($tnefProp2, [ref]$propValue2)

            if ($foundProperty1 -or $foundProperty2)
            {
                "TNEF props found on item: " + $item.Subject.ToString()

                if ($Fix)
                {
                    "    Removing TNEF properties..."

                    if ($foundProperty1)
                    {
                        $item.RemoveExtendedProperty($tnefProp1) | Out-Null
                    }

                    if ($foundProperty2)
                    {
                        $item.RemoveExtendedProperty($tnefProp2) | Out-Null
                    }

                    $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)

                    "    Finished removing TNEF properties from this item."
                }
            }
        }
     
        $offset += $results.Items.Count 
        $view = new-object Microsoft.Exchange.WebServices.Data.ItemView(100, $offset)
        $view.PropertySet = $itemviewPropertySet
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

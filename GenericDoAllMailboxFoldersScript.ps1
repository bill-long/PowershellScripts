#########################################
# GenericDoAllMailboxFoldersScript.ps1
# 

param([string]$HostName, [string]$UserName, [string]$Mailbox, [string]$InputFile)

#########################################
# Update the path below to match the actual path to the EWS managed API DLL.
#

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"

#
#########################################

if ($InputFile -eq "" -and $Mailbox -eq "")
{
    $Mailbox = Read-Host "Mailbox to process"
}
elseif ($InputFile -ne "" -and $Mailbox -ne "")
{
    "-Mailbox and -InputFile are mutually exclusive. Only one may be specified."
    return
}

$mailboxList = new-object 'System.Collections.Generic.List[string]'
if ($InputFile -ne "")
{
    $inputReader = new-object System.IO.StreamReader($InputFile)
    while ($null -ne ($buffer = $inputReader.ReadLine()))
    {
        if ($buffer.Length -gt 0)
        {
            $mailboxList.Add($buffer)
        }
    }
}
else
{
    $mailboxList.Add($Mailbox)
}

("Number of mailboxes to process: " + $mailboxList.Count.ToString())
if ($mailboxList.Count -lt 1)
{
    return
}

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
    ("Autodiscovering " + $mailboxList[0] + "...")
    $exchService.AutoDiscoverUrl($mailboxList[0], {$true}) 
}

if ($exchService.Url -eq $null) 
{ 
    return 
}

#########################################
# Properties we care about
# This is not actually used in this generic script
# 
$subjectProperty = [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject
$itemClassProperty = [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass
$dateTimeReceivedProperty = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived

$arrayOfPropertiesToLoad = @($subjectProperty, $itemClassProperty, $dateTimeReceivedProperty)
$propertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet($arrayOfPropertiesToLoad)
#
#########################################

function DoFolder($folder, $path)
{
    ("Processing folder: " + $path)
    $itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(100)
    while (($folderItems = $folder.FindItems($itemView)).Items.Count -gt 0)
    {
        foreach ($item in $folderItems)
        {
            # Do something
        }
       
        $offset += $folderItems.Items.Count
        $itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(100, $offset)
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
            if ($subfolder.SearchParameters -ne $null)
            {
                # If it's a search folder, we're not interested
                continue;
            }

            DoFolder $subfolder ($path + "\" + $subfolder.DisplayName)
            DoSubfoldersRecursive $subfolder ($path + "\" + $subfolder.DisplayName) 
        } 
        catch { "Error processing folder: " + $subfolder.DisplayName } 
    }
}

function DoMailbox($thisMailbox)
{
    $rootFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root
    $rootId = new-object Microsoft.Exchange.WebServices.Data.FolderId($rootFolderName, $mbx)
    $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $rootId)
    
    if ($rootFolder -eq $null)
    {
        ("Error. Could not open mailbox: " + $emailAddress)
        return
    }
    
    "Opened mailbox: " + $emailAddress
    DoSubfoldersRecursive $rootFolder ""
}

foreach ($emailAddress in $mailboxList)
{
    $mbx = new-object Microsoft.Exchange.WebServices.Data.Mailbox($emailAddress)
    DoMailbox($mbx)
}
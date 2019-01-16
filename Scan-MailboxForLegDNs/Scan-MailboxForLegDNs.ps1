# Scan-MailboxForLegDNs.ps1

# From my blog post Sep 16, 2011:
#
# I recently worked with a customer who had inadvertently deleted all their user accounts (and thus their Exchange mailboxes), 
# and with no backup available, they had to recreate them. Talk about a nightmare! After they did so, they were able to get 
# their email back, but they discovered that replying to email messages from before the problem resulted in a non-delivery report.
# 
# This is because of a fairly well-documented behavior. Email addresses in the From, To, and other fields are resolved to
# legacyExchangeDNs and stored in the message. When you reply to a message, we expect to be able to resolve that legacyExchangeDN.
# If we can’t, it causes an NDR. In various migration scenarios where the legacyExchangeDN of a user changes, we populate the user’s
# proxyAddresses with an X500 address that contains the old legacyExchangeDN. This allows the old value to resolve, preserving the
# ability to reply to old messages.
#
# In this case, when the users were recreated, they got new legacyExchangeDNs, which broke the ability to reply to old email. We
# needed to somehow get the old leg DNs back, but we didn’t know what they all were, and having to manually poke around in mailboxes
# looking for them was not realistic.
#
# To solve this problem, I wrote this script to scan a mailbox and output any unresolved legacyExchangeDNs. Note that EWS will not
# return the full recipient information as part of a FindItems() call – you only get back some basic information, which wasn’t enough
# to tell me if there was an unresolved legacyExchangeDN. To solve this, I had to actually bind to each individual email message. I
# did this using a specific property set of only a few properties that I was interested in, but it still made the script quite slow.
#
# The script will scan the Inbox and any subfolders, as well as Sent Items and Deleted Items, looking for unresolved leg DNs. Any it
# finds will be written out to a CSV file. It caches what it finds so that duplicates are not written to the CSV.

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"

$hostName = Read-Host "Hostname for EWS endpoint (leave blank to attempt Autodiscover)" 
$outputFile = Read-Host "Output file name" 
$emailAddress = Read-Host "Email address for authentication" 
$password = $host.ui.PromptForCredential("Credentials", "Please enter your password to authenticate to EWS.", $emailAddress, "").GetNetworkCredential().Password

# If specified, we'll try to open this mailbox instead of the one that authenticated 
$otherMailboxSmtp = Read-Host "SMTP address of other mailbox (optional)"

# Initialize the output file 
Set-Content -Path $outputFile -Value "Display Name,LegacyExchangeDN"

# Make variables for the properties to make them easier to type, 
# then stick them into a PropertySet 
$toRecipientsProperty = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ToRecipients 
$ccRecipientsProperty = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::CcRecipients 
$fromProperty = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From 
$subjectProperty = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject 
$arrayOfPropertiesToLoad = @($toRecipientsProperty, $ccRecipientsProperty, $fromProperty, $subjectProperty) 
$propertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet($arrayOfPropertiesToLoad)

# Here's where we'll store the ones we found to avoid duplicates in the CSV 
$legDNsFound = new-object 'System.Collections.Generic.List[string]'

# This function checks an individual EmailAddress to see if it's an unresolved legacyExchangeDN 
function CheckAddress($emailAddress) 
{ 
    if ($emailAddress.RoutingType -eq "EX") 
    { 
        ("        Legacy DN: " + $emailAddress.Address) 
        ("        Display Name: " + $emailAddress.Name) 
        if (!($legDNsFound.Contains($emailAddress.Address.ToLower()))) 
        { 
            $legDNsFound.Add($emailAddress.Address.ToLower()) 
            Add-Content -Path $outputFile -Value ($emailAddress.Name.Replace(",", ".") + "," + $emailAddress.Address) 
        } 
    } 
}

# This function loops through the items in a folder 
function ProcessFolder($folder) 
{ 
    ("Scanning folder: " + $folder.DisplayName) 
    $itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(100) 
    while (($folderItems = $folder.FindItems($itemView)).Items.Count -gt 0) 
    { 
        foreach ($item in $folderItems) 
        { 
            if ($item.GetType() -eq [Microsoft.Exchange.WebServices.Data.EmailMessage]) 
            { 
                $message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchService, $item.Id, $propertySet) 
                ("    " + $message.Subject) 
                foreach ($emailAddress in $message.ToRecipients) 
                { 
                    CheckAddress($emailAddress) 
                } 
                
                foreach ($emailAddress in $message.CcRecipients) 
                { 
                    CheckAddress($emailAddress) 
                } 
                
                CheckAddress($message.From) 
            } 
        } 
        
        $offset += $folderItems.Items.Count 
        $itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(100, $offset) 
    } 
}

# This function recursively processes subfolders 
function DoSubfoldersRecursive($folder) 
{ 
    if ($folder.ChildFolderCount -gt 0) 
    { 
        $folderView = new-object Microsoft.Exchange.WebServices.Data.FolderView($folder.ChildFolderCount) 
        $subfolders = $folder.FindFolders($folderView) 
        foreach ($subfolder in $subfolders) 
        { 
            ProcessFolder($subfolder) 
            DoSubfoldersRecursive($subfolder) 
        } 
    } 
}

# Here's where we try to connect 
# If a URL was specified we'll use that; otherwise we'll use Autodiscover 
$exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1) 
$exchService.Credentials = new-object System.Net.NetworkCredential($emailAddress, $password, "") 
if ($hostName -ne "") 
{ 
    ("Using EWS URL:" + "https://" + $hostName + "/EWS/Exchange.asmx") 
    $exchService.Url = new-object System.Uri(("https://" + $hostName + "/EWS/Exchange.asmx")) 
} 
elseif ($otherMailboxSmtp -ne "") 
{ 
    ("Autodiscovering " + $otherMailboxSmtp + "...") 
    $exchServce.AutoDiscoverUrl($otherMailboxSmtp, {$true}) 
} 
else 
{ 
    ("Autodiscovering " + $emailAddress + "...") 
    $exchService.AutodiscoverUrl($emailAddress, {$true}) 
}

if ($exchService.Url -eq $null) 
{ 
    return 
}

$mailbox = new-object Microsoft.Exchange.WebServices.Data.Mailbox($emailAddress)

# If some other mailbox was specified, open that one instead. 
if ($otherMailboxSmtp -ne "") 
{ 
    $mailbox = new-object Microsoft.Exchange.WebServices.Data.Mailbox($otherMailboxSmtp) 
}

# Create some variables for the folder names so they're easier to type 
$inboxFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox 
$sentItemsFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems 
$deletedItemsFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems

# We'll bind to each folder by instantiating a FolderId that points to the mailbox we want.

# First, scan the inbox 
$inboxId = new-object Microsoft.Exchange.WebServices.Data.FolderId($inboxFolder, $mailbox) 
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $inboxId) 
ProcessFolder($inbox) 
    
# Now any subfolders 
DoSubfoldersRecursive($inbox) 
    
# Now Sent Items 
$sentItemsId = new-object Microsoft.Exchange.WebServices.Data.FolderId($sentItemsFolder, $mailbox) 
$sentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $sentItemsId) 
ProcessFolder($sentItems) 
    
# Now Deleted Items 
$deletedItemsId = new-object Microsoft.Exchange.WebServices.Data.FolderId($deletedItemsFolder, $mailbox) 
$deletedItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $deletedItemsId) 
ProcessFolder($deletedItems)

"Done!"

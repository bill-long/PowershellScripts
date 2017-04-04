# SetPFDiscoveryMailbox.ps1
# 
# Syntax:
# 
# $mailboxes = Get-Mailbox | .\SetPFDiscoveryMailbox.ps1
# 
# You can use the usual parameters to filter the mailbox result, such
# as only returning mailboxes for a certain server:
# 
# $mailboxes = GetMailbox -Server EXCH1 | .\SetPFDiscoveryMailbox.ps1
#
############################################################################

############################################################################
#
# This script will not make any changes unless you flip this value to $true
#

$updateValues = $false

#
############################################################################

# First, get all the remote PF mailboxes
$remotePFMailboxes = (Get-OrganizationConfig).RemotePublicFolderMailboxes | Get-Mailbox
if ($remotePFMailboxes -eq $null)
{
    "No remote public folder mailboxes are set on the OrganizationConfig."
    return
}

# Now get the PF databases
foreach ($mbx in $remotePFMailboxes)
{
    $pfdb = Get-PublicFolderDatabase -Server $mbx.ServerName
    $mbx | Add-Member NoteProperty PFDatabase $pfdb.DistinguishedName
}

# Get all mailbox and PF databases in order to save time later
$mailboxDatabases = Get-MailboxDatabase -IncludePreExchange2013
$publicFolderDatabases = Get-PublicFolderDatabase

foreach ($mailbox in $input)
{
    # Skip mailboxes on Exchange 2010 or older
    if ($mailbox.ExchangeVersion.ExchangeBuild.Major -lt 15)
    {
        continue
    }

    # First, find the mailbox's default PF mailbox value. We have to check AD because
    # the cmdlet returns a random value when it's empty.
    $dn = $mailbox.DistinguishedName
    $adObject = [ADSI]("GC://" + $dn)
    $currentDefaultPFMailbox = $null
    if ($adObject.Properties["msExchPublicFolderMailbox"].Count -gt 0)
    {
        $currentDefaultPFMailbox = Get-Mailbox $adObject.Properties["msExchPublicFolderMailbox"][0].ToString()
    }

    # Now figure out what it *should* be 
    # What mailbox database is this mailbox on?
    $mailboxDatabase = $mailboxDatabases | WHERE { $_.DistinguishedName -eq $mailbox.Database.DistinguishedName }
    # What is the PublicFolderDatabase value on that mailbox database?
    $defaultPFDatabase = $publicFolderDatabases | WHERE { $_.DistinguishedName -eq $mailboxDatabase.PublicFolderDatabase.DistinguishedName }
    # What is the PF discovery mailbox on that server?
    $correctPFDiscoveryMailbox = $remotePFMailboxes | WHERE { $_.ServerName -eq $defaultPFDatabase.ServerName }

    # OK, so is the current value correct?
    $valueIsCorrect = [string]::Equals($correctPFDiscoveryMailbox.DistinguishedName, $currentDefaultPFMailbox.DistinguishedName, "OrdinalIgnoreCase")
    if (!($valueIsCorrect))
    {
        if ($updateValues)
        {
            if ($correctPFDiscoveryMailbox -ne $null)
            {
                Set-Mailbox $mailbox -DefaultPublicFolderMailbox $correctPFDiscoveryMailbox
            }

            # Output the result
            $object = New-Object -TypeName PSObject
            $object | Add-Member NoteProperty Mailbox $mailbox
            $object | Add-Member NoteProperty OldValue $currentDefaultPFMailbox
            $object | Add-Member NoteProperty NewValue $correctPFDiscoveryMailbox
            $object
        }
        else 
        {
            # Output the result
            $object = New-Object -TypeName PSObject
            $object | Add-Member NoteProperty Mailbox $mailbox
            $object | Add-Member NoteProperty CurrentValue $currentDefaultPFMailbox
            $object | Add-Member NoteProperty CorrectValue $correctPFDiscoveryMailbox
            $object
        }
    }
}

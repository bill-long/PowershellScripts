# Fix-MailEnabled.ps1
# 
# See http://blogs.technet.com/b/bill_long/archive/2013/07/02/fixing-mail-enabled-public-folders-per-kb-977921.aspx 
# 
# The purpose of this script is to read an ExFolders or PFDAVAdmin
# property export of public folders, and fix the folders where the
# mail-enabled state is not consistent.
# 
# The export must include the following properties:
# PR_PF_PROXY, PR_PF_PROXY_REQUIRED, DS:proxyAddresses
#
# The export can optionally include the following
# properties, and if they are included, the script will preserve
# that setting:
# DS:msExchHideFromAddressLists
# 
# This script must be run from Exchange Management Shell.

# File is required, ExchangeServer and DC are optional
# Example syntax:
# .\Fix-MailEnabled C:\ExFoldersExport.txt
# .\Fix-MailEnabled C:\ExFoldersExport.txt CONTOSO1 DC01
# .\Fix-MailEnabled -File C:\ExFoldersExport.txt -ExchangeServer CONTOSO1 -DC DC01

param([string]$File, [string]$ExchangeServer, [string]$DC, [string]$emailAddress, [string]$smtpServer)

# Log file stuff

$ScriptPath = Split-Path -Path $MyInvocation.MyCommand.Path -Parent 
$Global:LogFileName = "FixMailEnabled"
$Global:LogFile  = $ScriptPath + "\" + $LogFileName + ".log"
$Global:ErrorLogFile = $ScriptPath + "\FixMailEnabled-Errors.log"

$sendEmailOnError = $false
if ($emailAddress -ne $null -and $emailAddress.Length -gt 0 -and $smtpServer -ne $null -and $smtpServer.Length -gt 0)
{
    $sendEmailOnError = $true
}

function Test-Transcribing
{

    $externalHost = $host.gettype().getproperty("ExternalHost", [reflection.bindingflags]"NonPublic,Instance").getvalue($host, @())
    try
    {
        $externalHost.gettype().getproperty("IsTranscribing", [reflection.bindingflags]"NonPublic,Instance").getvalue($externalHost, @())
    }
    catch
    {
        write-warning "This host does not support transcription."

    }

}

function writelog([string]$value = "")
{
    $Global:LogDate = Get-Date -uformat "%Y %m-%d %H:%M:%S"
    ("$LogDate $value")
}

function writeerror([string]$value = "")
{
    $Global:LogDate = Get-Date -uformat "%Y %m-%d %H:%M:%S"
    Add-Content -Path $Global:ErrorLogFile -Value ("$LogDate $value")
    if ($sendEmailOnError)
    {
        writelog("Sending email notification...")
        Send-MailMessage -From "Fix-MailEnabled@Fix-MailEnabled" -To $emailAddress -Subject "Fix-MailEnabled script error" `
         -Body $value -SmtpServer $smtpServer
    }
}

$isTranscribing = Test-Transcribing
if (!($isTranscribing))
{
    $transcript = Start-Transcript $Global:LogFile -Append
    writelog ($transcript)
}
else
{
    writelog ("Transcript already started. Logging to the current file will continue.")
}

writelog ("Fix-MailEnabled starting.")

# Directory objects will be exported prior to deletion. This could
# potentionally create a lot of export files. By default these are
# put in the same folder as the script. If you want to put them
# elsewhere, change this path and make sure the folder exists.
$ExportPath = $ScriptPath

if ($ExchangeServer -eq "")
{
    # Choose a PF server
    $pfdbs = Get-PublicFolderDatabase
    if ($pfdbs.Length -ne $null)
    {
        $ExchangeServer = $pfdbs[0].Server.Name
    }
    else
    {
        $ExchangeServer = $pfdbs.Server.Name
    }
    
    writelog ("ExchangeServer parameter was not supplied. Using server: " + $ExchangeServer)
}

if ($DC -eq "")
{
    # Choose a DC
    $rootDSE = [ADSI]("LDAP://RootDSE")
    $DC = $rootDSE.Properties.dnsHostName
    writelog ("DC parameter was not supplied. Using DC: " + $DC)
}

$reader = new-object System.IO.StreamReader($File)

# The first line in this file must be the header line, so we can
# figure out which column is which. Folder Path is always the
# first column in an ExFolders property export.

$headerLine = $reader.ReadLine()
if (!($headerLine.StartsWith("Folder Path")))
{
    writelog "The input file doesn't seem to have the headers on the first line."
    return
}

# Figure out which column is which

$folderPathIndex = -1
$proxyIndex = -1
$proxyRequiredIndex = -1
$proxyAddressesIndex = -1
$hideFromAddressListsIndex = -1
$headers = $headerLine.Split("`t")
for ($x = 0; $x -lt $headers.Length; $x++)
{
    if ($headers[$x] -eq "Folder Path")
    {
        $folderPathIndex = $x
    }
    elseif ($headers[$x] -eq "PR_PF_PROXY: 0x671D0102")
    {
        $proxyIndex = $x
    }
    elseif ($headers[$x] -eq "PR_PF_PROXY_REQUIRED: 0x671F000B")
    {
        $proxyRequiredIndex = $x
    }
    elseif ($headers[$x] -eq "DS:proxyAddresses")
    {
        $proxyAddressesIndex = $x
    }
    elseif ($headers[$x] -eq "DS:msExchHideFromAddressLists")
    {
        $hideFromAddressListsIndex = $x
    }
}

if ($folderPathIndex -lt 0 -or `
    $proxyIndex -lt 0 -or `
    $proxyRequiredIndex -lt 0 -or `
    $proxyAddressesIndex -lt 0)
{
    writelog "Required columns were not present in the input file."
    writelog "Headers found:"
    writelog $headers
    return
}

# Loop through the lines in the file
while ($null -ne ($buffer = $reader.ReadLine()))
{
    $columns = $buffer.Split("`t")
    if ($columns.Length -lt 4)
    {
        continue
    }
    
    # Folder paths from ExFolders always start with "Public Folders" or
    # "System Folders", so trim the first 14 characters.
    $folderPath = $columns[$folderPathIndex].Substring(14)
    $guidString = $columns[$proxyIndex]
    $proxyRequired = $columns[$proxyRequiredIndex]
    $proxyAddresses = $columns[$proxyAddressesIndex]
    $hideFromAddressLists = $false
    if ($hideFromAddressListsIndex -gt -1)
    {
        if ($columns[$hideFromAddressListsIndex] -eq "True")
        {
            $hideFromAddressLists = $true
        }
    }
    
    if ($proxyRequired -ne "True" -and $proxyRequired -ne "1" -and $guidString -ne "PropertyError: NotFound" -and $guidString -ne "" -and $guidString -ne $null)
    {
        # does this objectGUID actually exist?
        $proxyObject = [ADSI]("LDAP://" + $DC + "/<GUID=" + $guidString + ">")
        if ($proxyObject.Path -eq $null)
        {
            # It's possible the object is in a different domain than the one
            # held by the specified DC, so let's try again without a specific DC.
            $proxyObject = [ADSI]("LDAP://<GUID=" + $guidString + ">")
        }
        
        if ($proxyObject.Path -ne $null)
        {
            # PR_PF_PROXY_REQUIRED is false or not set, but we have a directory object.
            # This means we need to mail-enable the folder. Ideally it would link up to
            # the existing directory object, but often, that doesn't seem to happen, and
            # we get a duplicate. So, what we're going to do here is delete the existing
            # directory object, mail-enable the folder, and then set the proxy addresses
            # from the old directory object onto the new directory object.
            
            # First, check if it's already mail-enabled. The input file could be out of
            # sync with the actual properties.
            $folder = Get-PublicFolder $folderPath -Server $ExchangeServer
            if ($folder -ne $null)
            {
                if ($folder.MailEnabled)
                {
                    writelog ("Skipping folder because it is already mail-enabled: " + $folderPath)
                    continue
                }
            }
            else
            {
                writelog ("Skipping folder because it was not found: " + $folderPath)
                continue
            }
            
            # If we got to this point, we found the PublicFolder object and it is not
            # already mail-enabled.
            writelog ("Found problem folder: " + $folderPath)
            
            # Export the directory object before we delete it, just in case
            $fileName = $ExportPath + "\" + $guidString + ".ldif"
            $ldifoutput = ldifde -d $proxyObject.Properties.distinguishedName -f ($fileName)
            writelog ("    " + $ldifoutput)
            writelog ("    Exported directory object to file: " + $fileName)

            # Save any explicit permissions
            $explicitPerms = Get-MailPublicFolder $proxyObject.Properties.distinguishedName[0] | Get-ADPermission | `
                WHERE { $_.IsInherited -eq $false -and (!($_.User.ToString().StartsWith("NT AUTHORITY"))) }

            # Save group memberships
            # We need to do this from a GC to make sure we get them all
            $memberOf = ([ADSI]("GC://" + $proxyObject.Properties.distinguishedName[0])).Properties.memberOf
            
            # Delete the current directory object
            # For some reason Parent comes back as a string in Powershell, so
            # we have to go bind to the parent.
            $parent = [ADSI]($proxyObject.Parent.Replace("LDAP://", ("LDAP://" + $DC + "/")))
            if ($parent.Path -eq $null)
            {
                $parent = [ADSI]($proxyObject.Parent)
            }
            
            if ($parent.Path -eq $null)
            {
                $proxyObject.Parent
                writelog ("Skipping folder because bind to parent container failed: " + $folderPath)
                continue
            }
            
            $parent.Children.Remove($proxyObject)
            writelog ("    Deleted old directory object.")
            
            # Mail-enable the folder
            Enable-MailPublicFolder $folderPath -Server $ExchangeServer
            writelog ("    Mail-enabled the folder.")
            
            # Disable the email address policy and set the addresses.
            # Because we just deleted the directory object a few seconds ago, it's
            # possible that that change has not replicated everywhere yet. If the
            # Exchange server still sees the object, setting the email addresses will
            # fail. The purpose of the following loop is to retry until it succeeds,
            # pausing in between. If this is constantly failing on the first try, it
            # may be helpful to increase the initial pause.
            $initialSleep = 30 # This is the initial pause. Increase it if needed.
            writelog ("    Sleeping for " + $initialSleep.ToString() + " seconds.")
            Start-Sleep $initialSleep
            
            # Filter out any addresses that aren't SMTP, and put the smtp addresses
            # into a single comma-separated string.
            $proxyAddressArray = $proxyAddresses.Split(" ")
            $proxyAddresses = ""
            foreach ($proxy in $proxyAddressArray)
            {
                if ($proxy.StartsWith("smtp:"))
                {
                    if ($proxyAddresses.Length -gt 0)
                    {
                        $proxyAddresses += ","
                    }
                    
                    $proxyAddresses += $proxy.Substring(5)
                }
                elseif ($proxy.StartsWith("SMTP:"))
                {
                    if ($proxyAddresses.Length -gt 0)
                    {
                        $proxyAddresses = $proxy.Substring(5) + "," + $proxyAddresses
                    }
                    else
                    {
                        $proxyAddresses = $proxy.Substring(5)
                    }
                }
            }

            $proxyAddresses = $proxyAddresses.Split(",")
            $retryCount = 0
            $maxRetry = 3 # The maximum number of times we'll retry
            $succeeded = $false
            while (!($succeeded))
            {
                writelog ("    Setting proxy addresses...")
                # Retrieve the new proxy object
                $newMailPublicFolder = Get-MailPublicFolder $folderPath -Server $ExchangeServer

                # Now set the properties
                $Error.Clear()
                Set-MailPublicFolder $newMailPublicFolder.Identity -EmailAddressPolicyEnabled $false -EmailAddresses $proxyAddresses `
                    -HiddenFromAddressListsEnabled $hideFromAddressLists -Server $ExchangeServer
                if ($Error[0] -eq $null)
                {
                    $succeeded = $true
                }
                else
                {
                    writelog ("    Error encountered in Set-MailPublicFolder: " + $Error[0].ToString())
                    if ($retryCount -lt $maxRetry)
                    {
                        $retryCount++
                        writelog ("    Pausing before retry. This will be retry number " `
                            + $retryCount.ToString() + ". Max retry attempts is " + $maxRetry.ToString() + ".")
                        Start-Sleep 60 # This is how long we'll pause before trying again
                    }
                    else
                    {
                        writelog ("    Max retries reached. You must manually set the properties.")
                        writelog ("    See the error log for more details.")
                        writeerror ("Failed to set proxyAddresses on folder.`r`nFolder: " + $folderPath + `
                            "`r`nProxy Addresses:`r`n" + $proxyAddresses + `
                            "`r`nGroup membership:`r`n" + $memberOf + `
                            "`r`nExplicit Permissions:`r`n" + ($explicitPerms | Select-Object User,AccessRights | out-string) + "`r`n")
                        
                        break
                    }
                }
            }

            if ($succeeded -and $explicitPerms -ne $null)
            {
                $succeeded = $true
                writelog ("    Setting explicit permissions on new directory object...")
                $newMailPublicFolder = Get-MailPublicFolder $folderPath -Server $ExchangeServer
                foreach ($permission in $explicitPerms)
                {
                    $Error.Clear()
                    $temp = Add-ADPermission $newMailPublicFolder.Identity -User $permission.User -AccessRights $permission.AccessRights
                    if ($Error[0] -ne $null)
                    {
                        $succeeded = $false
                        writelog ("    Error setting explicit permissions. You must manually set the permissions:")
                        writelog ($explicitPerms)
                        writeerror ("Failed to set explicit permissions on folder.`r`nFolder: " + $folderPath + `
                        "`r`nExplicit Permissions:`r`n" + ($explicitPerms | Select-Object User,AccessRights | out-string) + "`r`n")

                        break
                    }
                }
            }

            if ($succeeded -and $memberOf -ne $null)
            {
                writelog ("    Setting group memberships...")
                $newMailPublicFolder = Get-MailPublicFolder $folderPath -Server $ExchangeServer
                $proxy = [ADSI]("LDAP://<GUID=" + $newMailPublicFolder.Guid + ">")
                $proxyDn = $proxy.Properties.distinguishedName[0]
                $succeeded = $true
                foreach ($group in $memberOf)
                {
                    $Error.Clear()
                    $groupObject = [ADSI]("LDAP://" + $group)
                    $temp = $groupObject.Properties.member.Add($proxyDn)
                    $groupObject.CommitChanges()
                    if ($Error[0] -ne $null)
                    {
                        writelog ("    Error setting group memberships. You must add the folder to these groups:")
                        writelog ($memberOf)
                        writeerror ("Failed to set group memberships on folder.`r`nFolder: " + $folderPath + `
                        "`r`nGroup memberships:`r`n" + $memberOf + "`r`n")
                        $succeeded = $false
                        break
                    }
                }
            }
            
            writelog ("    Set the properties on the new directory object.")
            writelog ("    Done with this folder.")   
        }
        else
        {
            # This means the input file said this was a bad folder, but when we tried
            # to bind to the objectGUID from PR_PF_PROXY, we failed. Either the file
            # is out of sync with the folder settings, or something else went wrong.
            # Do we want to generate any output here?
            writelog ("Skipping folder because the objectGUID was not found: " + $folderPath)
        }

    }
    else
    {
        # If we got here, it means that according to the input file, the folder is
        # not in a state where it has a directory object but is not mail-enabled.
        # Nothing we need to do in that case. The folder is good.
    }
    
}

$reader.Close()
writelog "Done!"
if (!($isTranscribing))
{
    Stop-Transcript
}

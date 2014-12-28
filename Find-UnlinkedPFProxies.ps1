# Find-UnlinkedPFProxies.ps1
# 
# Find Exchange Server public folder directory objects that are not linked to any existing folder.

$gcRootDSE = [ADSI]"GC://RootDSE"
$gcRoot = [ADSI]("GC://" + $gcRootDSE.dnsHostName)
"Using " + $gcRoot.Path

# Find all the PF proxies
"Finding all public folder directory objects..."
$pfProxyFinder = new-object System.DirectoryServices.DirectorySearcher($gcRoot, "(objectClass=publicFolder)")
$pfProxyFinder.PageSize = 100
$allPfProxies = $pfProxyFinder.FindAll()
"Found " + $allPfProxies.Count.ToString() + " public folder directory objects."

# Make a dictionary with the GUID to DN mapping of each PF proxy
$pfProxyDictionary = new-object 'System.Collections.Generic.Dictionary[string, string]'
foreach ($pfProxy in $allPfProxies)
{
    $guid = new-object Guid(,$pfProxy.Properties["objectguid"][0])
    $pfProxyDictionary.Add($guid.ToString(), $pfProxy.Properties["distinguishedname"][0].ToString())
}

# Don't need the search results anymore, so give back that memory
$allPfProxies = $null

# Find all the mail-enabled PFs using Exchange Management Shell
"Finding all mail-enabled public folders..."
$mailEnabledPFs = Get-PublicFolder -recurse -resultsize unlimited | WHERE { $_.MailEnabled -eq $true }
"Found " + $mailEnabledPFs.Count.ToString() + " mail-enabled public folders."

# Is this Exchange 2013? If so, we can just look at MailRecipientGuid, which is faster
$pfobjectMembers = Get-PublicFolder | get-member
$hasRecipientGuid = ($pfobjectMembers | WHERE { $_.Name -eq "MailRecipientGuid" }).Length -gt 0

# Now figure out which directory objects aren't being used
"Determining which directory objects are not linked to a folder..."
foreach ($pf in $mailEnabledPFs)
{
    "Checking folder: " + $pf.ParentPath + $pf.Name
    $guidString = $null
    if ($hasRecipientGuid)
    {
        $guidString = $pf.MailRecipientGuid.ToString()
    }
    else
    {
        $guidString = ($pf | Get-MailPublicFolder).Guid.ToString()
    }

    $mailPF = ($pf | Get-MailPublicFolder)
    if ($pfProxyDictionary.ContainsKey($mailPF.Guid.ToString()))
    {
        $pfProxyDictionary.Remove($mailPF.Guid.ToString()) | Out-Null
    }
}

# The only stuff left in the dictionary at this point is unlinked directory objects
""
"There are " + $pfProxyDictionary.Count.ToString() + " public folder directory objects not linked to any folder. They are:"

foreach ($value in $pfProxyDictionary.Values)
{
    $value
}

"Done!"
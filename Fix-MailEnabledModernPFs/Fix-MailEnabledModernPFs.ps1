# Fix-MailEnabledModernPFs.ps1
# 
# Rev 53
# 
# In migration scenarios, we commonly see situations where
# mail-enabled public folders still have working directory
# objects while their MailEnabled property has been flipped
# to False. In older versions of Exchange, fixing this was
# complicated.
# 
# Exchange 2013 CU9 exposes the MailRecipientGuid property,
# which makes this much easier. This script combines the
# functionality of several older scripts, and takes advantage
# of MailRecipientGuid to correct these public folders.

Add-Type @"
    using System;
    using System.Collections.Generic;
    using System.DirectoryServices;
    using System.Linq;
    public class PfProxyCollection {
        public Dictionary<string, SearchResult> ByGuid = new Dictionary<string, SearchResult>();
        public Dictionary<string, SearchResult> ByEntryId = new Dictionary<string, SearchResult>();
        public Dictionary<string, List<SearchResult>> ByLegacyDN = new Dictionary<string, List<SearchResult>>();

        public void Add(SearchResult result)
        {
            var guid = new Guid((byte[])result.Properties["objectGUID"][0]).ToString().ToUpper();
            if (ByGuid.ContainsKey(guid))
            {
                Console.WriteLine("WARNING! Duplicate objectGUID: " + guid);
            }
            else
            {
                ByGuid.Add(guid, result);
            }

            var legDN = result.Properties["legacyExchangeDN"][0].ToString().ToUpper();
            if (ByLegacyDN.ContainsKey(legDN))
            {
                ByLegacyDN[legDN].Add(result);
            }
            else
            {
                ByLegacyDN.Add(legDN, new List<SearchResult>());
                ByLegacyDN[legDN].Add(result);
            }

            if (result.Properties.Contains("msExchPublicFolderEntryId"))
            {
                var entryId = result.Properties["msExchPublicFolderEntryId"][0].ToString();
                if (ByEntryId.ContainsKey(entryId))
                {
                    Console.WriteLine("WARNING! Duplicate msExchPublicFolderEntryId: " + entryId);
                }
                else
                {
                    ByEntryId.Add(entryId, result);
                }
            }
        }

        // If exactly one of the duplicates has msExchPublicFolderEntryId,
        // then that object is linked to the folder, and the others are
        // expendable. This method returns a list of expendable DNs and
        // removes them from the collections.
        public IEnumerable<string> FilterDuplicates()
        {
            try
            {
                Console.WriteLine("Filtering duplicate legacyExchangeDNs...");
                var dupes = ByLegacyDN.Where(p => p.Value.Count > 1).ToList();
                var total = 0;
                var expendableDistinguishedNames = new List<string>();
                foreach (var dupe in dupes)
                {
                    total += dupe.Value.Count;
                    var hasEntryId = dupe.Value.Where(p => p.Properties.Contains("msExchPublicFolderEntryId")).ToList();
                    if (hasEntryId.Count < 1 || hasEntryId.Count > 1)
                    {
                        // In this case, keep the oldest
                        var byOldest = dupe.Value.OrderBy(p => ((DateTime)p.Properties["whenCreated"][0])).ToList();
                        ByLegacyDN[dupe.Key] = new List<SearchResult>();
                        ByLegacyDN[dupe.Key].Add(byOldest[0]);
                        for (var x = 1; x < byOldest.Count; x++)
                        {
                            expendableDistinguishedNames.Add(byOldest[x].Properties["distinguishedName"][0].ToString());
                            ByGuid.Remove(new Guid((byte[])byOldest[x].Properties["objectGUID"][0]).ToString().ToUpper());
                        }
                    }
                    else
                    {
                        // Exactly one object with entry ID. The rest can be deleted.
                        foreach (var proxy in dupe.Value.Where(proxy => !proxy.Properties.Contains("msExchPublicFolderEntryId")))
                        {
                            expendableDistinguishedNames.Add(proxy.Properties["distinguishedName"][0].ToString());
                            ByGuid.Remove(new Guid((byte[])proxy.Properties["objectGUID"][0]).ToString().ToUpper());
                        }
                        
                        ByLegacyDN[dupe.Key] = new List<SearchResult>(hasEntryId);
                    }
                }
    
                Console.WriteLine("  Duplicate unique values: " + dupes.Count);
                Console.WriteLine("  Total duplicate proxies: " + total);
                Console.WriteLine("     Duplicates to delete: " + expendableDistinguishedNames.Count);
                return expendableDistinguishedNames;
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
                Console.WriteLine(exc.StackTrace);
                throw;
            }
        }

        private string GetByteStringFromBytes(byte[] bytes)
        {
            var byteString = "";
            for (var x = 0; x < bytes.Length; x+=2)
            {
                byteString += bytes[x].ToString("X2");
            }

            return byteString;
        }
    }
"@ -ReferencedAssemblies System.DirectoryServices.dll -ErrorAction SilentlyContinue

Add-Type @"
    using System;
    using System.Collections.Generic;
    using Microsoft.Exchange.Data.Storage.Management;
    public class PfCollection {
        public Dictionary<string, PublicFolder> ByMailRecipientGuid = new Dictionary<string, PublicFolder>();
        public Dictionary<string, PublicFolder> ByEntryId = new Dictionary<string, PublicFolder>();

        public void Add(PublicFolder pf)
        {
            if (!(pf.MailRecipientGuid == null || pf.MailRecipientGuid == Guid.Empty))
            {
                var guidString = pf.MailRecipientGuid.ToString().ToUpper();
                if (ByMailRecipientGuid.ContainsKey(guidString))
                {
                    Console.WriteLine("WARNING! Duplicate MailRecipientGuid: " + guidString);
                }
                else
                {
                    ByMailRecipientGuid.Add(guidString, pf);
                }
            }

            var entryId = pf.EntryId.ToString();
            if (ByEntryId.ContainsKey(entryId))
            {
                Console.WriteLine("WARNING! Duplicate EntryId: " + entryId);
                Console.WriteLine("    This Folder: " + pf.Identity.ToString());
                Console.WriteLine("    Already present: " + ByEntryId[entryId].Identity.ToString());
            }
            else
            {
                ByEntryId.Add(entryId, pf);
            }
        }
    }
"@ -ReferencedAssemblies @("C:\Program Files\Microsoft\Exchange Server\V15\bin\Microsoft.Exchange.Data.Storage.dll", `
    "C:\Program Files\Microsoft\Exchange Server\V15\bin\Microsoft.Exchange.Data.dll") -ErrorAction SilentlyContinue

function GetPartialEntryIdFromLegDN($legDn)
{
    if ($legDn.StartsWith("/CN=Mail Public Folder"))
    {
        "Can't parse this style of leg DN yet: " + $legDn
        return $null
    }
    $legDn = $legDn.TrimEnd(';')
    $legDnEnd = $legDn.Substring($legDn.Length - (19*2))
    if ($legDnEnd.Contains("-"))
    {
        $legDnEnd = $legDn.Substring($legDn.Length - (22*2) - 1).Replace("-", "")
        return $legDnEnd
    }
    else
    {
        $id = ""
        $firstPart = $legDnEnd.Substring(0, 8)
        $id += ReverseByteString $firstPart
        $secondPart = $legDnEnd.Substring(8, 8)
        $id += ReverseByteString $secondPart
        $thirdPart = $legDnEnd.Substring(16, 8)
        $id += ReverseByteString $thirdPart
        $fourthPart = $legDnEnd.Substring(24, 8)
        $id += ReverseByteString $fourthPart
        $id += "000000"
        $id += $legDnEnd.Substring(32, 6)
        return $id
    }
}

function ReverseByteString($byteString)
{
    $returnValue = ""
    for ($x = $byteString.Length - 2; $x -ge 0; $x-=2) 
    {
        $returnValue += $byteString.Substring($x, 2)
    }

    return $returnValue
}

function AddProxyToDeletionImport($dn)
{
    $script:proxiesToDelete.Add($dn)
    $domain = $dn.Substring($dn.IndexOf("DC="))
    Add-Content "$pwd\DeleteUnneededPFProxies-$domain.ldf" ("DN: " + $dn)
    Add-Content "$pwd\DeleteUnneededPFProxies-$domain.ldf" "changetype: delete"
    Add-Content "$pwd\DeleteUnneededPFProxies-$domain.ldf" ""
}

function AddToRelinkScript($thisPf, $guid, $dn)
{
    # sanity check
    if ($thisPf.MailRecipientGuid -ne $null -and $thisPf.MailRecipientGuid -ne [Guid]::Empty)
    {
        "WARNING! Blocked an attempt to link: "
        "    " + $dn
        "    " + $thisPf.Identity.ToString()
        "    That folder is already linked."
        return
    }
    #
    $entryId = $thisPf.EntryId.ToString()
    Add-Content "$pwd\ReLinkPFProxies.ps1" ("#  Proxy: " + $dn)
    Add-Content "$pwd\ReLinkPFProxies.ps1" ("# Folder: " + $thisPf.Identity.ToString())
    Add-Content "$pwd\ReLinkPFProxies.ps1" ("Set-PublicFolder `"" + $thisPf.Identity.ToString() + "`" -MailEnabled `$true -MailRecipientGuid " + $guid)
    $proxiesLinkedByScript.Add($dn)
    
    PopulateEntryIdOnProxyObject $dn $thisPf.EntryId.ToString()
}

function PopulateEntryIdOnProxyObject($dn, $entryId)
{
    $domain = $dn.Substring($dn.IndexOf("DC="))
    Add-Content "$pwd\PopulateEntryIds-$domain.ps1" ("Set-AdObject `"$dn`" -Replace @{msExchPublicFolderEntryId='$entryId'} -Server `$args[0]")
}

Set-ADServerSettings -ViewEntireForest $true

$proxiesToDelete = new-object 'System.Collections.Generic.List[string]'
$proxiesAlreadyLinked = new-object 'System.Collections.Generic.List[string]'
$proxiesLinkedByScript = new-object 'System.Collections.Generic.List[string]'

"Finding a GC..."

if ($allPfProxyResults -eq $null)
{
    # Find all the PF proxies
    $gcRootDSE = [ADSI]"GC://RootDSE"
    $gcRoot = [ADSI]("GC://" + $gcRootDSE.dnsHostName)
    "Using " + $gcRoot.Path
    "Finding all public folder directory objects..."
    $pfProxyFinder = new-object System.DirectoryServices.DirectorySearcher($gcRoot, "(objectClass=publicFolder)")
    $pfProxyFinder.PageSize = 1000
    $allPfProxyResults = $pfProxyFinder.FindAll()
    "Found " + $allPfProxyResults.Count.ToString() + " public folder directory objects."
}
else
{
    "`$allPfProxyResults is already populated. Using existing data."
}

# Store these in the collection
$pfProxies = new-object PfProxyCollection
foreach ($pfProxy in $allPfProxyResults)
{
    $pfProxies.Add($pfProxy)
}

$expendables = $pfProxies.FilterDuplicates()
foreach ($dn in $expendables)
{
    AddProxyToDeletionImport $dn
}

# Retrieve all PFs using Exchange Management Shell
# You can skip this step by putting all the PFs in $allPFs before running the script
# This saves time if you need to run this more than once
if ($allPFs -eq $null)
{
    "Finding all public folders..."
    $allPFs = Get-PublicFolder -recurse -resultsize unlimited
}
else
{
    "`$allPFs is already populated. Using existing data."
}

$mailEnabledPFs = $allPFs | WHERE { $_.MailEnabled -eq $true -and $_.MailRecipientGuid -ne [Guid]::Empty }
$nonMailEnabledPFs = $allPFs | WHERE { $_.MailEnabled -eq $false -or $_.MailRecipientGuid -eq [Guid]::Empty }
if (($allPFs[0].PSObject.Properties | WHERE { $_.Name -like "MailRecipientGuid" }).Length -lt 1)
{
    "MailRecipientGuid is not present. This script must be run on Exchange 2013 or later."
    return
}

"Found " + $mailEnabledPFs.Length + " mail-enabled public folders."
"Found " + $nonMailEnabledPFs.Length + " mail-disabled public folders."

$pfs = new-object PfCollection
foreach ($pf in $allPfs)
{
    $pfs.Add($pf)
}

# Now do the real work

"Figuring out which proxies are not linked..."
$unlinkedPFProxies = new-object 'System.Collections.Generic.Dictionary[string, object]'
foreach ($guid in $pfProxies.ByGuid.Keys)
{
    if (!$pfs.ByMailRecipientGuid.ContainsKey($guid))
    {
        $unlinkedPFProxies.Add($guid, $pfProxies.ByGuid[$guid])
    }
    else
    {
        $proxiesAlreadyLinked.Add($pfProxies.ByGuid[$guid].Properties["distinguishedName"][0].ToString())
        $pf = $pfs.ByMailRecipientGuid[$guid]
        $pfs.ByMailRecipientGuid.Remove($guid) | Out-Null
        $pfs.ByEntryId.Remove($pf.EntryId.ToString()) | Out-Null
    }
}

$proxiesNotMatched = new-object 'System.Collections.Generic.List[string]'

"Found " + $unlinkedPFProxies.Count.ToString() + " unlinked PF proxies."
if ($unlinkedPFProxies.Count -gt 0) 
{
    "Creating a script to link these to their folders..."
    Set-Content "$pwd\ReLinkPFProxies.ps1" ""
    $pfsAlreadyInOutput = new-object 'System.Collections.Generic.List[string]'
    foreach ($proxy in $unlinkedPFProxies.Values)
    {
        "Evaluating proxy: " + $proxy.Properties["distinguishedName"][0].ToString()
        $entryId = $null
        if ($proxy.Properties.Contains("msExchPublicFolderEntryId"))
        {
            "    Has msExchPublicFolderEntryId"
            $entryId = $proxy.Properties["msExchPublicFolderEntryId"][0].ToString()
        }
        $partialEntryId = $null
        if ($entryId -eq $null)
        {
            "    Getting partial entry ID"
            $partialEntryId = GetPartialEntryIdFromLegDN $proxy.Properties["legacyExchangeDN"][0].ToString()
            $partialEntryId += "0000"
        }
        
        $guid = (new-object Guid(,$proxy.Properties["objectguid"][0])).ToString().ToUpper()
    
        # If we have an entry ID, just use that
        if ($entryId -ne $null)
        {
            "    Getting public folder based on full entry ID"
            $thisPf = Get-PublicFolder $entryId
            # Is it already linked?
            if ($thisPf.MailRecipientGuid -ne $null -and $thisPf.MailRecipientGuid -ne [Guid]::Empty)
            {
                if ($thisPf.MailRecipientGuid.ToString().ToUpper() -ne $guid)
                {
                    # Then this is a duplicate
                    AddProxyToDeletionImport $proxy.Properties["distinguishedName"][0].ToString()
                }
            }
            elseif ($pfsAlreadyInOutput.Contains($thisPf.Identity.ToString()))
            {
                "WARNING! Multiple pf proxies have this entryId. This was supposed to be filtered out already."
                "    " + $entryId
                "    " + $thisPf.Identity.ToString()
            }
            else
            {
                "    Linking to this object."
                AddToRelinkScript $thisPf $guid $proxy.Properties["distinguishedName"][0].ToString()
                $pfsAlreadyInOutput.Add($thisPf.Identity.ToString())
            }
        }
        # Otherwise use the one we built from leg DN
        elseif ($partialEntryId -ne $null)
        {
            "    Using partial entry ID"
            $matches = @()
            foreach ($key in $pfs.ByEntryId.Keys)
            {
                if ($key.EndsWith($partialEntryId))
                {
                    $matches += $key
                }
            }
    
            "    matches: " + $matches.Length
            # If we still have no matches, match on the last 10 characters
            if ($matches.Length -eq 0)
            {
                "    Matching by last 10"
                foreach ($key in $pfs.ByEntryId.Keys)
                {
                    if ($key.EndsWith($partialEntryId.Substring($partialEntryId.Length - 10)))
                    {
                        $matches += $key
                    }
                }
            }
            
            "    matches: " + $matches.Length
            # If we have more than one, try to narrow it down using name
            if ($matches.Length -gt 1)
            {
                "    narrowing by display name"
                $oldMatches = $matches
                $matches = @()
                foreach ($oldMatch in $oldMatches)
                {
                     if ($proxy.Properties["displayName"] -eq $pfs.ByEntryId[$oldMatch].Name)
                     {
                         $matches += $oldMatch
                     } 
                }
            }
            # If we STILL have none, try to match by name alone
            elseif ($matches.Length -eq 0)
            {
                "    matching by name alone"
                foreach ($key in $pfs.ByEntryId.Keys)
                {
                    if ($proxy.Properties["displayName"] -eq $pfs.ByEntryId[$key].Name)
                    {
                        if ($pfs.ByEntryId[$key].MailRecipientGuid -ne $null -and $pfs.ByEntryId[$key].MailRecipientGuid -ne [Guid]::Empty)
                        {
                            $matches += $key
                        }
                    }
                }
            }
            
            # If we found no match at all, this proxy can be deleted
            if ($matches.Length -eq 0)
            {
                "    no match at all"
                AddProxyToDeletionImport $proxy.Properties["distinguishedName"][0].ToString()
            }
            elseif ($matches.Length -eq 1)
            {
                "    exactly one match"
                $thisPf = $pfs.ByEntryId[$matches[0]]
                if ($thisPf.MailRecipientGuid -ne $null -and $thisPf.MailRecipientGuid -ne [Guid]::Empty)
                {
                    # This PF is already linked to something else. Since that was the only match,
                    # this proxy can be deleted.
                    AddProxyToDeletionImport $proxy.Properties["distinguishedName"][0].ToString()
                }
                elseif ($pfsAlreadyInOutput.Contains($thisPf.Identity.ToString()))
                {
                    "WARNING! Multiple pf proxies match this folder. This was supposed to be filtered out already."
                    "    " + $matches[0]
                    "    " + $thisPf.Identity.ToString()
                }
                else
                {
                    AddToRelinkScript $thisPf $guid $proxy.Properties["distinguishedName"][0].ToString()
                    $pfsAlreadyInOutput.Add($thisPf.Identity.ToString())
                }
            }
            elseif ($matches.Length -gt 1)
            {
                "WARNING: Found more than one match for:"
                "    " + $proxy.Properties["distinguishedName"][0].ToString()
                foreach ($match in $matches)
                {
                    "    " + $pfs.ByEntryId[$match].Identity.ToString()
                }
                
                $proxiesNotMatched.Add($proxy.Properties["distinguishedName"][0].ToString())
            }
        }
        else 
        {
            "WARNING: Could not produce an entryId or partial entryId for: "
            "    " + $proxy.Properties["distinguishedName"][0].ToString()
        }
    }
}

""
"           Proxies total: " + $allPfProxyResults.Count.ToString()
"  Proxies already linked: " + $proxiesAlreadyLinked.Count.ToString()
"Proxies linked by script: " + $proxiesLinkedByScript.Count.ToString()
"     Proxies not matched: " + $proxiesNotMatched.Count.ToString()
"       Proxies to delete: " + $proxiesToDelete.Count.ToString()
""
"Saving these results for sanity checking..."

Set-Content "$pwd\ProxiesTotal.txt" ""
foreach ($proxy in $allPfProxyResults) { Add-Content "$pwd\ProxiesTotal.txt" $proxy.Properties["distinguishedName"][0].ToString() }
Set-Content "$pwd\ProxiesAlreadyLinked.txt" ""
foreach ($dn in $proxiesAlreadyLinked) { Add-Content "$pwd\ProxiesAlreadyLinked.txt" $dn }
Set-Content "$pwd\ProxiesLinkedByScript.txt" ""
foreach ($dn in $proxiesLinkedByScript) { Add-Content "$pwd\ProxiesLinkedByScript.txt" $dn }
Set-Content "$pwd\ProxiesNotMatched.txt" ""
foreach ($dn in $proxiesNotMatched) { Add-Content "$pwd\ProxiesNotMatched.txt" $dn }
Set-Content "$pwd\ProxiesToDelete.txt" ""
foreach ($dn in $proxiesToDelete) { Add-Content "$pwd\ProxiesToDelete.txt" $dn }
    
"Done!"
"This script generates the following files:"
""
"$pwd\ReLinkPFProxies.ps1           Commands to link proxies to their folders."
"$pwd\DeleteUnneededPFProxies.ldf   Can be imported to delete unneeded proxies."
"$pwd\ProxiesTotal.txt              All public folder proxies that were found."
"$pwd\ProxiesAlreadyLinked.txt      The proxies that were already linked to folders."
"$pwd\ProxiesLinkedByScript.txt     The proxies that will be linked by running the .ps1."
"$pwd\ProxiesNotMatched.txt         The proxies that couldn't be matched due to ambiguous name."
"$pwd\ProxiesToDelete.txt           The proxies that would be deleted by the ldf."
""
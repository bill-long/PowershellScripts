# Fix-DelegatedSetup.ps1
#
# In Exchange 2013, delegated setup fails if legacy admin groups still exist. However,
# it's recommended that these admin groups not be deleted. To make delegated setup work,
# you can temporarily place an explicit deny so that the Delegated Setup group
# cannot see them. This script automates that process.
# 
# This script takes no parameters. The syntax is simply:
# .\Fix-DelegatedSetup.ps1
# 
# The script will add the explicit deny if it is not present, and it will remove it
# if it is present. This means you simply run the script to add the deny, perform
# the delegated setup, and then run the script again to remove the deny.

$delegatedSetupRoleGroup = Get-RoleGroup "Delegated Setup"
if ($delegatedSetupRoleGroup -eq $null)
{
    "Could not get Delegated Setup role group."
    return
}

$rootDSE = [ADSI]("LDAP://RootDSE")
if ($rootDSE.configurationNamingContext -eq $null)
{
    "Could not read RootDSE."
    return
}

$adminGroupFinder = new-object System.DirectoryServices.DirectorySearcher
$adminGroupFinder.SearchRoot = [ADSI]("LDAP://" + $rootDSE.configurationNamingContext.ToString())
$adminGroupFinder.Filter = "(&(objectClass=msExchAdminGroup)(!(cn=Exchange Administrative Group (FYDIBOHF23SPDLT))))"
$adminGroupFinder.SearchScope = "Subtree"
$adminGroupResults = $adminGroupFinder.FindAll()
if ($adminGroupResults.Count -lt 1)
{
    "No legacy admin groups were found."
    return
}

# Check to see if we already have the explicit perms set

$foundExplicitPerms = $false
foreach ($result in $adminGroupResults)
{
    $explicitPerms = Get-ADPermission $result.Properties["distinguishedname"][0].ToString() | `
        WHERE { $_.IsInherited -eq $false -and $_.Deny -eq $true -and $_.User -like "*\Delegated Setup" }

    if ($explicitPerms.Length -gt 0)
    {
        $foundExplicitPerms = $true
        break
    }
}

if ($foundExplicitPerms)
{
    # We already have explicit perms on at least some legacy admin groups, so
    # remove them from all admin groups where they exist
    "Removing explicit deny for Delegated Setup..."
    foreach ($result in $adminGroupResults)
    {
        $explicitPerms = Get-ADPermission $result.Properties["distinguishedname"][0].ToString() | `
            WHERE { $_.IsInherited -eq $false -and $_.Deny -eq $true -and $_.User -like "*\Delegated Setup" }
        
        if ($explicitPerms.Length -gt 0)
        {
            foreach ($perm in $explicitPerms)
            {
                Remove-ADPermission -Instance $perm -Confirm:$false
            }
        }
    }
}
else
{
    # We don't have explicit perms on any legacy admin groups, so add them to
    # all legacy admin groups
    "Adding explicit deny for Delegated Setup..."
    foreach ($result in $adminGroupResults)
    {
        Add-ADPermission $result.Properties["distinguishedname"][0].ToString() `
            -User $delegatedSetupRoleGroup.DistinguishedName -Deny -AccessRights GenericAll | out-null
    }
}
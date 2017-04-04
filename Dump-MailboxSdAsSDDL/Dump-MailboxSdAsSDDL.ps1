param($alias)

$searcher = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().FindGlobalCatalog().GetDirectorySearcher()
$searcher.Filter = "(mailnickname=$alias)"
$user = $searcher.FindOne()
$mbxSd = $user.Properties["msExchMailboxSecurityDescriptor"][0]
$sd = New-Object System.Security.AccessControl.RawSecurityDescriptor([byte[]]$mbxSd, 0)
$sd.GetSddlForm("All")

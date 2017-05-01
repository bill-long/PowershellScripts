# Reset-MailboxSD
# 
# This version is intended for Exchange 2013 and 2016, where we have to
# update the value in AD.
# 
# Usage:
# 
# To do one user:
# 
# .\Reset-MailboxSD.ps1 "CN=SomeUser,OU=Wherever,DC=contoso,DC=com"
# 
# To do all mailboxes that start with "Foo":
# 
# Get-Mailbox Foo* | % { .\Reset-MailboxSD.ps1 $_.DistinguishedName } 

param($dn)

$securityDescriptor = New-Object -TypeName System.Security.AccessControl.RawSecurityDescriptor -ArgumentList "O:PSG:PSD:(A;CI;CCRC;;;PS)"
$user = [ADSI]("LDAP://" + $dn)
$user.Properties["msExchMailboxSecurityDescriptor"].Clear()
[byte[]]$mbxSdBytes = [System.Array]::CreateInstance([System.Byte], $securityDescriptor.BinaryLength)
$securityDescriptor.GetBinaryForm($mbxSdBytes, 0)
$user.Properties["msExchMailboxSecurityDescriptor"].Add($mbxSdBytes)
$user.CommitChanges()

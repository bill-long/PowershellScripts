param([string]$sidString)

$authzInitializeResourceManagerSig = @'
[DllImport("authz.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern bool AuthzInitializeResourceManager(int flags, IntPtr pfnAccessCheck, IntPtr pfnComputeDynamicGroups,
    IntPtr pfnFreeDynamicGroups, string szResourceManagerName, out IntPtr phAuthzResourceManager);
'@

$authzInitializeContextFromSidSig = @'
[DllImport("authz.dll", EntryPoint = "AuthzInitializeContextFromSid", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
public static extern bool AuthzInitializeContextFromSid(int flags, byte[] UserSid, IntPtr hAuthzResourceManager, IntPtr pExpirationTime,
    int id, IntPtr DynamicGroupArgs, out IntPtr pAuthzClientContext);
'@

$authzFreeContextSig = @'
[DllImport("authz.dll", EntryPoint = "AuthzFreeContext", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
public static extern bool AuthzFreeContext(IntPtr hAuthzCLientContext);
'@

$authzFreeResourceManagerSig = @'
[DllImport("authz.dll", EntryPoint = "AuthzFreeResourceManager", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
public static extern bool AuthzFreeResourceManager(IntPtr hAuthzResourceManager);
'@

$getLastErrorSig = @'
[DllImport("Kernel32.dll")]
public static extern uint GetLastError();
'@

Add-Type -MemberDefinition $authzInitializeResourceManagerSig -Name Api1 -Namespace AuthZ
Add-Type -MemberDefinition $authzInitializeContextFromSidSig -Name Api2 -Namespace AuthZ
Add-Type -MemberDefinition $authzFreeContextSig -Name Api3 -Namespace AuthZ
Add-Type -MemberDefinition $authzFreeResourceManagerSig -Name Api4 -Namespace AuthZ
Add-Type -MemberDefinition $getLastErrorSig -Name Api -Namespace Win32

$sid = New-Object System.Security.Principal.SecurityIdentifier($sidString)
[byte[]]$sidBytes = [System.Array]::CreateInstance([System.Byte], $sid.BinaryLength)
$sid.GetBinaryForm($sidBytes, 0)

$hClientContext = [IntPtr]::Zero
$hResourceManager = [IntPtr]::Zero

$success = [AuthZ.Api1]::AuthzInitializeResourceManager(1, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, "", [ref]$hResourceManager)
if (!($success))
{
    $error = [Win32.Api]::GetLastError()
    "Error $error"
    return
}

#$unused = New-Object LUID()
$success = [AuthZ.Api2]::AuthzInitializeContextFromSid(0, $sidBytes, $hResourceManager, [IntPtr]::Zero, 0, [IntPtr]::Zero, [ref]$hClientContext)
if (!($success))
{
    $error = [Win32.Api]::GetLastError()
    "Error $error"
    return
}

"Succeeded."

$success = [AuthZ.Api3]::AuthzFreeContext($hClientContext)
$success = [AuthZ.Api4]::AuthzFreeResourceManager($hResourceManager)

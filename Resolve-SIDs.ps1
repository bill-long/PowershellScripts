# Resolve-SIDs.ps1
# 
# Resolve a bunch of SIDs from a text file, in a loop.

param($file)

$sids = New-Object 'System.Collections.Generic.List[string]'
$reader = New-Object System.IO.StreamReader($file)
while ($null -ne ($buffer = $reader.ReadLine()))
{
    if ($buffer.StartsWith("S-1"))
    {
        $sidString = $buffer
        if ($sidString.Contains(" "))
        {
            $sidString = $sidString.Substring(0, $buffer.IndexOf(" "))
        }
        
        $sids.Add($sidString)
    }
}

$reader.Close()

"Read " + $sids.Count + " SIDs from file."
""

if ($sids.Count -lt 1)
{
    return
}

$stopwatch = New-Object System.Diagnostics.Stopwatch

while ($true)
{
    foreach ($sidString in $sids)
    {
        $stopwatch.Reset()
        $sid = New-Object System.Security.Principal.SecurityIdentifier($sidString)
        $stopwatch.Start()
        $result = $sid.Translate([System.Security.Principal.NTAccount])
        $stopwatch.Stop()
        if ($result -ne $null)
        {
            $stopwatch.ElapsedMilliseconds.ToString() + " msec to resolve " + $sidString + " to " + $result.Value
        }
        else
        {
            "Failed to resolve " + $sidString + " after " + $stopwatch.ElapsedMilliseconds.ToString() + " msec"
        }
    }
}
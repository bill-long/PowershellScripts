# Convert-LdifObjectsForImport
#
# The purpose of this script is to read an ldifde export
# from AD and produce a new ldifde file that can be used
# for importing objects from the original
# file into a totally different domain.

param($inputFile, $outputFile)

$reader = new-object System.IO.StreamReader($inputFile)
$writer = new-object System.IO.StreamWriter($outputFile)

if ($reader -eq $null)
{
	"Cannot read the specified input file: " + $inputFile
	return
}

if ($writer -eq $null)
{
	"Cannot write to the specified output file: " + $outputFile
	return
}

$propertiesToExclude = new-object 'System.Collections.Generic.List[string]'
if ($propertiesToExclude -eq $null)
{
	"This script requires Powershell 2.0."
	return
}

$propertiesToExclude.AddRange([string[]]@(
	"objectGUID",
	"dSCorePropagationData",
	"whenChanged",
	"whenCreated",
	"uSNCreated",
	"uSNChanged",
	"cn",
	"msExchResponsibleMTAServer",
	"homeMTA",
	"homeMDB")
	)

$outputBuffer = new-object 'System.Collections.Generic.List[string]'
$isPublicFolderObject = $false
$keepThisProperty = $false

while ($null -ne ($buffer = $reader.ReadLine()))
{
	if ($buffer -eq "")
	{
		foreach ($line in $outputBuffer)
		{
			$writer.WriteLine($line)
		}

		$writer.WriteLine("")
		$outputBuffer.Clear()
	}
	elseif ($buffer.StartsWith(" "))
	{
		if ($keepThisProperty -eq $true)
		{
			$outputBuffer.Add($buffer)
		}
	}
	else
	{
		$propertyName = $buffer.Substring(0, $buffer.IndexOf(":"))
		if ((!($propertiesToExclude.Contains($propertyName))) -and (!($propertyName.EndsWith("BL"))))
		{
			$keepThisProperty = $true
			$outputBuffer.Add($buffer)
		}
		else
		{
			$keepThisProperty = $false
		}
	}
}

$reader.Close()
$writer.Close()
"Done!"
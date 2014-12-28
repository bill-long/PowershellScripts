# Create-LotsOfPFs.ps1
#
# Creates 10,000 public folders

function CreateFoldersRecursive($parentFolder)
{
	$currentDepth++
	for ($x = 1; $x -le $foldersPerLevel; $x++)
	{
		$thisFolder = New-PublicFolder ("Folder " + $x.ToString()) -Path $parentFolder.Identity.ToString()
		$thisFolder
		if ($currentDepth -lt $desiredDepth)
		{
			CreateFoldersRecursive $thisFolder
		}
	}
}

$topFolder = New-PublicFolder "Lots Of Folders Here"
$desiredDepth = 4
$foldersPerLevel = 10
$currentDepth = 0
CreateFoldersRecursive $topFolder
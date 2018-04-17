# This reports the number of pages taken up by various tables based on
# an eseutil /ms /v dump. It will also work with /ms /f#all /csv /v.
#
# This script will only work with Exchange 2010 mailbox database space
# dumps.

param([string]$inputFilePath)
$fileReader = new-object System.IO.StreamReader($inputFilePath)
$foundHeaderLine = $false
while ($null -ne ($buffer = $fileReader.ReadLine()))
{
    if ($buffer.StartsWith("Name "))
    {
        $foundHeaderLine = $true
        # found the header line
        # is this a csv or not?
        $isCSV = ($buffer.IndexOf(",") -gt 0)
        if ($isCSV)
        {
            $headerSplit = $buffer.Split(@(","))
            for ($x = 0; $x -lt $headerSplit.Length; $x++)
            {
                if ($headerSplit[$x].Trim() -eq "Owned(MB)")
                {
                    $ownedColumnIndex = $x
                }
                elseif ($headerSplit[$x].Trim() -eq "Avail(MB)")
                {
                    $availColumnIndex = $x
                }
            }
            
            break
        }
        else
        {
            # now find the Owned header and figure out where that column starts and ends
            $typeLabelIndex = $buffer.IndexOf("Type")
            $typeLabelEnd = $typeLabelIndex + 4
            $ownedLabelIndex = $buffer.IndexOf("Owned(MB)")
            $ownedColumnEnd = $ownedLabelIndex + 9
            $ownedColumnStart = $ownedColumnEnd - 12

            $ofTableLabelIndex = $buffer.IndexOf("%OfTable")
            $ofTableLabelEnd = $ofTableLabelIndex + 8
            $availableColumnIndex = $buffer.IndexOf("Avail(MB)")
            $availableColumnEnd = $availableColumnIndex + 9
            $availableColumnStart = $availableColumnEnd - 12

            break
        }
    }
}

if (!($foundHeaderLine))
{
    "Couldn't find the header line in the space dump."
    return
}

$ownedColumnLength = $ownedColumnEnd - $ownedColumnStart
$availableColumnLength = $availableColumnEnd - $availableColumnStart

# Skip 3 lines to get to the start of the tables
for ($x = 0; $x -lt 3; $x++)
{
	$blah = $fileReader.ReadLine()
}

if ($isCSV)
{
	$blah = $fileReader.ReadLine()
}

$bodyTableSizes = new-object 'System.Collections.Generic.Dictionary[string, double]'
$bodyTableFree = new-object 'System.Collections.Generic.Dictionary[string, double]'
$msgViewTablesPerMailbox = new-object 'System.Collections.Generic.Dictionary[string, double]'

[double]$spaceOwnedByBodyTables = 0
[double]$freeSpaceByBodyTables = 0
[double]$numberOfBodyTables = 0
[double]$spaceOwnedByCategTables = 0
[double]$freeSpaceByCategTables = 0
[double]$numberOfCategTables = 0
[double]$spaceOwnedByDVUTables = 0
[double]$freeSpaceByDVUTables = 0
[double]$numberOfDVUTables = 0
[double]$spaceOwnedByFoldersTables = 0
[double]$freeSpaceByFoldersTables = 0
[double]$numberOfFoldersTables = 0
[double]$spaceOwnedByMsgHeaderTables = 0
[double]$freeSpaceByMsgHeaderTables = 0
[double]$numberofMsgHeaderTables = 0
[double]$spaceOwnedByMsgViewTables = 0
[double]$freeSpaceByMsgViewTables = 0
[double]$numberOfMsgViewTables = 0
[double]$spaceOwnedByCIPropStoreTables = 0
[double]$freeSpaceByCIPropStoreTables = 0
[double]$numberOfCIPropStoreTables = 0
[double]$spaceOwnedByOtherTables = 0
[double]$freeSpaceByOtherTables = 0
[double]$numberOfOtherTables = 0

while ($null -ne ($buffer = $fileReader.ReadLine()))
{
    if (!($buffer.StartsWith("    ")) -and $buffer -ne "")
    {
        if ($buffer.StartsWith("-----"))
        {
            break;
        }

        if ($isCSV)
        {
            $bufferSplit = $buffer.Split(@(","))
            $thisOwnedSpace = [System.Double]::Parse($bufferSplit[$ownedColumnIndex])
            $thisAvailSpace = [System.Double]::Parse($bufferSplit[$availColumnIndex])
        }
        else
        {
            $thisOwnedSpace = [System.Double]::Parse($buffer.Substring($ownedColumnStart, $ownedColumnLength))
            $thisAvailSpace = [System.Double]::Parse($buffer.Substring($availableColumnStart, $availableColumnLength))
        }
        
        if ($buffer.StartsWith("  Body-"))
        {
            $numberOfBodyTables++
            $spaceOwnedByBodyTables += $thisOwnedSpace
            if ($isCSV)
            {
                $bodyTableName = $bufferSplit[0].Trim()
            }
            else
            {
                $bodyTableName = $buffer.Substring(0, $typeLabelIndex).Trim()
            }
            $bodyTableSizes.Add($bodyTableName, $thisOwnedSpace)

            $freeSpaceByBodyTables += $thisAvailSpace
            $bodyTableFree.Add($bodyTableName, $thisAvailSpace)
        }
        elseif ($buffer.StartsWith("  Categ-"))
        {
            $numberOfCategTables++
            $spaceOwnedByCategTables += $thisOwnedSpace
            $freeSpaceByCategTables += $thisAvailSpace
        }
        elseif ($buffer.StartsWith("  DVU-"))
        {
            $numberOfDVUTables++
            $spaceOwnedByDVUTables += $thisOwnedSpace
            $freeSpaceByDVUTables += $thisAvailSpace
        }
        elseif ($buffer.StartsWith("  Folders-"))
        {
            $numberOfFoldersTables++
            $spaceOwnedByFoldersTables += $thisOwnedSpace
            $freeSpaceByFoldersTables += $thisAvailSpace
        }
        elseif ($buffer.StartsWith("  MsgHeader-"))
        {
            $numberofMsgHeaderTables++
            $spaceOwnedByMsgHeaderTables += $thisOwnedSpace
            $freeSpaceByMsgHeaderTables += $thisAvailSpace
        }
        elseif ($buffer.StartsWith("  MsgView-"))
        {
            $numberOfMsgViewTables++
            $spaceOwnedByMsgViewTables += $thisOwnedSpace
            $freeSpaceByMsgViewTables += $thisAvailSpace

            $firstSpaceIndex = $buffer.IndexOf(" ", 10)
            $mailboxFID = $buffer.Substring(10, $firstSpaceIndex - 10)
            $firstMinusIndex = $mailboxFID.IndexOf("-")
            $secondMinusIndex = $mailboxFID.IndexOf("-", $firstMinusIndex + 1)
            if ($secondMinusIndex -gt 0)
            {
                $mailboxFID = $mailboxFID.Substring(0, $secondMinusIndex)
            }

            $tableCount = $null
            if ($msgViewTablesPerMailbox.TryGetValue($mailboxFID, [ref]$tableCount))
            {
                $msgViewTablesPerMailbox[$mailboxFID] = $tableCount + 1
            }
            else
            {
                $msgViewTablesPerMailbox.Add($mailboxFID, 1)
            }
        }
        else
        {
            $numberOfOtherTables++
            $spaceOwnedByOtherTables += $thisOwnedSpace
            $freeSpaceByOtherTables += $thisAvailSpace
        }
    }
}

$totalSpace = $spaceOwnedByBodyTables + $spaceOwnedByCategTables + $spaceOwnedByDVUTables + $spaceOwnedByFoldersTables + $spaceOwnedByMsgHeaderTables + $spaceOwnedByMsgViewTables + $spaceOwnedByOtherTables
$totalFree = $freeSpaceByBodyTables + $freeSpaceByCategTables + $freeSpaceByDVUTables + $freeSpaceByFoldersTables + $freeSpaceByMsgHeaderTables + $freeSpaceByMsgViewTables + $freeSpaceByOtherTables

"     Space owned by Body tables: " + $spaceOwnedByBodyTables.ToString("F3")
"         % owned by Body tables: " + (($spaceOwnedByBodyTables / $totalSpace) * 100).ToString("F2")
"      Free space in Body tables: " + $freeSpaceByBodyTables.ToString("F3")
"          Number of Body tables: " + $numberOfBodyTables.ToString()
"    Space owned by Categ tables: " + $spaceOwnedByCategTables.ToString("F3")
"        % owned by Categ tables: " + (($spaceOwnedByCategTables / $totalSpace) * 100).ToString("F2")
"     Free space in Categ tables: " + $freeSpaceByCategTables.ToString("F3")
"         Number of Categ tables: " + $numberOfCategTables.ToString()
"      Space owned by DVU tables: " + $spaceOwnedByDVUTables.ToString("F3")
"          % owned by DVU tables: " + (($spaceOwnedByDVUTables / $totalSpace) * 100).ToString("F2")
"       Free space in DVU tables: " + $freeSpaceByDVUTables.ToString("F3")
"           Number of DVU tables: " + $numberOfDVUTables.ToString()
"  Space owned by Folders tables: " + $spaceOwnedByFoldersTables.ToString("F3")
"      % owned by Folders tables: " + (($spaceOwnedByFoldersTables / $totalSpace) * 100).ToString("F2")
"   Free space in Folders tables: " + $freeSpaceByFoldersTables.ToString("F3")
"       Number of Folders tables: " + $numberOfFoldersTables.ToString()
"Space owned by MsgHeader tables: " + $spaceOwnedByMsgHeaderTables.ToString("F3")
"    % owned by MsgHeader tables: " + (($spaceOwnedByMsgHeaderTables / $totalSpace) * 100).ToString("F2")
" Free space in MsgHeader tables: " + $freeSpaceByMsgHeaderTables.ToString("F3")
"     Number of MsgHeader tables: " + $numberOfMsgHeaderTables.ToString()
"  Space owned by MsgView tables: " + $spaceOwnedByMsgViewTables.ToString("F3")
"      % owned by MsgView tables: " + (($spaceOwnedByMsgViewTables / $totalSpace) * 100).ToString("F2")
"   Free space in MsgView tables: " + $freeSpaceByMsgViewTables.ToString("F3")
"       Number of MsgView tables: " + $numberOfMsgViewTables.ToString()
"    Space owned by other tables: " + $spaceOwnedByOtherTables.ToString("F3")
"        % owned by other tables: " + (($spaceOwnedByOtherTables / $totalSpace) * 100).ToString("F2")
"     Free space in other tables: " + $freeSpaceByOtherTables.ToString("F3")
"         Number of other tables: " + $numberOfOtherTables.ToString()
""
"Total space owned by all tables: " + $totalSpace.ToString("F3")
" Total space free in all tables: " + $totalFree.ToString("F3")
""
"Largest body tables:"
$top10BodyTables = $bodyTableSizes.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10
foreach ($kv in $top10BodyTables)
{
    ("  " + $kv.Key + " Owned: " + $kv.Value.ToString("F3") + " Free: " + $bodyTableFree[$kv.Key].ToString("F3"))
}

"Body tables with most free space:"
$top10FreeBodyTables = $bodyTableFree.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10
foreach ($kv in $top10FreeBodyTables)
{
    ("  " + $kv.Key + " Owned: " + $bodyTableSizes[$kv.Key].ToString("F3") + " Free: " + $kv.Value.ToString("F3"))
}

"Mailboxes with the most MsgView tables:"
$top10MsgView = $msgViewTablesPerMailbox.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 10
foreach ($kv in $top10MsgView)
{
    ("  " + $kv.Key + " MsgView Count: " + $kv.Value.ToString())
}

$fileReader.Close()

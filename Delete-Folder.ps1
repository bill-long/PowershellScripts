# Delete-Folder.ps1
# 
# A Powershell script to delete folders that contain files and directories exceeding the MAX_PATH limit on Windows.
# Requires Microsoft.Experimental.IO.dll, which can be downloaded from Codeplex:
# http://bcl.codeplex.com/wikipage?title=Long%20Path

param($folder)

$ExperimentalIOBinary = 'C:\Users\administrator\Desktop\Microsoft.Experimental.IO.dll'
[System.Reflection.Assembly]::LoadFile($ExperimentalIOBinary)

function DeleteAllFilesRecursive($path)
{
    "Getting folders in folder: " + $path
    $subfolders = [Microsoft.Experimental.IO.LongPathDirectory]::EnumerateDirectories($path)
    foreach ($subfolder in $subfolders)
    {
        "Recursing folder: " + $subfolder
        DeleteAllFilesRecursive($subfolder)
        "Deleting folder: " + $subfolder
        [Microsoft.Experimental.IO.LongPathDirectory]::Delete($subfolder)
    }

    $files = [Microsoft.Experimental.IO.LongPathDirectory]::EnumerateFiles($path)
    foreach ($file in $files)
    {
        "Deleting file: " + $file
        [Microsoft.Experimental.IO.LongPathFile]::Delete($file)
    }
}

DeleteAllFilesRecursive $folder
"Deleting folder: " + $folder
[Microsoft.Experimental.IO.LongPathDirectory]::Delete($folder)
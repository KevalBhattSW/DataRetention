clear

#Function to loop through a collection of files and check their age 
function List-FileAgeProperties {
    param (
        [System.Collections.ArrayList]$Files,
        [String]$filename,
        [String]$progressFile,
        [string]$driveLetter = $null,
        [string]$mappedPath = $null
    )

    if(($driveLetter -and -not $mappedPath) -or (-not $driveLetter -and $mappedPath)) {
        Write-Error "Both driveLetter mappedPath should be provided or both be blank"
        return $null
    }


    # Create new list
    $filesScanned = New-Object System.Collections.ArrayList

    # Test if a progress file exists. If not, quit.
    if (!(Test-Path -Path $progressFile -PathType Leaf)) {
        Write-Error "File $progressFile does not exist"
        return $null
    }

    # Create progress file entry
    $currentTime = Get-Date
    $currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "Catalogue for $filename started at $($currentTimeF)"
    Add-Content -Path $progressFile -Value $logEntry

    # Create file list header entry
    $listEntry = @("Name","Containing Path","Size","Last Modified","Last Accessed","Creation Date","Extension","Last Save Date","Date Checked") -Join [char]9
    $filesScanned.Add($listEntry)

    # Loop through each file in collection parameter
    foreach ($objFile in $Files) {
        if($objfile.Substring(0,1) -ne "~") {
            $item = Get-Item $objFile

            # Get file metadata
            $currentFileSize = ($item.Length).ToString()
            $fileNameOnly = $item.Name
            $filePath = $item.DirectoryName
            $extension = $item.Extension

            if ($driveLetter -and $mappedPath) {
                if ($filePath.StartsWith("$driveLetter`:")) {
                    $relativePath = $filepath.Substring(3)
                    $filePath = Join-Path -Path $mappedPath -ChildPath $relativePath
                }
            }


            # Get file time metadata
            $dtLastAccessedDoc = $item.LastAccessTime
            $dtCreated = $item.CreationTime
            $dtLastModified = $item.LastWriteTime

            # If file is read-only, set as read/write just to ensure file timestamps can be corrected if required
            $fileReadOnly = $false
            if((Get-Item $objFile).IsReadOnly -eq $true) {
                (Get-Item $objFile).IsReadOnly = $false
                $fileReadOnly = $true
            }

            # Update file timestamps if they have been changed
            if(($item.LastWriteTime -ne $dtLastModified) -or ($item.LastAccessTime -ne $dtLastAccessedDoc)) {
                $item.LastWriteTime = $dtLastModified
                Start-Sleep -Milliseconds 100   # If we don't pause here, the dates do not get updated correctly
                $item.LastAccessTime = $dtLastAccessedDoc
            }
            if($fileReadOnly -eq $true){
                $item.IsReadOnly = $fileReadOnly
            }
            $runTime = Get-Date
            $runTimeF = $runTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")

            if($item -ne $null) {
                #Write to data that file has been updated
                $listEntry = @($fileNameOnly, $filePath, $currentFileSize, $dtLastModified, $dtLastAccessedDoc, $dtCreated, $extension, $dtLastModified,$runTime)  -Join [char]9
                $filesScanned.Add($listEntry)
            }
        }
    }

    $currentTime = Get-Date
    $currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "Catalogue for $filename ended at $($currentTimeF)"
    Add-Content -Path $progressFile -Value $logEntry

    $currentTime = Get-Date
    $currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "Listing for $filename started at $($currentTimeF)"
    Add-Content -Path $progressFile -Value $logEntry


    # Create the output file
    New-Item -Path $filename -ItemType File -Force
    $filesScanned | Tee-Object -Append $filename

    $currentTime = Get-Date
    $currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "Listing for $filename ended at $($currentTimeF)"
    Add-Content -Path $progressFile -Value $logEntry
    
}

function Get-ApplicableFiles {
    param (
        [System.Collections.ArrayList]$Files,
        [string]$FolderName
    )

    # Ensure the folder exists
    if (!(Test-Path -Path $FolderName -PathType Container)) {
        Write-Error "Folder $FolderName does not exist."
        return
    }

    # Process root folder files
    $rootFiles = Get-ChildItem -Path $FolderName -File #| Where-Object {($_.LastAccessedTime -lt (Get-Date).AddDays(-540)) -and  ($_.CreationTime -lt (Get-Date).AddDays(-1095))} 
    foreach ($file in $rootFiles) {
        $Files.Add($file.FullName)

    }

    # Process subfolders
    $subFolders = Get-ChildItem -Path $FolderName -Directory
    foreach ($subFolder in $subFolders) {
        $subFiles = Get-ChildItem -Path $subFolder.FullName  -File #| Where-Object {($_.LastAccessedTime -lt (Get-Date).AddDays(-540)) -and  ($_.CreationTime -lt (Get-Date).AddDays(-1095))} 
        foreach ($file in $subFiles) {
            $Files.Add($file.FullName)
        }

        # Recursively process subfolders
        Write-Host "Recursing into subfolder: $($subFolder.FullName)"
        Get-ApplicableFiles -files $Files -FolderName $subFolder.FullName
    }

}



# Import functions
#. 'C:\users\BHATTK\OneDrive - RSA Group\Developers\Get-FileListing-Functions.ps1'

$server = "\\saazrsaundatfsprduks001.file.core.windows.net\"
$workingPath = "C:\temp\Unstructured\FileListing\"
$destinationPath = "$($workingPath)FileShare\"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$progressFile = "$($workingPath)$($timestamp)_FileShareListingProgress.txt"
New-Item -Path $progressFile -ItemType File -Force

$filesToScan = New-Object System.Collections.ArrayList

$folderToScan = "fsunstr01prd01"
$pathToScan = "$server$folderToScan"

#for ($letter = [char]'Z'; $letter -ge [char]'A'; $letter--) {
for ($test = 122;$test -ge 97;$test--) {
    $letter = [char]$test
    $drive = "$letter`:"
    if (-not (Test-Path $drive)) {
        # Map the drive
        New-PSDrive -Name $letter -PSProvider FileSystem -Root $pathToScan -Persist
        Write-Output "Mapped $drive to $pathToScan"

        $pathToScanNew = "$drive"
        break
    }
}

$filename = $folderToScan
$filenameInterim = "$($filename)_FileShareListing.txt"
$filenameZip = "$destinationPath$($timestamp)_$($filename)_FileShareListing.zip"
$filePath = "$destinationPath$filenameInterim"

if(Test-Path -Path  $filePath -PathType Leaf) {
    Compress-Archive -Path $filePath -DestinationPath $filenameZip
    Remove-Item -Path $filePath
}

$currentTime = Get-Date
$currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
$logEntry = "Scanning for $filename started at $($currentTimeF)"
Add-Content -Path $progressFile -Value $logEntry

# Create collection of files from which metadata is to be extracted
Get-ApplicableFiles -Files $filesToScan -FolderName $pathToScanNew
   
# Ensure only unique files are extracted (duplicates occur with nested folders)
$filesToScanUnique = $filesToScan | sort -Unique
    
$currentTime = Get-Date
$currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
$logEntry = "Scanning for $filename ended at $($currentTimeF)"
Add-Content -Path $progressFile -Value $logEntry

    if($filesToScan.Count -gt 0) {

        $currentTime = Get-Date
        $currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
        $logEntry = "Listing for $filename started at $($currentTimeF)"
        Add-Content -Path $progressFile -Value $logEntry

        # Create file list
        List-FileAgeProperties -Files $filesToScanUnique -Filename $filePath -progressFile $progressFile -driveLetter $letter -mappedPath $pathToScan

        $currentTime = Get-Date
        $currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
        $logEntry = "Listing for $filename ended at $($currentTimeF)"
        Add-Content -Path $progressFile -Value $logEntry

    }

try{
    Get-PSDrive $letter
    Remove-PSDrive
    }
catch{}
    
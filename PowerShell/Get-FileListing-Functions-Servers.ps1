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

clear

$workingPath = "C:\temp\Unstructured\FileListing\"
$destinationPath = "$($workingPath)Servers\"
$serverFile = "$($workingPath)Servers.txt"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Import list of servers into memory
$serverList = (Get-Content -Path $serverFile).Trim()

# Get list of files already created
$filemap = Get-ChildItem -Path $destinationPath | Group-Object {$_.BaseName} 

# Create lookup for quick access 
$fileLookup = @{} 
foreach ($group in $fileMap) { 
    $fileLookup[$group.Name] = $group.Group | Sort-Object CreationTime 
} 

# Separate servers 
$serversWithoutFiles = @() 
$serversWithFiles = @() 

# Loop through servers to get run order (servers without file lists in increasing size order, servers with file lists in increased creation date order)
foreach ($server in $serverList) { 
    if($server.Substring(0,2) -ne '--') {
        $trimmedServer = $server.Trim()
        $filename = ($trimmedServer.Replace('\','_').Replace(':',''))
        $filenameInterim = "$($filename)_DriveListing"
        if ($fileLookup.ContainsKey($filenameInterim)) { 
            # If file for server is already created
            $oldestFile = $fileLookup[$filenameInterim][0]
            if($oldestFile) {
                $serversWithFiles += [PSCustomObject]@{ 
                    Server = $trimmedServer 
                    File   = $oldestFile  # Oldest file 
                } 
            }
        } else { 
            $serversWithoutFiles += $trimmedServer 
        } 
    } 
}

# Sort servers with files by file creation date 
$sortedWithFiles = $serversWithFiles | Sort-Object { $_.File.CreationTime } | ForEach-Object { $_.Server } 

# Combine final list 
$finalList = $serversWithoutFiles + $sortedWithFiles 

$filesToScan = New-Object System.Collections.ArrayList

$progressFile = "$($workingPath)$($timestamp)_DriveListingProgress.txt"

New-Item -Path $progressFile -ItemType File -Force

# Loop through ordered list
foreach($server in $finalList) {
    $filename = ($server.Replace('\','_').Replace(':',''))
    $filenameInterim = "$($filename)_DriveListing.txt"
    $filenameZip = "$destinationPath$($timestamp)_$($filename)_DriveListing.zip"
    $filePath = "$destinationPath$filenameInterim"

    # Create zip of previous file list
    if(Test-Path -Path  $filePath -PathType Leaf) {
        Compress-Archive -Path $filePath -DestinationPath $filenameZip
        Remove-Item -Path $filePath
    }

    $filesToScan.Clear()

    $currentTime = Get-Date
    $currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "Scanning for $filename started at $($currentTimeF)"
    Add-Content -Path $progressFile -Value $logEntry

    # Create collection of files from which metadata is to be extracted
    Get-ApplicableFiles -Files $filesToScan -FolderName $server
        
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
        List-FileAgeProperties -Files $filesToScanUnique -Filename $filePath -progressFile $progressFile

        $currentTime = Get-Date
        $currentTimeF = $currentTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
        $logEntry = "Listing for $filename ended at $($currentTimeF)"
        Add-Content -Path $progressFile -Value $logEntry

    }
}

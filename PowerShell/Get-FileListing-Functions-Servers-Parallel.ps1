param(
    [Parameter(Mandatory=$true)]
    [string]$DrivePath,

    [Parameter(Mandatory=$false)]
    [int]$ParallelItems = 5,

    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 200,

    [Parameter(Mandatory=$false)]
    [string]$WorkingPath = "C:\temp\Unstructured\FileListing\",

    [Parameter(Mandatory=$false)]
    [string]$DriveLetter = $null,

    [Parameter(Mandatory=$false)]
    [string]$MappedPath = $null
)

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $DrivePath -PathType Container)) {
    Write-Error "DrivePath '$DrivePath' does not exist or is not accessible. Aborting."
    exit 1
}

if (($DriveLetter -and -not $MappedPath) -or (-not $DriveLetter -and $MappedPath)) {
    Write-Error "Both -DriveLetter and -MappedPath must be provided together, or neither."
    exit 1
}

if (-not (Test-Path -Path $WorkingPath -PathType Container)) {
    New-Item -Path $WorkingPath -ItemType Directory -Force | Out-Null
}

$destinationPath = $WorkingPath
$timestamp       = Get-Date -Format "yyyyMMdd_HHmmss"


# ---------------------------------------------------------------------------
# Add-ContentSafe
# Thread-safe content writer with retry, since multiple runspaces will be
# writing to the same progress log concurrently.
# ---------------------------------------------------------------------------
function Add-ContentSafe {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Value,
        [int]$MaxRetries = 5,
        [int]$RetryDelayMs = 200
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Add-Content -Path $Path -Value $Value -ErrorAction Stop
            return $true
        }
        catch {
            if ($attempt -eq $MaxRetries) {
                Write-Warning "Add-ContentSafe: failed to write to $Path after $MaxRetries retries: $($_.Exception.Message)"
                return $false
            }
            Start-Sleep -Milliseconds $RetryDelayMs
        }
    }
}


# ---------------------------------------------------------------------------
# Get-ApplicableFiles
# Recursively walks the drive/folder and builds the full file list.
# Unchanged from the original, single-threaded — this is a cheap filesystem
# walk and doesn't benefit meaningfully from parallelisation; the expensive
# part is the per-file metadata/timestamp work done afterwards.
#
# Excludes:
#   - Snapshot folders (any folder named "~snapshot", case-insensitive)
#     at any depth, so a single nested snapshot doesn't get walked and
#     doesn't multiply every file beneath it across snapshot generations.
# ---------------------------------------------------------------------------
function Get-ApplicableFiles {
    param (
        [System.Collections.ArrayList]$Files,
        [string]$FolderName
    )

    if (-not (Test-Path -Path $FolderName -PathType Container)) {
        Write-Error "Folder $FolderName does not exist."
        return
    }

    $rootFiles = Get-ChildItem -Path $FolderName -File -ErrorAction SilentlyContinue
    foreach ($file in $rootFiles) {
        $Files.Add($file.FullName) | Out-Null
    }

    $subFolders = Get-ChildItem -Path $FolderName -Directory -ErrorAction SilentlyContinue
    foreach ($subFolder in $subFolders) {

        if ($subFolder.Name -ieq "~snapshot") {
            Write-Host "Skipping snapshot folder: $($subFolder.FullName)"
            continue
        }

        $subFiles = Get-ChildItem -Path $subFolder.FullName -File -ErrorAction SilentlyContinue
        foreach ($file in $subFiles) {
            $Files.Add($file.FullName) | Out-Null
        }

        Write-Host "Recursing into subfolder: $($subFolder.FullName)"
        Get-ApplicableFiles -Files $Files -FolderName $subFolder.FullName
    }
}


# ---------------------------------------------------------------------------
# Process-FileListingBatch
# Runs a chunk of files through the metadata-extraction logic in its own
# runspace. Mirrors the structure used in the tagging script's batch
# functions — function bodies injected as named functions, results
# returned as objects rather than written directly (avoids contention on
# the output file itself; only the progress log is written concurrently).
# ---------------------------------------------------------------------------
function Process-FileListingBatch {
    param(
        [System.Collections.ArrayList]$batch,
        [int]$parallelItems,
        [string]$progressFile,
        [string]$driveLetter,
        [string]$mappedPath
    )

    if ($null -eq $batch -or $batch.Count -eq 0) { return @() }

    $pool = [runspacefactory]::CreateRunspacePool(1, $parallelItems)
    $pool.Open()

    $fnAddContentSafe = "function Add-ContentSafe { ${function:Add-ContentSafe} }"

    $jobs = @()

    foreach ($file in $batch) {

        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        $ps.AddScript($fnAddContentSafe) | Out-Null

        $ps.AddScript({
            param($objFile, $progressFile, $driveLetter, $mappedPath)

            try {
                $item = Get-Item -LiteralPath $objFile -ErrorAction Stop
            }
            catch {
                Add-ContentSafe -Path $progressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $objFile could not be opened: $($_.Exception.Message)"
                return $null
            }

            # Skip Office temp/lock files — these have the "~$" prefix on the
            # FILENAME itself (e.g. "~$document.docx"), not the path. The
            # original check tested the full path's first character, which
            # would incorrectly skip every file inside a folder that happened
            # to start with "~" (such as a literal "~snapshot" folder if it
            # weren't already excluded during the walk).
            if ($item.Name.StartsWith("~$")) {
                return $null
            }

            $currentFileSize = $item.Length.ToString()
            $fileNameOnly    = $item.Name
            $filePath        = $item.DirectoryName
            $extension       = $item.Extension

            if ($driveLetter -and $mappedPath) {
                if ($filePath.StartsWith("$driveLetter`:")) {
                    $relativePath = $filePath.Substring(3)
                    $filePath = Join-Path -Path $mappedPath -ChildPath $relativePath
                }
            }

            $dtLastAccessedDoc = $item.LastAccessTime
            $dtCreated         = $item.CreationTime
            $dtLastModified    = $item.LastWriteTime

            # Preserve original timestamps across the read — IsReadOnly flag
            # toggled only if needed, restored afterwards
            $fileReadOnly = $false
            try {
                if ($item.IsReadOnly -eq $true) {
                    $item.IsReadOnly = $false
                    $fileReadOnly = $true
                }

                if (($item.LastWriteTime -ne $dtLastModified) -or ($item.LastAccessTime -ne $dtLastAccessedDoc)) {
                    $item.LastWriteTime = $dtLastModified
                    Start-Sleep -Milliseconds 100
                    $item.LastAccessTime = $dtLastAccessedDoc
                }

                if ($fileReadOnly -eq $true) {
                    $item.IsReadOnly = $true
                }
            }
            catch {
                $msg     = $_.Exception.Message
                $hresult = if ($_.Exception.HResult) { '{0:X8}' -f $_.Exception.HResult } else { $null }
                if ($msg -match 'being used by another process' -or $hresult -eq '80070020') {
                    Add-ContentSafe -Path $progressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $objFile timestamp restore skipped; file in use"
                }
                else {
                    Add-ContentSafe -Path $progressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $objFile timestamp restore failed: $msg"
                }
            }

            $runTime  = Get-Date
            $runTimeF = $runTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")

            return [pscustomobject]@{
                Name             = $fileNameOnly
                ContainingPath   = $filePath
                Size             = $currentFileSize
                LastModified     = $dtLastModified
                LastAccessed     = $dtLastAccessedDoc
                CreationDate     = $dtCreated
                Extension        = $extension
                LastSaveDate     = $dtLastModified
                DateChecked      = $runTimeF
            }

        }) | Out-Null

        $ps.AddArgument($file)         | Out-Null
        $ps.AddArgument($progressFile) | Out-Null
        $ps.AddArgument($driveLetter)  | Out-Null
        $ps.AddArgument($mappedPath)   | Out-Null

        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }

    return [pscustomobject]@{ Jobs = $jobs; Pool = $pool }
}


# ---------------------------------------------------------------------------
# Wait-AndCollectJobs
# ---------------------------------------------------------------------------
function Wait-AndCollectJobs {
    param($BatchResult)

    $collected = @()

    if ($null -eq $BatchResult) { return $collected }

    foreach ($job in $BatchResult.Jobs) {
        if ($null -eq $job -or $null -eq $job.Pipe -or $null -eq $job.Handle) { continue }

        $job.Handle.AsyncWaitHandle.WaitOne()
        try {
            $output = $job.Pipe.EndInvoke($job.Handle)
            if ($output) { $collected += $output }
        }
        catch {
            Write-Warning "Runspace error: $($_.Exception.Message)"
        }
        finally {
            $job.Pipe.Dispose()
        }
    }

    if ($BatchResult.Pool) {
        $BatchResult.Pool.Close()
        $BatchResult.Pool.Dispose()
    }

    return $collected
}


# ---------------------------------------------------------------------------
# List-FileAgeProperties
# Chunks the unique file list into batches and dispatches them to
# Process-FileListingBatch, writing results to the output file as each
# batch completes (rather than holding everything in memory until the end).
# ---------------------------------------------------------------------------
function List-FileAgeProperties {
    param (
        [System.Collections.ArrayList]$Files,
        [string]$Filename,
        [string]$ProgressFile,
        [int]$ParallelItems,
        [int]$BatchSize,
        [string]$DriveLetter = $null,
        [string]$MappedPath = $null
    )

    if (-not (Test-Path -Path $ProgressFile -PathType Leaf)) {
        Write-Error "File $ProgressFile does not exist"
        return $null
    }

    $currentTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    Add-ContentSafe -Path $ProgressFile -Value "Catalogue for $Filename started at $currentTimeF"

    # Create output file with header
    New-Item -Path $Filename -ItemType File -Force | Out-Null
    $header = @("Name","Containing Path","Size","Last Modified","Last Accessed","Creation Date","Extension","Last Save Date","Date Checked") -Join [char]9
    Add-Content -Path $Filename -Value $header

    $totalFiles = $Files.Count
    $processed  = 0

    # Chunk the file list into batches for parallel dispatch
    for ($start = 0; $start -lt $totalFiles; $start += $BatchSize) {
        $end = [Math]::Min($start + $BatchSize - 1, $totalFiles - 1)
        $chunk = [System.Collections.ArrayList]@($Files[$start..$end])

        Add-ContentSafe -Path $ProgressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Dispatching batch $($start+1)-$($end+1) of $totalFiles"

        $batchResult = Process-FileListingBatch `
            -batch         $chunk `
            -parallelItems $ParallelItems `
            -progressFile  $ProgressFile `
            -driveLetter   $DriveLetter `
            -mappedPath    $MappedPath

        $rows = Wait-AndCollectJobs -BatchResult $batchResult

        foreach ($row in $rows) {
            $listEntry = @(
                $row.Name, $row.ContainingPath, $row.Size, $row.LastModified,
                $row.LastAccessed, $row.CreationDate, $row.Extension,
                $row.LastSaveDate, $row.DateChecked
            ) -Join [char]9
            Add-Content -Path $Filename -Value $listEntry
        }

        $processed += $chunk.Count
        Write-Progress -Activity "Scanning files" -Status "$processed of $totalFiles" -PercentComplete (($processed / $totalFiles) * 100)
    }

    Write-Progress -Activity "Scanning files" -Completed

    $currentTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    Add-ContentSafe -Path $ProgressFile -Value "Catalogue for $Filename ended at $currentTimeF"
}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$filename     = ($DrivePath.Replace('\','_').Replace(':',''))
$filePath     = "$destinationPath$($timestamp)_$($filename)_DriveListing.txt"
$progressFile = "$destinationPath$($timestamp)_DriveListingProgress.txt"

New-Item -Path $progressFile -ItemType File -Force | Out-Null

$filesToScan = New-Object System.Collections.ArrayList

$currentTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-ContentSafe -Path $progressFile -Value "Scanning for $DrivePath started at $currentTimeF"

Get-ApplicableFiles -Files $filesToScan -FolderName $DrivePath

$filesToScanUnique = $filesToScan | Sort-Object -Unique

$currentTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-ContentSafe -Path $progressFile -Value "Scanning for $DrivePath ended at $currentTimeF — $($filesToScanUnique.Count) unique files found"

if ($filesToScanUnique.Count -gt 0) {
    List-FileAgeProperties `
        -Files         ([System.Collections.ArrayList]$filesToScanUnique) `
        -Filename      $filePath `
        -ProgressFile  $progressFile `
        -ParallelItems $ParallelItems `
        -BatchSize     $BatchSize `
        -DriveLetter   $DriveLetter `
        -MappedPath    $MappedPath

    Write-Host "Listing complete: $filePath"
}
else {
    Write-Host "No files found under $DrivePath"
}

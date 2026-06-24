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
# Process-FileListingBatch
# Runs a chunk of files through the metadata-extraction logic in its own
# runspace pool. Returns a batch-result object for Wait-AndCollectJobs.
# ---------------------------------------------------------------------------
function Process-FileListingBatch {
    param(
        [string[]]$batch,
        [int]$parallelItems,
        [string]$progressFile,
        [string]$driveLetter,
        [string]$mappedPath
    )

    if ($null -eq $batch -or $batch.Count -eq 0) { return $null }

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

            # Skip Office temp/lock files — "~$" prefix on the filename itself.
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

            # Preserve original timestamps across the read
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

            $runTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")

            return [pscustomobject]@{
                Name           = $fileNameOnly
                ContainingPath = $filePath
                Size           = $currentFileSize
                LastModified   = $dtLastModified
                LastAccessed   = $dtLastAccessedDoc
                CreationDate   = $dtCreated
                Extension      = $extension
                LastSaveDate   = $dtLastModified
                DateChecked    = $runTimeF
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

    $collected = [System.Collections.Generic.List[object]]::new()

    if ($null -eq $BatchResult) { return $collected }

    foreach ($job in $BatchResult.Jobs) {
        if ($null -eq $job -or $null -eq $job.Pipe -or $null -eq $job.Handle) { continue }

        $job.Handle.AsyncWaitHandle.WaitOne()
        try {
            $output = $job.Pipe.EndInvoke($job.Handle)
            if ($output) { $collected.AddRange([object[]]$output) }
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
# Invoke-FlushBatch
# Dispatches a pending batch to the runspace pool, collects results, and
# writes them immediately to the output file. Called from within the
# directory walk so that files are written as soon as a batch fills up
# rather than after the entire tree has been enumerated.
# ---------------------------------------------------------------------------
function Invoke-FlushBatch {
    param(
        [string[]]$batch,
        [string]$outputFile,
        [string]$progressFile,
        [int]$parallelItems,
        [string]$driveLetter,
        [string]$mappedPath,
        [ref]$processedCount
    )

    if ($batch.Count -eq 0) { return }

    $batchResult = Process-FileListingBatch `
        -batch         $batch `
        -parallelItems $parallelItems `
        -progressFile  $progressFile `
        -driveLetter   $driveLetter `
        -mappedPath    $mappedPath

    $rows = Wait-AndCollectJobs -BatchResult $batchResult

    foreach ($row in $rows) {
        $listEntry = @(
            $row.Name, $row.ContainingPath, $row.Size, $row.LastModified,
            $row.LastAccessed, $row.CreationDate, $row.Extension,
            $row.LastSaveDate, $row.DateChecked
        ) -Join [char]9
        Add-Content -Path $outputFile -Value $listEntry
    }

    $processedCount.Value += $batch.Count
}


# ---------------------------------------------------------------------------
# Walk-AndProcess
# Recursive directory walk that fills a rolling batch buffer and flushes it
# to disk as soon as it reaches $BatchSize — without ever holding the full
# file list in memory. The $seenPaths HashSet deduplicates across the walk
# (replacing the post-walk Sort-Object -Unique that previously required the
# entire list to be in memory at once).
#
# Excludes snapshot folders (any folder named "~snapshot", case-insensitive).
# ---------------------------------------------------------------------------
function Walk-AndProcess {
    param(
        [string]$FolderName,
        [string]$OutputFile,
        [string]$ProgressFile,
        [int]$ParallelItems,
        [int]$BatchSize,
        [string]$DriveLetter,
        [string]$MappedPath,
        [System.Collections.Generic.HashSet[string]]$SeenPaths,
        [string[]]$PendingBatch,          # passed by value; returned via [ref] pattern below
        [ref]$PendingBatchRef,
        [ref]$ProcessedCount
    )

    # Enumerate files in this folder and add unseen ones to the pending batch
    $localFiles = Get-ChildItem -Path $FolderName -File -ErrorAction SilentlyContinue
    foreach ($file in $localFiles) {
        $fp = $file.FullName
        if ($SeenPaths.Add($fp)) {           # Add returns $false if already present
            $PendingBatchRef.Value += $fp

            if ($PendingBatchRef.Value.Count -ge $BatchSize) {
                Add-ContentSafe -Path $ProgressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Flushing batch at $($ProcessedCount.Value + $PendingBatchRef.Value.Count) files (in $FolderName)"
                Invoke-FlushBatch `
                    -batch          $PendingBatchRef.Value `
                    -outputFile     $OutputFile `
                    -progressFile   $ProgressFile `
                    -parallelItems  $ParallelItems `
                    -driveLetter    $DriveLetter `
                    -mappedPath     $MappedPath `
                    -processedCount $ProcessedCount
                $PendingBatchRef.Value = @()
                [System.GC]::Collect()
            }
        }
    }

    # Recurse into sub-folders
    $subFolders = Get-ChildItem -Path $FolderName -Directory -ErrorAction SilentlyContinue
    foreach ($subFolder in $subFolders) {

        if ($subFolder.Name -ieq "~snapshot") {
            Write-Host "Skipping snapshot folder: $($subFolder.FullName)"
            continue
        }

        Write-Host "Recursing into subfolder: $($subFolder.FullName)"
        Walk-AndProcess `
            -FolderName      $subFolder.FullName `
            -OutputFile      $OutputFile `
            -ProgressFile    $ProgressFile `
            -ParallelItems   $ParallelItems `
            -BatchSize       $BatchSize `
            -DriveLetter     $DriveLetter `
            -MappedPath      $MappedPath `
            -SeenPaths       $SeenPaths `
            -PendingBatch    $PendingBatchRef.Value `
            -PendingBatchRef $PendingBatchRef `
            -ProcessedCount  $ProcessedCount
    }
}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$filename     = ($DrivePath.Replace('\','_').Replace(':',''))
$filePath     = "$destinationPath$($timestamp)_$($filename)_DriveListing.txt"
$progressFile = "$destinationPath$($timestamp)_DriveListingProgress.txt"

New-Item -Path $progressFile -ItemType File -Force | Out-Null

# Create output file with header immediately
New-Item -Path $filePath -ItemType File -Force | Out-Null
$header = @("Name","Containing Path","Size","Last Modified","Last Accessed","Creation Date","Extension","Last Save Date","Date Checked") -Join [char]9
Add-Content -Path $filePath -Value $header

$currentTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-ContentSafe -Path $progressFile -Value "Scanning $DrivePath started at $currentTimeF"

# Shared state threaded through the recursive walk
$seenPaths      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$pendingBatch   = @()
$pendingBatchR  = [ref]$pendingBatch
$processedCount = [ref]0

Walk-AndProcess `
    -FolderName      $DrivePath `
    -OutputFile      $filePath `
    -ProgressFile    $progressFile `
    -ParallelItems   $ParallelItems `
    -BatchSize       $BatchSize `
    -DriveLetter     $DriveLetter `
    -MappedPath      $MappedPath `
    -SeenPaths       $seenPaths `
    -PendingBatch    $pendingBatchR.Value `
    -PendingBatchRef $pendingBatchR `
    -ProcessedCount  $processedCount

# Flush any remaining files that didn't fill a complete batch
if ($pendingBatchR.Value.Count -gt 0) {
    Add-ContentSafe -Path $progressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Flushing final batch of $($pendingBatchR.Value.Count) files"
    Invoke-FlushBatch `
        -batch          $pendingBatchR.Value `
        -outputFile     $filePath `
        -progressFile   $progressFile `
        -parallelItems  $ParallelItems `
        -driveLetter    $DriveLetter `
        -mappedPath     $MappedPath `
        -processedCount $processedCount
}

$currentTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-ContentSafe -Path $progressFile -Value "Scanning $DrivePath ended at $currentTimeF — $($processedCount.Value) files processed"

Write-Host "Listing complete: $filePath"

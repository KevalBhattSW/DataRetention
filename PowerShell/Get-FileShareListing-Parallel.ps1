<#
.SYNOPSIS
    Parallel, batched, resumable file-metadata listing for a single Azure
    file share. Produces one tab-delimited output file for the whole share.

.DESCRIPTION
    This is the Azure-file-share counterpart to
    Get-FileListing-Functions-Servers-Parallel.ps1. It keeps that script's
    architecture — runspace-pool parallelism, streamed batch flushing (never
    holding the full file list in memory), thread-safe progress logging,
    snapshot-folder exclusion, and R/O/X resume detection — but folds in the
    drive-mapping that the original sequential FileShareListing script did:

      1. Finds a free drive letter (Z -> A) and maps it to the Azure file
         share (\\<account>.file.core.windows.net\<share>) with New-PSDrive.
      2. Walks the mapped drive in parallel, writing one listing row per file.
         Stored ContainingPath values are translated from the temporary drive
         letter back to the share's UNC path, so the output is portable and
         doesn't depend on which letter happened to be free.
      3. Always unmaps the drive on exit (success, failure, or Ctrl-C via the
         finally block).

    Output is a SINGLE file per share (fixed name, no datestamp) so a resumed
    run appends to the same file. The progress log is datestamped at the
    working-path root so each run/resume gets its own clean log.

    The tab-delimited columns are identical to the servers listing and the
    catalogue listing, so Load-FileShareDataToSql.ps1 (companion loader) can
    stage it into SQL with the same enrichment logic.

.NOTES
    Re-sign this script with your usual Sign-* signing script before running
    under an execution policy that requires signed scripts — the edits here
    mean any previous signature block no longer matches.
#>

param(
    # Azure file share account root, including trailing backslash, e.g.
    # \\saazrsaundatfsprduks001.file.core.windows.net\
    [Parameter(Mandatory=$false)]
    [string]$Server = "\\saazrsaundatfsprduks001.file.core.windows.net\",

    # The share (top-level folder) under $Server to scan, e.g. fsunstr01prd01
    [Parameter(Mandatory=$false)]
    [string]$ShareName = "fsunstr01prd01",

    [Parameter(Mandatory=$false)]
    [int]$ParallelItems = 5,

    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 1000,

    [Parameter(Mandatory=$false)]
    [string]$WorkingPath = "\\AZUKSWVPUNSD01\Temp\Unstructured\FileListing\",

    # Optional credential for the file share. Leave $null to use the ambient
    # identity (matches the original script). For a storage-account-key mount,
    # supply a PSCredential whose username is "AZURE\<storageaccountname>"
    # (or "localhost\<storageaccountname>") and whose password is the key.
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential = $null
)

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
$Server    = $Server.TrimEnd('\') + '\'          # normalise trailing slash
$ShareRoot = "$Server$ShareName"                 # \\...\<share>

if (-not (Test-Path -Path $WorkingPath -PathType Container)) {
    New-Item -Path $WorkingPath -ItemType Directory -Force | Out-Null
}

# Single output for the whole share lives in a 'FileShare' subfolder — fixed
# name, no datestamp, so a resumed run appends to the same file.
$fileSharePath = Join-Path -Path $WorkingPath -ChildPath "FileShare"
if (-not (Test-Path -Path $fileSharePath -PathType Container)) {
    New-Item -Path $fileSharePath -ItemType Directory -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"


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

            # Translate the temporary drive letter back to the share's UNC path
            # so the output doesn't depend on which letter was free this run.
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
            if ($output) {
                $valid = [object[]]($output | Where-Object { $_ -ne $null })
                if ($valid.Count -gt 0) { $collected.AddRange($valid) }
            }
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
        if ($null -eq $row -or -not $row.Name -or -not $row.ContainingPath) { continue }
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
# file list in memory. The $seenPaths HashSet deduplicates across the walk.
#
# Excludes snapshot folders (any folder named "~snapshot", case-insensitive)
# — Azure file shares expose share snapshots under a ~snapshot directory.
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
# Map the Azure file share to a free drive letter (Z -> A)
# ---------------------------------------------------------------------------
$mappedLetter = $null

try {

    for ($code = 122; $code -ge 97; $code--) {          # 'z' .. 'a'
        $letter = [char]$code
        $drive  = "$letter`:"
        if (-not (Test-Path $drive)) {
            $newDriveParams = @{
                Name       = $letter
                PSProvider = 'FileSystem'
                Root       = $ShareRoot
                Persist    = $true
                ErrorAction = 'Stop'
            }
            if ($Credential) { $newDriveParams['Credential'] = $Credential }

            New-PSDrive @newDriveParams | Out-Null
            $mappedLetter = $letter
            Write-Output "Mapped $drive to $ShareRoot"
            break
        }
    }

    if (-not $mappedLetter) {
        Write-Error "No free drive letter (A-Z) available to map $ShareRoot. Aborting."
        exit 1
    }

    $DriveLetter = [string]$mappedLetter
    $MappedPath  = $ShareRoot
    $DrivePath   = "$mappedLetter`:"

    # -----------------------------------------------------------------------
    # Paths — single output per share, datestamped progress log
    # -----------------------------------------------------------------------
    $filePath     = Join-Path -Path $fileSharePath -ChildPath "$($ShareName)_FileShareListing.txt"
    $progressFile = Join-Path -Path $WorkingPath   -ChildPath "$($timestamp)_FileShareListingProgress.txt"

    New-Item -Path $progressFile -ItemType File -Force | Out-Null

    # -----------------------------------------------------------------------
    # Resume detection — prompt if a listing file already exists for this share
    # -----------------------------------------------------------------------
    $isResume  = $false
    $seenPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if (Test-Path -LiteralPath $filePath -PathType Leaf) {

        Write-Host ""
        Write-Host "An existing listing file was found for this share:" -ForegroundColor Yellow
        Write-Host "  $filePath" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  [R] Resume    — continue from where the previous run stopped, appending to this file"
        Write-Host "  [O] Overwrite — delete the existing file and start fresh"
        Write-Host "  [X] Exit      — abort this run"
        Write-Host ""

        do {
            $choice = Read-Host "Enter choice (R / O / X)"
        } while ($choice -notmatch '^[ROXrox]$')

        switch ($choice.ToUpper()) {

            'R' {
                $isResume = $true
                Write-Host "Resuming. Loading already-processed paths from existing listing..." -ForegroundColor Cyan
                Add-ContentSafe -Path $progressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') RESUME run started for $ShareRoot — loading existing listing from $filePath"

                # Read every non-header data row and reconstruct the full path
                # from columns: Name (col 0) and ContainingPath (col 1). The
                # stored path is the share UNC, so reverse-map it back to the
                # temporary drive letter to match what Get-ChildItem yields.
                $rowCount = 0
                Get-Content -LiteralPath $filePath |
                    Select-Object -Skip 1 |   # skip header row
                    ForEach-Object {
                        $cols = $_ -split [char]9
                        if ($cols.Count -ge 2 -and $cols[0] -and $cols[1]) {
                            $storedPath = Join-Path -Path $cols[1] -ChildPath $cols[0]

                            if ($DriveLetter -and $MappedPath -and $storedPath.StartsWith($MappedPath)) {
                                $relative   = $storedPath.Substring($MappedPath.Length).TrimStart('\')
                                $storedPath = "$DriveLetter`:\$relative"
                            }

                            $seenPaths.Add($storedPath) | Out-Null
                            $rowCount++
                        }
                    }

                Write-Host "  Loaded $rowCount previously-processed paths. Walk will skip these." -ForegroundColor Cyan
                Add-ContentSafe -Path $progressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Loaded $rowCount previously-processed paths from existing listing"
            }

            'O' {
                Write-Host "Overwriting existing listing." -ForegroundColor Cyan
                Add-ContentSafe -Path $progressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') OVERWRITE — deleting existing listing and starting fresh for $ShareRoot"
                Remove-Item -LiteralPath $filePath -Force
            }

            'X' {
                Write-Host "Exiting." -ForegroundColor Yellow
                Add-ContentSafe -Path $progressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Run aborted by user at startup prompt"
                exit 0
            }
        }
    }

    # Create the listing file (fresh run) or leave it in place (resume).
    # Header is only written for a fresh file.
    if (-not $isResume) {
        New-Item -Path $filePath -ItemType File -Force | Out-Null
        $header = @("Name","Containing Path","Size","Last Modified","Last Accessed","Creation Date","Extension","Last Save Date","Date Checked") -Join [char]9
        Add-Content -Path $filePath -Value $header
    }

    $currentTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    Add-ContentSafe -Path $progressFile -Value "Scanning $ShareRoot started at $currentTimeF"

    # Shared state threaded through the recursive walk.
    # $seenPaths is already populated if this is a resume.
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
    Add-ContentSafe -Path $progressFile -Value "Scanning $ShareRoot ended at $currentTimeF — $($processedCount.Value) new files added this run"

    Write-Host "Listing complete: $filePath"
}
finally {
    # Always unmap the drive, whatever happened above.
    if ($mappedLetter) {
        try {
            Remove-PSDrive -Name $mappedLetter -Force -ErrorAction Stop
            Write-Output "Unmapped $mappedLetter`: from $ShareRoot"
        }
        catch {
            Write-Warning "Failed to unmap drive $mappedLetter`: — $($_.Exception.Message)"
        }
    }
}

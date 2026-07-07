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
# Runs a chunk of files through the metadata-extraction logic in parallel
# runspaces. Returns a batch result for Wait-AndCollectJobs to drain.
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
            # filename itself (e.g. "~$document.docx"), not the path.
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
# Invoke-StreamingFileListing
#
# Replaces the old two-phase approach (Get-ApplicableFiles to build a full
# list, then Sort-Object -Unique, then List-FileAgeProperties) with a single
# streaming pass that:
#
#   1. Walks the directory tree using a stack and [System.IO.Directory]
#      instead of Get-ChildItem — significantly faster on large shares
#      (avoids PowerShell pipeline overhead per file).
#
#   2. Deduplicates using a HashSet (O(1) per insert) rather than
#      Sort-Object -Unique which requires holding all paths in memory
#      and sorting them — impractical at 25M files.
#
#   3. Dispatches batches to Process-FileListingBatch as soon as
#      $BatchSize files accumulate rather than waiting for the full
#      tree walk to finish — keeps memory usage flat regardless of
#      drive size.
#
#   4. Writes each completed batch's rows directly to the output file
#      so the output grows incrementally and the run is resumable if
#      interrupted.
# ---------------------------------------------------------------------------
function Invoke-StreamingFileListing {
    param(
        [string]$RootPath,
        [string]$OutputFile,
        [string]$ProgressFile,
        [int]$ParallelItems,
        [int]$BatchSize,
        [string]$DriveLetter = $null,
        [string]$MappedPath  = $null
    )

    # Create output file with header
    New-Item -Path $OutputFile -ItemType File -Force | Out-Null
    $header = @("Name","Containing Path","Size","Last Modified","Last Accessed",
                "Creation Date","Extension","Last Save Date","Date Checked") -Join [char]9
    Add-Content -Path $OutputFile -Value $header

    $startTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    Add-ContentSafe -Path $ProgressFile -Value "Catalogue for $OutputFile started at $startTimeF"

    # Deduplication set and running batch
    $seen  = [System.Collections.Generic.HashSet[string]]::new(
                 [System.StringComparer]::OrdinalIgnoreCase)
    $batch = [System.Collections.ArrayList]::new()

    $totalFound     = 0
    $totalProcessed = 0
    $batchNumber    = 0

    # Stack-based iterative directory walk — avoids PowerShell recursion
    # depth limits on deep trees and is faster than recursive calls
    $stack = [System.Collections.Generic.Stack[string]]::new()
    $stack.Push($RootPath)

    while ($stack.Count -gt 0) {

        $currentDir = $stack.Pop()

        # Enumerate files in current directory
        try {
            $dirFiles = [System.IO.Directory]::GetFiles($currentDir)
        }
        catch {
            Add-ContentSafe -Path $ProgressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Cannot enumerate files in $currentDir : $($_.Exception.Message)"
            $dirFiles = @()
        }

        foreach ($file in $dirFiles) {
            if ($seen.Add($file)) {
                $batch.Add($file) | Out-Null
                $totalFound++

                # Dispatch when batch is full
                if ($batch.Count -ge $BatchSize) {
                    $batchNumber++
                    Add-ContentSafe -Path $ProgressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Dispatching batch $batchNumber — $totalFound files found so far"

                    $batchCopy   = [System.Collections.ArrayList]@($batch)
                    $batchResult = Process-FileListingBatch `
                        -batch         $batchCopy `
                        -parallelItems $ParallelItems `
                        -progressFile  $ProgressFile `
                        -driveLetter   $DriveLetter `
                        -mappedPath    $MappedPath

                    $rows = Wait-AndCollectJobs -BatchResult $batchResult
                    foreach ($row in $rows) {
                        if ($null -eq $row) { continue }
                        $line = @($row.Name, $row.ContainingPath, $row.Size,
                                  $row.LastModified, $row.LastAccessed,
                                  $row.CreationDate, $row.Extension,
                                  $row.LastSaveDate, $row.DateChecked) -Join [char]9
                        Add-Content -Path $OutputFile -Value $line
                    }

                    $totalProcessed += $batchCopy.Count
                    $batch.Clear()

                    Write-Progress -Activity "Cataloguing files" `
                                   -Status "$totalProcessed files written, $totalFound found" `
                                   -PercentComplete -1   # indeterminate — total unknown upfront
                }
            }
        }

        # Push subdirectories onto stack (skip snapshot folders)
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($currentDir)
        }
        catch {
            Add-ContentSafe -Path $ProgressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Cannot enumerate subdirs in $currentDir : $($_.Exception.Message)"
            $subDirs = @()
        }

        foreach ($sub in $subDirs) {
            $leafName = [System.IO.Path]::GetFileName($sub)
            if ($leafName -ieq "~snapshot") {
                Write-Host "Skipping snapshot folder: $sub"
                Add-ContentSafe -Path $ProgressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Skipping snapshot folder: $sub"
                continue
            }
            $stack.Push($sub)
        }
    }

    # Drain any remaining files in the last partial batch
    if ($batch.Count -gt 0) {
        $batchNumber++
        Add-ContentSafe -Path $ProgressFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Dispatching final batch $batchNumber — $($batch.Count) remaining files"

        $batchResult = Process-FileListingBatch `
            -batch         $batch `
            -parallelItems $ParallelItems `
            -progressFile  $ProgressFile `
            -driveLetter   $DriveLetter `
            -mappedPath    $MappedPath

        $rows = Wait-AndCollectJobs -BatchResult $batchResult
        foreach ($row in $rows) {
            if ($null -eq $row) { continue }
            $line = @($row.Name, $row.ContainingPath, $row.Size,
                      $row.LastModified, $row.LastAccessed,
                      $row.CreationDate, $row.Extension,
                      $row.LastSaveDate, $row.DateChecked) -Join [char]9
            Add-Content -Path $OutputFile -Value $line
        }

        $totalProcessed += $batch.Count
    }

    Write-Progress -Activity "Cataloguing files" -Completed

    $endTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    Add-ContentSafe -Path $ProgressFile -Value "Catalogue for $OutputFile ended at $endTimeF — $totalFound unique files found, $totalProcessed rows written"

    return $totalFound
}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$filename     = ($DrivePath.Replace('\','_').Replace(':',''))
$filePath     = "$destinationPath$($timestamp)_$($filename)_DriveListing.txt"
$progressFile = "$destinationPath$($timestamp)_DriveListingProgress.txt"

New-Item -Path $progressFile -ItemType File -Force | Out-Null

$startTimeF = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-ContentSafe -Path $progressFile -Value "Drive listing for $DrivePath started at $startTimeF"
Write-Host "Starting catalogue of $DrivePath ..."
Write-Host "Output : $filePath"
Write-Host "Log    : $progressFile"
Write-Host ""

$totalFound = Invoke-StreamingFileListing `
    -RootPath      $DrivePath `
    -OutputFile    $filePath `
    -ProgressFile  $progressFile `
    -ParallelItems $ParallelItems `
    -BatchSize     $BatchSize `
    -DriveLetter   $DriveLetter `
    -MappedPath    $MappedPath

if ($totalFound -gt 0) {
    Write-Host "Catalogue complete: $filePath ($totalFound files)"
}
else {
    Write-Host "No files found under $DrivePath"
}

param(
    [Parameter(Mandatory=$true)]
    [string]$DrivePath,
    [Parameter(Mandatory=$true)]
    [string]$ScriptPath,
    [Parameter(Mandatory=$false)]
    [ValidateSet("Depth","Full","Both")]
    [string]$Mode = "Both",
    [Parameter(Mandatory=$false)]
    [int]$MaxDepth = 6,
    [Parameter(Mandatory=$false)]
    [int]$ParallelItems = 8
)

if (-not (Test-Path -LiteralPath $DrivePath -PathType Container)) {
    Write-Error "DrivePath '$DrivePath' does not exist or is not accessible. Aborting."
    Exit 1
}

# ---------------------------------------------------------------------------
# Helper: Add-ContentSafe
# Thread-safe file append used by runspaces
# ---------------------------------------------------------------------------
function Add-ContentSafe {
    param(
        [string]$Path,
        [string]$Value,
        [int]$MaxRetries = 5,
        [int]$DelayMs    = 50
    )
    for ($i = 0; $i -lt $MaxRetries; $i++) {
        try {
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Append,
                                         [System.IO.FileAccess]::Write,
                                         [System.IO.FileShare]::None)
            $sw = New-Object System.IO.StreamWriter($fs)
            $sw.WriteLine($Value)
            $sw.Close()
            $fs.Close()
            return
        }
        catch {
            Start-Sleep -Milliseconds $DelayMs
        }
    }
    Write-Warning "Failed to write to $Path after $MaxRetries retries"
}

# ---------------------------------------------------------------------------
# Helper: Get-FolderDepth
# Returns the depth of a path relative to the root drive path
# ---------------------------------------------------------------------------
function Get-FolderDepth {
    param([string]$FolderPath, [string]$RootPath)
    $relative = $FolderPath.Substring($RootPath.TrimEnd('\').Length).TrimStart('\')
    if ($relative -eq '') { return 0 }
    return ($relative -split '\\').Count
}

# ---------------------------------------------------------------------------
# Core scan function — called per top-level subfolder in a runspace
# Uses robocopy /L /NFL /NDL /NJH /NJS /NC /BYTES to get counts fast
# without loading all filenames into PowerShell memory
# ---------------------------------------------------------------------------
function Get-FolderStats {
    param(
        [string]$FolderPath,
        [string]$RootPath,
        [string]$OutputCsv,
        [string]$LogPath,
        [bool]$DepthLimited,
        [int]$MaxDepth
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        # Build list of folders to scan at the right depth
        $foldersToScan = [System.Collections.Generic.List[string]]::new()
        $foldersToScan.Add($FolderPath)

        # Get all subfolders up to MaxDepth (for depth-limited) or all (for full)
        $queue = [System.Collections.Generic.Queue[string]]::new()
        $queue.Enqueue($FolderPath)

        while ($queue.Count -gt 0) {
            $current = $queue.Dequeue()
            $depth   = Get-FolderDepth -FolderPath $current -RootPath $RootPath

            try {
                $subDirs = [System.IO.Directory]::GetDirectories($current)
            }
            catch {
                $subDirs = @()
            }

            foreach ($sub in $subDirs) {
                $subDepth = Get-FolderDepth -FolderPath $sub -RootPath $RootPath
                $foldersToScan.Add($sub)

                if (-not $DepthLimited -or $subDepth -lt $MaxDepth) {
                    $queue.Enqueue($sub)
                }
            }
        }

        # For each folder, count direct children only (non-recursive per folder)
        foreach ($folder in $foldersToScan) {
            $depth = Get-FolderDepth -FolderPath $folder -RootPath $RootPath

            if ($DepthLimited -and $depth -gt $MaxDepth) { continue }

            $fileCount      = 0
            $subFolderCount = 0
            $totalSizeBytes = 0L

            try {
                $di = [System.IO.DirectoryInfo]::new($folder)

                # Count direct files
                try {
                    $files = $di.GetFiles()
                    $fileCount = $files.Count
                    foreach ($f in $files) {
                        try { $totalSizeBytes += $f.Length } catch {}
                    }
                }
                catch { $fileCount = -1 }

                # Count direct subfolders
                try {
                    $subFolderCount = $di.GetDirectories().Count
                }
                catch { $subFolderCount = -1 }
            }
            catch {
                $fileCount      = -1
                $subFolderCount = -1
            }

            $sizeMB = if ($totalSizeBytes -ge 0) {
                [math]::Round($totalSizeBytes / 1MB, 2)
            } else { -1 }

            $relative = $folder.Substring($RootPath.TrimEnd('\').Length).TrimStart('\')
            if ($relative -eq '') { $relative = '\' }

            $results.Add([PSCustomObject]@{
                FolderPath      = $folder
                RelativePath    = $relative
                Depth           = $depth
                DirectFiles     = $fileCount
                DirectSubfolders = $subFolderCount
                DirectSizeMB    = $sizeMB
            })
        }
    }
    catch {
        Add-ContentSafe -Path $LogPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ERROR scanning $FolderPath : $($_.Exception.Message)"
    }

    # Write results to CSV (thread-safe, one row at a time)
    foreach ($row in $results) {
        $line = '"{0}","{1}",{2},{3},{4},{5}' -f `
            $row.FolderPath, $row.RelativePath, $row.Depth,
            $row.DirectFiles, $row.DirectSubfolders, $row.DirectSizeMB
        Add-ContentSafe -Path $OutputCsv -Value $line
    }

    return $results.Count
}

# ---------------------------------------------------------------------------
# Parallel dispatch — splits top-level subfolders across runspaces
# ---------------------------------------------------------------------------
function Invoke-ParallelFolderScan {
    param(
        [string]$RootPath,
        [string]$OutputCsv,
        [string]$LogPath,
        [bool]$DepthLimited,
        [int]$MaxDepth,
        [int]$ParallelItems
    )

    # Write CSV header
    Add-ContentSafe -Path $OutputCsv -Value '"FolderPath","RelativePath","Depth","DirectFiles","DirectSubfolders","DirectSizeMB"'

    # Write root folder itself first
    $rootDi         = [System.IO.DirectoryInfo]::new($RootPath)
    $rootFiles      = try { $rootDi.GetFiles() }      catch { @() }
    $rootSubDirs    = try { $rootDi.GetDirectories() } catch { @() }
    $rootSizeBytes  = ($rootFiles | ForEach-Object { try { $_.Length } catch { 0 } } | Measure-Object -Sum).Sum
    $rootSizeMB     = [math]::Round($rootSizeBytes / 1MB, 2)

    $rootLine = '"{0}","{1}",{2},{3},{4},{5}' -f `
        $RootPath, '\', 0, $rootFiles.Count, $rootSubDirs.Count, $rootSizeMB
    Add-ContentSafe -Path $OutputCsv -Value $rootLine

    # Get top-level subfolders to distribute across runspaces
    $topLevel = try {
        [System.IO.Directory]::GetDirectories($RootPath)
    } catch { @() }

    if ($topLevel.Count -eq 0) {
        Write-Host "No subfolders found under $RootPath"
        return
    }

    Write-Host "Scanning $($topLevel.Count) top-level folders with $ParallelItems parallel runspaces..."

    $fnAddContentSafe  = "function Add-ContentSafe { ${function:Add-ContentSafe} }"
    $fnGetFolderDepth  = "function Get-FolderDepth { ${function:Get-FolderDepth} }"
    $fnGetFolderStats  = "function Get-FolderStats { ${function:Get-FolderStats} }"

    $pool = [runspacefactory]::CreateRunspacePool(1, $ParallelItems)
    $pool.Open()

    $jobs = @()

    foreach ($topFolder in $topLevel) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        $ps.AddScript($fnAddContentSafe)  | Out-Null
        $ps.AddScript($fnGetFolderDepth)  | Out-Null
        $ps.AddScript($fnGetFolderStats)  | Out-Null

        $ps.AddScript({
            param($FolderPath, $RootPath, $OutputCsv, $LogPath, $DepthLimited, $MaxDepth)

            $count = Get-FolderStats `
                -FolderPath    $FolderPath `
                -RootPath      $RootPath `
                -OutputCsv     $OutputCsv `
                -LogPath       $LogPath `
                -DepthLimited  $DepthLimited `
                -MaxDepth      $MaxDepth

            Write-Output "Scanned $FolderPath — $count folders"
        }) | Out-Null

        $ps.AddArgument($topFolder)    | Out-Null
        $ps.AddArgument($RootPath)     | Out-Null
        $ps.AddArgument($OutputCsv)    | Out-Null
        $ps.AddArgument($LogPath)      | Out-Null
        $ps.AddArgument($DepthLimited) | Out-Null
        $ps.AddArgument($MaxDepth)     | Out-Null

        $jobs += [PSCustomObject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }

    # Wait and collect
    $completed = 0
    foreach ($job in $jobs) {
        $job.Handle.AsyncWaitHandle.WaitOne() | Out-Null
        try {
            $output = $job.Pipe.EndInvoke($job.Handle)
            foreach ($line in $output) {
                Write-Host $line
            }
        }
        catch {
            Write-Warning "Runspace error: $($_.Exception.Message)"
        }
        finally {
            $job.Pipe.Dispose()
        }
        $completed++
        if ($completed % 10 -eq 0) {
            Write-Host "$completed / $($jobs.Count) top-level folders complete..."
        }
    }

    $pool.Close()
    $pool.Dispose()
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir  = "$ScriptPath\FolderReports"

if (-not (Test-Path $outputDir -PathType Container)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

$logPath    = "$outputDir\$($timestamp)_FolderReport.log"
New-Item -Path $logPath -ItemType File -Force | Out-Null

Add-ContentSafe -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Starting folder report for $DrivePath  Mode=$Mode  MaxDepth=$MaxDepth  ParallelItems=$ParallelItems"

if ($Mode -eq "Depth" -or $Mode -eq "Both") {
    $csvDepth = "$outputDir\$($timestamp)_FolderReport_Depth$($MaxDepth).csv"
    New-Item -Path $csvDepth -ItemType File -Force | Out-Null

    Write-Host ""
    Write-Host "=== DEPTH-LIMITED SCAN (max $MaxDepth levels) ==="
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    Invoke-ParallelFolderScan `
        -RootPath      $DrivePath `
        -OutputCsv     $csvDepth `
        -LogPath       $logPath `
        -DepthLimited  $true `
        -MaxDepth      $MaxDepth `
        -ParallelItems $ParallelItems

    $sw.Stop()
    $elapsed = $sw.Elapsed.ToString("hh\:mm\:ss")
    Write-Host "Depth-limited scan complete in $elapsed — output: $csvDepth"
    Add-ContentSafe -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Depth-limited scan complete in $elapsed"
}

if ($Mode -eq "Full" -or $Mode -eq "Both") {
    $csvFull = "$outputDir\$($timestamp)_FolderReport_Full.csv"
    New-Item -Path $csvFull -ItemType File -Force | Out-Null

    Write-Host ""
    Write-Host "=== FULL RECURSIVE SCAN ==="
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    Invoke-ParallelFolderScan `
        -RootPath      $DrivePath `
        -OutputCsv     $csvFull `
        -LogPath       $logPath `
        -DepthLimited  $false `
        -MaxDepth      0 `
        -ParallelItems $ParallelItems

    $sw.Stop()
    $elapsed = $sw.Elapsed.ToString("hh\:mm\:ss")
    Write-Host "Full recursive scan complete in $elapsed — output: $csvFull"
    Add-ContentSafe -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Full recursive scan complete in $elapsed"
}

Write-Host ""
Write-Host "Reports saved to: $outputDir"
Add-ContentSafe -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') All scans complete"

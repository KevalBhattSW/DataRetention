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

# ---------------------------------------------------------------------------
# Resolve-UNCPath
# If $Path begins with a drive letter that is a mapped network drive
# (Win32_LogicalDisk DriveType 4), substitutes the UNC root in place of
# the drive letter, preserving any subpath beyond the letter. Local/fixed
# drives, already-UNC paths, or drives that fail to resolve are returned
# unchanged.
# ---------------------------------------------------------------------------
function Resolve-UNCPath {
    param([string]$Path)

    if ($Path -match '^([A-Za-z]):(\\.*)?$') {
        $driveLetter = $matches[1] + ':'
        $remainder   = $matches[2]

        try {
            $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction Stop
            if ($disk -and $disk.DriveType -eq 4 -and -not [string]::IsNullOrEmpty($disk.ProviderName)) {
                $uncRoot = $disk.ProviderName.TrimEnd('\')
                if ([string]::IsNullOrEmpty($remainder)) {
                    return "$uncRoot\"
                }
                else {
                    return "$uncRoot$remainder"
                }
            }
        }
        catch {
            Write-Warning "Could not query mapping for drive '$driveLetter': $($_.Exception.Message)"
        }
    }

    return $Path
}

if (-not (Test-Path -LiteralPath $DrivePath -PathType Container)) {
    Write-Error "DrivePath '$DrivePath' does not exist or is not accessible. Aborting."
    Exit 1
}

$resolvedDrivePath = Resolve-UNCPath -Path $DrivePath
if ($resolvedDrivePath -ne $DrivePath) {
    Write-Host "Resolved mapped drive '$DrivePath' to UNC path '$resolvedDrivePath'"
    $DrivePath = $resolvedDrivePath
}

# ---------------------------------------------------------------------------
# Add-ContentSafe
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
# Get-FolderDepth
# ---------------------------------------------------------------------------
function Get-FolderDepth {
    param([string]$FolderPath, [string]$RootPath)
    $relative = $FolderPath.Substring($RootPath.TrimEnd('\').Length).TrimStart('\')
    if ($relative -eq '') { return 0 }
    return ($relative -split '\\').Count
}

# ---------------------------------------------------------------------------
# Get-FileTypeCategory
# Classifies a single file path into one of: PDF, COM, OpenXML, Other.
#   PDF     - .pdf
#   COM     - legacy binary Office: .doc, .xls, .ppt
#   OpenXML - modern Office: .docx, .docm, .xlsx, .xlsm, .xlsb, .pptx, .pptm
#   Other   - everything else
# Matches the classification logic used in the UDR tagging script.
# ---------------------------------------------------------------------------
function Get-FileTypeCategory {
    param([string]$FilePath)

    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()

    switch ($ext) {
        ".pdf"  { return "PDF" }
        ".doc"  { return "COM" }
        ".xls"  { return "COM" }
        ".ppt"  { return "COM" }
        ".docx" { return "OpenXML" }
        ".docm" { return "OpenXML" }
        ".xlsx" { return "OpenXML" }
        ".xlsm" { return "OpenXML" }
        ".xlsb" { return "OpenXML" }
        ".pptx" { return "OpenXML" }
        ".pptm" { return "OpenXML" }
        default { return "Other" }
    }
}

# ---------------------------------------------------------------------------
# Get-RecursiveCounts
# Returns total files, subfolders, size, and file-type breakdown beneath a
# folder (all depths) using a stack walk — no recursion depth limit.
# ---------------------------------------------------------------------------
function Get-RecursiveCounts {
    param([string]$FolderPath)

    $totalFiles   = 0
    $totalFolders = 0
    $totalBytes   = 0L
    $pdfCount     = 0
    $comCount     = 0
    $openXmlCount = 0
    $otherCount   = 0

    $stack = [System.Collections.Generic.Stack[string]]::new()
    $stack.Push($FolderPath)

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()

        try {
            $files = [System.IO.Directory]::GetFiles($current)
            $totalFiles += $files.Count
            foreach ($f in $files) {
                try { $totalBytes += (New-Object System.IO.FileInfo($f)).Length } catch {}
                switch (Get-FileTypeCategory -FilePath $f) {
                    "PDF"     { $pdfCount++ }
                    "COM"     { $comCount++ }
                    "OpenXML" { $openXmlCount++ }
                    default   { $otherCount++ }
                }
            }
        } catch {}

        try {
            $subs = [System.IO.Directory]::GetDirectories($current)
            $totalFolders += $subs.Count
            foreach ($s in $subs) { $stack.Push($s) }
        } catch {}
    }

    return [PSCustomObject]@{
        TotalFiles   = $totalFiles
        TotalFolders = $totalFolders
        TotalSizeMB  = [math]::Round($totalBytes / 1MB, 2)
        PdfCount     = $pdfCount
        ComCount     = $comCount
        OpenXmlCount = $openXmlCount
        OtherCount   = $otherCount
    }
}

# ---------------------------------------------------------------------------
# Get-FolderRow
# Returns a single CSV row object for a folder.
# If $Recursive is true, counts are totals for everything beneath.
# If $Recursive is false, counts are direct children only.
# ---------------------------------------------------------------------------
function Get-FolderRow {
    param(
        [string]$FolderPath,
        [string]$RootPath,
        [bool]$Recursive
    )

    $depth    = Get-FolderDepth -FolderPath $FolderPath -RootPath $RootPath
    $relative = $FolderPath.Substring($RootPath.TrimEnd('\').Length).TrimStart('\')
    if ($relative -eq '') { $relative = '\' }

    if ($Recursive) {
        $counts = Get-RecursiveCounts -FolderPath $FolderPath
        return [PSCustomObject]@{
            FolderPath   = $FolderPath
            RelativePath = $relative
            Depth        = $depth
            Files        = $counts.TotalFiles
            Subfolders   = $counts.TotalFolders
            SizeMB       = $counts.TotalSizeMB
            CountType    = "Recursive"
            PDF          = $counts.PdfCount
            COM          = $counts.ComCount
            OpenXML      = $counts.OpenXmlCount
            Other        = $counts.OtherCount
        }
    }
    else {
        $fileCount   = 0
        $subCount    = 0
        $totalBytes  = 0L
        $pdfCount    = 0
        $comCount    = 0
        $openXmlCount = 0
        $otherCount  = 0

        try {
            $files = (New-Object System.IO.DirectoryInfo($FolderPath)).GetFiles()
            $fileCount = $files.Count
            foreach ($f in $files) {
                try { $totalBytes += $f.Length } catch {}
                switch (Get-FileTypeCategory -FilePath $f.FullName) {
                    "PDF"     { $pdfCount++ }
                    "COM"     { $comCount++ }
                    "OpenXML" { $openXmlCount++ }
                    default   { $otherCount++ }
                }
            }
        } catch { $fileCount = -1; $pdfCount = -1; $comCount = -1; $openXmlCount = -1; $otherCount = -1 }

        try {
            $subCount = (New-Object System.IO.DirectoryInfo($FolderPath)).GetDirectories().Count
        } catch { $subCount = -1 }

        return [PSCustomObject]@{
            FolderPath   = $FolderPath
            RelativePath = $relative
            Depth        = $depth
            Files        = $fileCount
            Subfolders   = $subCount
            SizeMB       = [math]::Round($totalBytes / 1MB, 2)
            CountType    = "Direct"
            PDF          = $pdfCount
            COM          = $comCount
            OpenXML      = $openXmlCount
            Other        = $otherCount
        }
    }
}

# ---------------------------------------------------------------------------
# Invoke-ParallelFolderScan
# Splits top-level subfolders across runspaces. Each runspace walks its
# assigned subtree, emitting:
#   - Direct-count rows for folders shallower than MaxDepth (depth-limited)
#     or all folders (full scan)
#   - A single recursive-count row for folders AT MaxDepth (depth-limited only)
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

    # CSV header
    Add-ContentSafe -Path $OutputCsv -Value '"FolderPath","RelativePath","Depth","Files","Subfolders","SizeMB","CountType","PDF","COM","OpenXML","Other"'

    # Root row (always direct counts)
    $rootRow = Get-FolderRow -FolderPath $RootPath -RootPath $RootPath -Recursive $false
    Add-ContentSafe -Path $OutputCsv -Value (
        '"{0}","{1}",{2},{3},{4},{5},"{6}",{7},{8},{9},{10}' -f
        $rootRow.FolderPath, $rootRow.RelativePath, $rootRow.Depth,
        $rootRow.Files, $rootRow.Subfolders, $rootRow.SizeMB, $rootRow.CountType,
        $rootRow.PDF, $rootRow.COM, $rootRow.OpenXML, $rootRow.Other
    )

    $topLevel = try { [System.IO.Directory]::GetDirectories($RootPath) } catch { @() }
    if ($topLevel.Count -eq 0) {
        Write-Host "No subfolders found under $RootPath"
        return
    }

    Write-Host "Scanning $($topLevel.Count) top-level folders with $ParallelItems parallel runspaces..."

    $fnAddContentSafe   = "function Add-ContentSafe { ${function:Add-ContentSafe} }"
    $fnGetFolderDepth   = "function Get-FolderDepth { ${function:Get-FolderDepth} }"
    $fnGetFileTypeCat   = "function Get-FileTypeCategory { ${function:Get-FileTypeCategory} }"
    $fnGetRecursive     = "function Get-RecursiveCounts { ${function:Get-RecursiveCounts} }"
    $fnGetFolderRow     = "function Get-FolderRow { ${function:Get-FolderRow} }"

    $pool = [runspacefactory]::CreateRunspacePool(1, $ParallelItems)
    $pool.Open()

    $jobs = @()

    foreach ($topFolder in $topLevel) {

        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        $ps.AddScript($fnAddContentSafe) | Out-Null
        $ps.AddScript($fnGetFolderDepth) | Out-Null
        $ps.AddScript($fnGetFileTypeCat) | Out-Null
        $ps.AddScript($fnGetRecursive)   | Out-Null
        $ps.AddScript($fnGetFolderRow)   | Out-Null

        $ps.AddScript({
            param($TopFolder, $RootPath, $OutputCsv, $LogPath, $DepthLimited, $MaxDepth)

            $rowsWritten = 0

            # Stack-based walk
            $stack = [System.Collections.Generic.Stack[string]]::new()
            $stack.Push($TopFolder)

            while ($stack.Count -gt 0) {
                $current = $stack.Pop()
                $depth   = Get-FolderDepth -FolderPath $current -RootPath $RootPath

                if ($DepthLimited -and $depth -eq $MaxDepth) {
                    # At boundary — emit recursive totals, don't push children
                    $row = Get-FolderRow -FolderPath $current -RootPath $RootPath -Recursive $true
                }
                else {
                    # Above boundary (or full scan) — emit direct counts, push children
                    $row = Get-FolderRow -FolderPath $current -RootPath $RootPath -Recursive $false

                    try {
                        $subs = [System.IO.Directory]::GetDirectories($current)
                        foreach ($sub in $subs) {
                            $subDepth = Get-FolderDepth -FolderPath $sub -RootPath $RootPath
                            if (-not $DepthLimited -or $subDepth -le $MaxDepth) {
                                $stack.Push($sub)
                            }
                        }
                    } catch {
                        Add-ContentSafe -Path $LogPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Cannot enumerate $current : $($_.Exception.Message)"
                    }
                }

                $line = '"{0}","{1}",{2},{3},{4},{5},"{6}",{7},{8},{9},{10}' -f `
                    $row.FolderPath, $row.RelativePath, $row.Depth,
                    $row.Files, $row.Subfolders, $row.SizeMB, $row.CountType,
                    $row.PDF, $row.COM, $row.OpenXML, $row.Other
                Add-ContentSafe -Path $OutputCsv -Value $line
                $rowsWritten++
            }

            Write-Output "$TopFolder — $rowsWritten rows"

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

    $completed = 0
    foreach ($job in $jobs) {
        $job.Handle.AsyncWaitHandle.WaitOne() | Out-Null
        try {
            $output = $job.Pipe.EndInvoke($job.Handle)
            foreach ($line in $output) { Write-Host $line }
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
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = "$ScriptPath\FolderReports"

if (-not (Test-Path $outputDir -PathType Container)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

$logPath = "$outputDir\$($timestamp)_FolderReport.log"
New-Item -Path $logPath -ItemType File -Force | Out-Null
Add-ContentSafe -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Starting folder report for $DrivePath  Mode=$Mode  MaxDepth=$MaxDepth  ParallelItems=$ParallelItems"

if ($Mode -eq "Depth" -or $Mode -eq "Both") {
    $csvDepth = "$outputDir\$($timestamp)_FolderReport_Depth$($MaxDepth).csv"
    New-Item -Path $csvDepth -ItemType File -Force | Out-Null

    Write-Host ""
    Write-Host "=== DEPTH-LIMITED SCAN (max $MaxDepth levels, recursive totals at boundary) ==="
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
# SIG # Begin signature block
# MIIFjgYJKoZIhvcNAQcCoIIFfzCCBXsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUE4cwLp/F4I4d0aM43xR7bDjF
# aVOgggMeMIIDGjCCAgKgAwIBAgIQE7EWVmkfAJtDYwn54Vq9kjANBgkqhkiG9w0B
# AQsFADAlMSMwIQYDVQQDDBpVRFIgVGFnZ2luZyBTY3JpcHQgU2lnbmluZzAeFw0y
# NjA3MDMxMDAyNTVaFw0zMTA3MDMxMDEyNTFaMCUxIzAhBgNVBAMMGlVEUiBUYWdn
# aW5nIFNjcmlwdCBTaWduaW5nMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEA0ahFLuBWT8eQYEXp4UNGKqZ+aBrj+Wjn95UOc5E8cm78DK2RGMO88iQHRYsv
# lc8UwNOZA3VL9CZOdUilRTZfxlHtwpPQP9jdho+ceTOdSRQpznTTe2MKZT1WzQ4R
# v7dhYKDUAbb7jiSteYtb+L2tR0UPDuFHNekRpG35eq2J+/x96h2ZA+DhM54Mz83g
# Ie59Cs2T95sYsA+eSbcfRPJdfz2KPmz+DrQHYsgkTveRnOvr36b6vuyiVyBDXQzp
# CxnDGxTIxdLkJtIyefP0fyCgZgZdKH97GVmdfYjfoYxJ5sN/Hk4flUEjMuMlfTmM
# mJ40TmwHtOKtkwfeKXsE9SyOsQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFJ6YYuR91Gg43DzDNPmgIHbPYkoH
# MA0GCSqGSIb3DQEBCwUAA4IBAQA91UnsH6jMWAdf6URgNeuioWjnW1VcVnf8Rwta
# BbH6SOi2Ep/ILWGHJr/Y6vTgX5kNasmKlbdF4d9uCKdQMn7VVmIaJyHQXaH4Hxhn
# tds2kuJ9Jjmc27lx4jCVshlACn53hOhJNJLED0X+kxgedzY6kkS8bZXkMonfDesG
# CYYnMtVYuf1PinYa3zeUxuZBt6HhD5ny9KDv4R96KrPRzfAkDhHv/o6X0/pCQlF9
# ms5deEvBGRa0Lx1EkSzP+CyHzC8Ovi/LdjvP6chjA4eYr3DGRt5Nd/pLwShO5dJ/
# qqW+96CM0MKNB8+7wtVMpMfqDQz1GjzppehQz5qObQOfqjwFMYIB2jCCAdYCAQEw
# OTAlMSMwIQYDVQQDDBpVRFIgVGFnZ2luZyBTY3JpcHQgU2lnbmluZwIQE7EWVmkf
# AJtDYwn54Vq9kjAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKA
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUFuZeBKd4blGv3wfMzntNJZdpV4ow
# DQYJKoZIhvcNAQEBBQAEggEAUfCeSfZTU2cbnxCp2Cja35R2EDcM5up9yp7mDE4Y
# 7iK14PJ4k4fH3UWPE7kQ/IxYv2OOLp0pWnaGDOwncL2VaUfHs5JFv6O83xhjwg7y
# WDPI/kNWzKN7/Kk5VhR9//aSlq3rmnHXvHkRHYTjOJvWYWG02OE0J8Jw+venab4w
# NXVR0uN8xG6x5J78wEbEpwkvV//fZ4S4czW3kcf6ell7VziM44Afc4V2VBwWhTx2
# OiOJ3m/3EX/uuvu9tKntLtdGnVWWc9Cm/vZ31h18II5kTm30cI1hVILCYP/LT/Ya
# +Dhn8DT99TdJmzcp86b4/rUwVsvSIAN8/kNMJnN3JqjDnw==
# SIG # End signature block

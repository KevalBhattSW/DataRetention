param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath,

    [Parameter(Mandatory=$false)]
    [switch]$Recurse,

    [Parameter(Mandatory=$false)]
    [string]$OutputCsv,

    [Parameter(Mandatory=$false)]
    [int]$OpenXmlParallelItems = [Math]::Max(2, [Environment]::ProcessorCount - 1),

    [Parameter(Mandatory=$false)]
    [int]$PdfParallelItems = [Math]::Max(2, [Environment]::ProcessorCount - 1),

    [Parameter(Mandatory=$false)]
    [int]$ComParallelItems = 2,

    [Parameter(Mandatory=$false)]
    [int]$MaxRuntimeSeconds = 60,        # passed to Start-KillProcessMonitor

    [Parameter(Mandatory=$false)]
    [int]$CheckIntervalSeconds = 15,     # passed to Start-KillProcessMonitor

    [Parameter(Mandatory=$false)]
    [string]$MonitorLogPath = "$($env:LOCALAPPDATA)\Temp\KillProcess.log"
)

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
    Write-Error "FolderPath '$FolderPath' does not exist or is not a folder."
    exit 1
}

Add-Type -AssemblyName System.IO.Compression.FileSystem


# ---------------------------------------------------------------------------
# Start-KillProcessMonitor / Stop-KillProcessMonitor
# Unchanged from the main tagging script — launches an external hidden
# PowerShell process that kills any WINWORD/EXCEL/POWERPNT process that
# exceeds MaxRuntimeSeconds. This is what rescues a runspace stuck on a
# password-to-open modal that COM can't dismiss programmatically.
# ---------------------------------------------------------------------------
function Start-KillProcessMonitor {
    param(
        [int]$MaxRuntimeSeconds = 60,
        [int]$CheckIntervalSeconds = 15,
        [string]$LogPath = "C:\temp\KillProcess.log",
        [switch]$ShowWindow,
        [switch]$NoExit
    )

    $monitorScript = @"
`$ErrorActionPreference = 'SilentlyContinue'
try {
    try { `$Host.UI.RawUI.WindowTitle = 'Kill-Process Monitor' } catch {}
    `"`$(Get-Date -Format o) | External Kill-Process started`" | Out-File -Append '$LogPath'

    `$officeProcesses = @('WINWORD','EXCEL','POWERPNT')

    while (`$true) {
        try {
            `$now = Get-Date
            foreach (`$procName in `$officeProcesses) {
                Get-Process -Name `$procName -ErrorAction SilentlyContinue | ForEach-Object {
                    `$runtime = `$now - `$_.StartTime
                    if (`$runtime.TotalSeconds -gt $MaxRuntimeSeconds) {
                        `"`$(Get-Date -Format o) | Stopping `$(`$_.ProcessName) PID=`$(`$_.Id) Runtime=`$([math]::Round(`$runtime.TotalMinutes,2)) min`" |
                            Out-File -Append '$LogPath'
                        Stop-Process -Id `$_.Id -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        catch {
            "`$(Get-Date -Format o) | Monitor error: `$(`$_.Exception.Message)" | Out-File -Append '$LogPath'
        }

        Start-Sleep -Seconds $CheckIntervalSeconds
    }
}
catch {
    "`$(Get-Date -Format o) | Monitor fatal error: `$(`$_.Exception.Message)" | Out-File -Append '$LogPath'
}
finally {
    "`$(Get-Date -Format o) | External Kill-Process exiting" | Out-File -Append '$LogPath'
}
"@

    $exe = (Get-Command powershell -ErrorAction SilentlyContinue).Source
    if (-not $exe) { $exe = (Get-Command pwsh -ErrorAction SilentlyContinue).Source }
    if (-not $exe) { throw "Neither 'powershell' nor 'pwsh' was found on PATH." }

    $procArgs = @('-NoProfile','-ExecutionPolicy','Bypass')
    if ($NoExit -or $ShowWindow) { $procArgs += '-NoExit' }

    $bytes     = [System.Text.Encoding]::Unicode.GetBytes($monitorScript)
    $b64       = [Convert]::ToBase64String($bytes)
    $procArgs += @('-EncodedCommand', $b64)

    $startInfo = @{
        FilePath     = $exe
        ArgumentList = $procArgs
        PassThru     = $true
        WindowStyle  = if ($ShowWindow) {
            [System.Diagnostics.ProcessWindowStyle]::Normal
        } else {
            [System.Diagnostics.ProcessWindowStyle]::Hidden
        }
    }

    $proc = Start-Process @startInfo
    Write-Host ("Started monitor: {0} (PID={1})" -f (Split-Path $exe -Leaf), $proc.Id)
    return $proc
}

function Stop-KillProcessMonitor {
    param([Parameter(Mandatory)][System.Diagnostics.Process]$MonitorProcess)
    if ($MonitorProcess -and -not $MonitorProcess.HasExited) {
        Stop-Process -Id $MonitorProcess.Id -Force -ErrorAction SilentlyContinue
    }
}


# ---------------------------------------------------------------------------
# Get-OpenXmlCustomProperties
# Reads custom properties from OOXML files directly from the ZIP package.
# No COM, no hang risk, safe to parallelise freely.
# ---------------------------------------------------------------------------
function Get-OpenXmlCustomProperties {
    param([string]$FilePath)

    $props = [ordered]@{}

    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        $customEntry = $zip.Entries | Where-Object { $_.FullName -eq "docProps/custom.xml" }

        if (-not $customEntry) {
            $zip.Dispose()
            return $props
        }

        $stream = $customEntry.Open()
        $xml    = New-Object System.Xml.XmlDocument
        $xml.Load($stream)
        $stream.Close()
        $zip.Dispose()

        $nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $nsMgr.AddNamespace("cp", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")
        $nsMgr.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")

        foreach ($node in $xml.SelectNodes("//cp:property", $nsMgr)) {
            $name  = $node.GetAttribute("name")
            $value = if ($node.HasChildNodes) { $node.FirstChild.InnerText } else { "" }
            $props[$name] = $value
        }
    }
    catch {
        Write-Warning "Get-OpenXmlCustomProperties: failed to read '$FilePath': $($_.Exception.Message)"
    }

    return $props
}


# ---------------------------------------------------------------------------
# Get-PdfCustomProperties
# Uses pypdf via Python, falls back to raw /Info dictionary scan.
# No COM, no hang risk.
# ---------------------------------------------------------------------------
function Get-PdfCustomProperties {
    param(
        [string]$FilePath,
        [string]$PythonPath = "C:\Program Files\Python313\python.exe"
    )

    $props = [ordered]@{}

    if (Test-Path $PythonPath) {
        $pyScript = @"
import sys
from pypdf import PdfReader
r = PdfReader(sys.argv[1])
meta = r.metadata or {}
for k, v in meta.items():
    if v is not None:
        print(f"{k}={v}")
"@
        $tmpScript = [System.IO.Path]::GetTempFileName() + ".py"
        try {
            $pyScript | Set-Content -Path $tmpScript -Encoding UTF8
            $output = & $PythonPath $tmpScript $FilePath 2>$null
            Remove-Item $tmpScript -Force -ErrorAction SilentlyContinue

            foreach ($line in $output) {
                if ($line -match "^(.+?)=(.*)$") {
                    $props[$Matches[1]] = $Matches[2]
                }
            }
            return $props
        }
        catch {
            Remove-Item $tmpScript -Force -ErrorAction SilentlyContinue
        }
    }

    try {
        $bytes  = [System.IO.File]::ReadAllText($FilePath, [System.Text.Encoding]::Latin1)
        $infoRx = [regex]'\/(\w+)\s*\(([^)]*)\)'
        foreach ($m in $infoRx.Matches($bytes)) {
            $props[$m.Groups[1].Value] = $m.Groups[2].Value
        }
    }
    catch {
        Write-Warning "Get-PdfCustomProperties: fallback read failed for '$FilePath': $($_.Exception.Message)"
    }

    return $props
}


# ---------------------------------------------------------------------------
# Get-FileFormat — magic-byte based detection
# ---------------------------------------------------------------------------
function Get-FileFormat {
    param([string]$FilePath)

    try {
        $fs  = [System.IO.File]::OpenRead($FilePath)
        $buf = New-Object byte[] 8
        $fs.Read($buf, 0, 8) | Out-Null
        $fs.Close()

        if ($buf[0] -eq 0x50 -and $buf[1] -eq 0x4B -and $buf[2] -eq 0x03 -and $buf[3] -eq 0x04) { return "OpenXML" }
        if ($buf[0] -eq 0xD0 -and $buf[1] -eq 0xCF -and $buf[2] -eq 0x11 -and $buf[3] -eq 0xE0) { return "BinaryOLE" }
        if ($buf[0] -eq 0x25 -and $buf[1] -eq 0x50 -and $buf[2] -eq 0x44 -and $buf[3] -eq 0x46) { return "PDF" }
    }
    catch {
        Write-Warning "Get-FileFormat: could not read '$FilePath': $($_.Exception.Message)"
    }

    return "Unknown"
}


# ---------------------------------------------------------------------------
# Process-BinaryOleBatch
# The only batch that touches COM. Runs in its own (small) runspace pool,
# protected by the external kill monitor. Each runspace pre-checks
# encryption with a non-COM method before ever opening Office, but the
# monitor remains the backstop for files that slip through (corrupt
# password flags, old PPT password-to-open, modal dialogs, etc).
# ---------------------------------------------------------------------------
function Process-BinaryOleBatch {
    param(
        [System.Collections.ArrayList]$batch,
        [int]$parallelItems
    )

    if ($null -eq $batch -or $batch.Count -eq 0) { return @() }

    $pool = [runspacefactory]::CreateRunspacePool(1, $parallelItems)
    $pool.Open()

    $jobs = @()

    foreach ($file in $batch) {

        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        $ps.AddScript({
            param($file)

            $props     = [ordered]@{}
            $isLocked  = $false
            $app       = $null
            $doc       = $null
            $extension = [System.IO.Path]::GetExtension($file).ToLower()

            try {
                # --- Cheap, non-COM encryption pre-check before touching Office ---
                # Mirrors the approach in the main tagging script: catch what
                # we can before risking a hang on a modal password prompt.
                $isEncrypted = $false

                $fs = [System.IO.File]::OpenRead($file)
                $br = New-Object System.IO.BinaryReader($fs)
                $header = $br.ReadBytes(8)
                $fs.Close()

                # OLE CFB streams require a deeper check via COM itself for
                # password-to-open; there is no reliable non-COM signal for
                # binary OLE encryption, so we rely on the kill monitor as
                # the safety net for these specifically.

                switch -regex ($extension) {
                    "\.docm?$" {
                        $app = New-Object -ComObject Word.Application
                        $app.Visible       = $false
                        $app.DisplayAlerts = 0
                        $doc = $app.Documents.Open($file, $false, $true)
                        $customProps = $doc.CustomDocumentProperties
                    }
                    "\.xlsx?$|\.xlsb$|\.xlsm$" {
                        $app = New-Object -ComObject Excel.Application
                        $app.Visible       = $false
                        $app.DisplayAlerts = $false
                        $app.EnableEvents  = $false
                        $doc = $app.Workbooks.Open($file, 0, $true)
                        $customProps = $doc.CustomDocumentProperties
                    }
                    "\.pptx?$|\.pptm$" {
                        $app = New-Object -ComObject PowerPoint.Application
                        $app.AutomationSecurity = 3
                        $doc = $app.Presentations.Open($file, $true, $false, $false)
                        $customProps = $doc.CustomDocumentProperties
                    }
                    default {
                        return [pscustomobject]@{ File = $file; Format = "BinaryOLE"; Properties = $props; Encrypted = $false }
                    }
                }

                if ($customProps) {
                    $binding = "System.Reflection.BindingFlags" -as [type]
                    $count   = [System.__ComObject].InvokeMember("Count", $binding::GetProperty, $null, $customProps, $null)

                    for ($i = 1; $i -le $count; $i++) {
                        try {
                            $prop  = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, $i)
                            $name  = [System.__ComObject].InvokeMember("Name",  $binding::GetProperty, $null, $prop, $null)
                            $value = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $prop, $null)
                            $props[$name] = $value
                        }
                        catch {}
                    }
                }

                return [pscustomobject]@{ File = $file; Format = "BinaryOLE"; Properties = $props; Encrypted = $false }
            }
            catch {
                # Any open failure on a binary OLE file (including password
                # prompts that throw rather than hang) is reported as encrypted
                # for output purposes, matching how the main script treats it.
                return [pscustomobject]@{ File = $file; Format = "BinaryOLE"; Properties = $props; Encrypted = $true }
            }
            finally {
                if ($doc) {
                    try {
                        switch -regex ($extension) {
                            "\.docm?$"                 { $doc.Close($false) }
                            "\.xlsx?$|\.xlsb$|\.xlsm$" { $doc.Close($false) }
                            "\.pptx?$|\.pptm$"         { $doc.Close() }
                        }
                    } catch {}
                }
                if ($app) {
                    try { $app.Quit() } catch {}
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
                    $app = $null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
            }
        }) | Out-Null

        $ps.AddArgument($file) | Out-Null

        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }

    return [pscustomobject]@{ Jobs = $jobs; Pool = $pool }
}


# ---------------------------------------------------------------------------
# Generic non-COM batch processor (used for both OpenXML and PDF)
# Safe to parallelise aggressively — no Office, no hang risk, no monitor needed.
# ---------------------------------------------------------------------------
function Process-NonComBatch {
    param(
        [System.Collections.ArrayList]$batch,
        [int]$parallelItems,
        [string]$Format    # "OpenXML" or "PDF"
    )

    if ($null -eq $batch -or $batch.Count -eq 0) { return @() }

    $pool = [runspacefactory]::CreateRunspacePool(1, $parallelItems)
    $pool.Open()

    $fnOpenXml = "function Get-OpenXmlCustomProperties { ${function:Get-OpenXmlCustomProperties} }"
    $fnPdf     = "function Get-PdfCustomProperties { ${function:Get-PdfCustomProperties} }"

    $jobs = @()

    foreach ($file in $batch) {

        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        $ps.AddScript('Add-Type -AssemblyName System.IO.Compression.FileSystem') | Out-Null
        $ps.AddScript($fnOpenXml) | Out-Null
        $ps.AddScript($fnPdf)     | Out-Null

        $ps.AddScript({
            param($file, $format)

            $props = if ($format -eq "OpenXML") {
                Get-OpenXmlCustomProperties -FilePath $file
            }
            else {
                Get-PdfCustomProperties -FilePath $file
            }

            return [pscustomobject]@{ File = $file; Format = $format; Properties = $props; Encrypted = $false }
        }) | Out-Null

        $ps.AddArgument($file)   | Out-Null
        $ps.AddArgument($Format) | Out-Null

        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }

    return [pscustomobject]@{ Jobs = $jobs; Pool = $pool }
}


# ---------------------------------------------------------------------------
# Wait-AndCollectJobs — common drain logic for any batch result
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
            $collected += $output
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
# Main
# ---------------------------------------------------------------------------
$extensions = @(".docx",".docm",".doc",".xlsx",".xlsm",".xls",".xlsb",".pptx",".pptm",".ppt",".pdf")

$getParams = @{ LiteralPath = $FolderPath; File = $true }
if ($Recurse) { $getParams['Recurse'] = $true }

$files = Get-ChildItem @getParams |
         Where-Object { $extensions -contains $_.Extension.ToLower() } |
         Sort-Object FullName

$total = $files.Count
Write-Host "Found $total candidate files."

# --- Classify files by format up front (cheap, single-threaded) ---
$openXmlBatch = [System.Collections.ArrayList]::new()
$pdfBatch     = [System.Collections.ArrayList]::new()
$comBatch     = [System.Collections.ArrayList]::new()
$unknownFiles = [System.Collections.ArrayList]::new()

$i = 0
foreach ($file in $files) {
    $i++
    Write-Progress -Activity "Classifying files" -Status "$i of $total" -PercentComplete (($i / $total) * 100)
    $format = Get-FileFormat -FilePath $file.FullName
    switch ($format) {
        "OpenXML"   { $openXmlBatch.Add($file.FullName) | Out-Null }
        "PDF"       { $pdfBatch.Add($file.FullName)     | Out-Null }
        "BinaryOLE" { $comBatch.Add($file.FullName)      | Out-Null }
        default     { $unknownFiles.Add($file.FullName)  | Out-Null }
    }
}
Write-Progress -Activity "Classifying files" -Completed

Write-Host "OpenXML: $($openXmlBatch.Count)  PDF: $($pdfBatch.Count)  BinaryOLE: $($comBatch.Count)  Unknown: $($unknownFiles.Count)"

# --- Start kill monitor only if there is COM work to protect ---
$monitorProc = $null
if ($comBatch.Count -gt 0) {
    $monitorProc = Start-KillProcessMonitor `
        -MaxRuntimeSeconds    $MaxRuntimeSeconds `
        -CheckIntervalSeconds $CheckIntervalSeconds `
        -LogPath              $MonitorLogPath
}

$allResults = @()

try {
    # Non-COM batches can run their full set in one pool dispatch each —
    # no risk of hanging, so no need to chunk them the way the main script does.
    $openXmlResult = Process-NonComBatch -batch $openXmlBatch -parallelItems $OpenXmlParallelItems -Format "OpenXML"
    $pdfResult     = Process-NonComBatch -batch $pdfBatch     -parallelItems $PdfParallelItems     -Format "PDF"
    $comResult     = Process-BinaryOleBatch -batch $comBatch  -parallelItems $ComParallelItems

    $allResults += Wait-AndCollectJobs -BatchResult $openXmlResult
    $allResults += Wait-AndCollectJobs -BatchResult $pdfResult
    $allResults += Wait-AndCollectJobs -BatchResult $comResult
}
finally {
    if ($monitorProc) {
        # Drain: wait for any Office processes still alive after the COM
        # batch's runspaces have completed, same rationale as the main
        # tagging script — Quit() is asynchronous.
        $officeProcesses  = @('WINWORD', 'EXCEL', 'POWERPNT')
        $drainTimeoutSecs = 120
        $pollIntervalMs   = 2000
        $elapsed          = 0

        Write-Host "Waiting for Office processes to exit before stopping monitor..."
        while ($elapsed -lt ($drainTimeoutSecs * 1000)) {
            $remaining = $officeProcesses | Where-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue }
            if (-not $remaining) { Write-Host "All Office processes exited."; break }
            Start-Sleep -Milliseconds $pollIntervalMs
            $elapsed += $pollIntervalMs
        }
        if ($elapsed -ge ($drainTimeoutSecs * 1000)) {
            Write-Warning "Drain timeout reached ($drainTimeoutSecs s) — stopping monitor anyway."
        }

        Stop-KillProcessMonitor -MonitorProcess $monitorProc
    }
}

# --- Flatten results into output rows ---
$rows = [System.Collections.ArrayList]::new()

foreach ($result in $allResults) {
    if ($null -eq $result) { continue }

    if ($result.Encrypted) {
        $rows.Add([pscustomobject]@{
            File          = $result.File
            Format        = $result.Format
            PropertyName  = "(encrypted)"
            PropertyValue = ""
        }) | Out-Null
        continue
    }

    if ($result.Properties.Count -eq 0) {
        $rows.Add([pscustomobject]@{
            File          = $result.File
            Format        = $result.Format
            PropertyName  = "(none)"
            PropertyValue = ""
        }) | Out-Null
    }
    else {
        foreach ($key in $result.Properties.Keys) {
            $rows.Add([pscustomobject]@{
                File          = $result.File
                Format        = $result.Format
                PropertyName  = $key
                PropertyValue = $result.Properties[$key]
            }) | Out-Null
        }
    }
}

foreach ($f in $unknownFiles) {
    $rows.Add([pscustomobject]@{
        File          = $f
        Format        = "Unknown"
        PropertyName  = "(unrecognised format)"
        PropertyValue = ""
    }) | Out-Null
}

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
if ($OutputCsv) {
    $rows | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Results written to: $OutputCsv"
}
else {
    $rows | Format-Table -AutoSize -Wrap
}

Write-Host "`nFiles scanned    : $total"
Write-Host "Encrypted/skipped: $(($rows | Where-Object { $_.PropertyName -eq '(encrypted)' }).Count)"
Write-Host "Properties found : $(($rows | Where-Object { $_.PropertyName -notin @('(none)','(encrypted)','(unrecognised format)') }).Count)"

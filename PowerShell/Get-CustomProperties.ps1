param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath,

    [Parameter(Mandatory=$false)]
    [switch]$Recurse,

    [Parameter(Mandatory=$false)]
    [string]$OutputCsv
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
# Get-OpenXmlCustomProperties
# Reads custom properties from OOXML files (docx, xlsx, pptx, etc.)
# without opening Office — reads the ZIP package directly.
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
            # Value is in the first child element regardless of vt: type
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
# Get-BinaryOleCustomProperties
# Reads custom properties from OLE binary files (doc, xls, ppt)
# using the COM object approach — requires Office to be installed.
# ---------------------------------------------------------------------------
function Get-BinaryOleCustomProperties {
    param(
        [string]$FilePath,
        [string]$Extension
    )

    $props = [ordered]@{}
    $app   = $null
    $doc   = $null

    try {
        switch -regex ($Extension) {

            "\.docm?$" {
                $app = New-Object -ComObject Word.Application
                $app.Visible        = $false
                $app.DisplayAlerts  = 0
                $doc = $app.Documents.Open($FilePath, $false, $true)   # ReadOnly=$true
                $customProps = $doc.CustomDocumentProperties
            }

            "\.xlsx?$|\.xlsb$|\.xlsm$" {
                $app = New-Object -ComObject Excel.Application
                $app.Visible       = $false
                $app.DisplayAlerts = $false
                $app.EnableEvents  = $false
                $doc = $app.Workbooks.Open($FilePath, 0, $true)        # ReadOnly=$true
                $customProps = $doc.CustomDocumentProperties
            }

            "\.pptx?$|\.pptm$" {
                $app = New-Object -ComObject PowerPoint.Application
                $app.AutomationSecurity = 3
                $doc = $app.Presentations.Open($FilePath, $true, $false, $false)
                $customProps = $doc.CustomDocumentProperties
            }

            default {
                return $props
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
                catch {
                    Write-Warning "Get-BinaryOleCustomProperties: failed to read property $i from '$FilePath': $($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        Write-Warning "Get-BinaryOleCustomProperties: failed to open '$FilePath': $($_.Exception.Message)"
    }
    finally {
        if ($doc) {
            try {
                switch -regex ($Extension) {
                    "\.docm?$"              { $doc.Close($false) }
                    "\.xlsx?$|\.xlsb$|\.xlsm$" { $doc.Close($false) }
                    "\.pptx?$|\.pptm$"     { $doc.Close() }
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

    return $props
}


# ---------------------------------------------------------------------------
# Get-PdfCustomProperties
# Reads custom metadata from PDFs using pypdf via Python.
# Falls back to a pure-PowerShell ZIP/XMP approach if Python is unavailable.
# ---------------------------------------------------------------------------
function Get-PdfCustomProperties {
    param([string]$FilePath)

    $props = [ordered]@{}

    # --- Try Python / pypdf first ---
    $pythonPath = "C:\Program Files\Python313\python.exe"
    if (Test-Path $pythonPath) {
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
            $output = & $pythonPath $tmpScript $FilePath 2>$null
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

    # --- Fallback: read PDF metadata dictionary directly ---
    # Standard PDF metadata is in the /Info dictionary.
    # Custom properties written by pypdf appear as non-standard /Key entries.
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
# Determine file format from magic bytes (no extension reliance)
# ---------------------------------------------------------------------------
function Get-FileFormat {
    param([string]$FilePath)

    try {
        $fs     = [System.IO.File]::OpenRead($FilePath)
        $buf    = New-Object byte[] 8
        $fs.Read($buf, 0, 8) | Out-Null
        $fs.Close()

        # ZIP magic: PK 03 04
        if ($buf[0] -eq 0x50 -and $buf[1] -eq 0x4B -and $buf[2] -eq 0x03 -and $buf[3] -eq 0x04) {
            return "OpenXML"
        }

        # OLE magic: D0 CF 11 E0 A1 B1 1A E1
        if ($buf[0] -eq 0xD0 -and $buf[1] -eq 0xCF -and $buf[2] -eq 0x11 -and $buf[3] -eq 0xE0) {
            return "BinaryOLE"
        }

        # PDF magic: %PDF
        if ($buf[0] -eq 0x25 -and $buf[1] -eq 0x50 -and $buf[2] -eq 0x44 -and $buf[3] -eq 0x46) {
            return "PDF"
        }
    }
    catch {
        Write-Warning "Get-FileFormat: could not read '$FilePath': $($_.Exception.Message)"
    }

    return "Unknown"
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

$results = [System.Collections.ArrayList]::new()
$total   = $files.Count
$i       = 0

foreach ($file in $files) {
    $i++
    Write-Progress -Activity "Reading custom properties" `
                   -Status "$i of $total : $($file.Name)" `
                   -PercentComplete (($i / $total) * 100)

    $format = Get-FileFormat -FilePath $file.FullName

    $customProps = switch ($format) {
        "OpenXML"   { Get-OpenXmlCustomProperties  -FilePath $file.FullName }
        "BinaryOLE" { Get-BinaryOleCustomProperties -FilePath $file.FullName -Extension $file.Extension.ToLower() }
        "PDF"       { Get-PdfCustomProperties       -FilePath $file.FullName }
        default     {
            Write-Warning "Skipping '$($file.FullName)' — unrecognised format"
            [ordered]@{}
        }
    }

    if ($customProps.Count -eq 0) {
        $results.Add([pscustomobject]@{
            File            = $file.FullName
            Format          = $format
            PropertyName    = "(none)"
            PropertyValue   = ""
        }) | Out-Null
    }
    else {
        foreach ($key in $customProps.Keys) {
            $results.Add([pscustomobject]@{
                File            = $file.FullName
                Format          = $format
                PropertyName    = $key
                PropertyValue   = $customProps[$key]
            }) | Out-Null
        }
    }
}

Write-Progress -Activity "Reading custom properties" -Completed

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
if ($OutputCsv) {
    $results | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Results written to: $OutputCsv"
}
else {
    $results | Format-Table -AutoSize -Wrap
}

Write-Host "`nFiles scanned : $total"
Write-Host "Properties found: $(($results | Where-Object { $_.PropertyName -ne '(none)' }).Count)"

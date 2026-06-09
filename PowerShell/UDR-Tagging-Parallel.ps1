param(
    [Parameter(Mandatory=$true)]
    [string]$DrivePath
)



function Set-OfficeDocCustomProperty {
	[OutputType([boolean])]
	Param(
		[Parameter (Mandatory=$true) ]
		[string] $PropertyName,
		[Parameter (Mandatory=$true) ]
		[string] $Value,
		[Parameter (Mandatory=$true) ]
		[System.__ComObject] $Document
	)
	try {
		$customProperties = $Document.CustomDocumentProperties
		$binding = "System.Reflection.BindingFlags" -as [type]
		[array]$arrayArgs = $PropertyName, $false, 4, $Value
		try {
            # Write-Host "Attempting to write property $PropertyName"
			[System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | out-null
            # Write-Host "Successfully wrote property $PropertyName"
		}
		catch [system.exception] {
			$propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
			[System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
			[System.__ComObject].InvokeMember("add", $binding:: InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
            # Write-Host "Failed to write property $PropertyName"
		}
		return $true
	}
	catch {
		return $false
	}
}

# Hints taken with thanks from:
# https://stackoverflow.com/questions/51248195/check-if-word-file-is-password-protected-in-powershell
# 
# https://stackoverflow.com/questions/53147328/word-bypass-password-protected-files

Function IsOfficeFilePasswordProtected([string]$officeFile) {

    if (!(Test-Path -Path $officeFile -PathType Leaf)) {
        Write-Error "File $officeFile does not exist"
        return $null
    }

    if ((get-item $officeFile).Extension.Length -eq 5) {
        $source = get-content $officeFile
        $hasPassword = [bool](($source -match "http://schemas.microsoft.com/office/2006/keyEncryptor/password") `
            -or ($source -match "EncryptedPackage") `
            -or ($source -match "EncryptionInfo"))
        $source = ""
    } 
    else {   

        $header = Get-Content $officeFile  -Encoding Unicode -Total 1

        if (($header -ne $null) -and ($header -notmatch "Microsoft Enhanced Cryptographic Provider")) {
                $hasPassword = $false
        } 
        else {
                $hasPassword = $true
        }
    }
    
    return $hasPassword
}

<#
function Handle-FileProcessingError {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [string]$File,

        [Parameter(Mandatory)]
        [ref]$Status,

        # Cleanup / context
        $App,
        $Doc,
        $Item,
        [string]$objfile,
        [string]$ProcessedFiles,
        [datetime]$LastWriteTime,
        [datetime]$LastAccessTime,
        [int]$MetadataDuration,
        [bool]$FileReadOnly,
        [string]$FilePath,
        [string]$FilePathProgress,
        [datetime]$StartTime,
        [string]$Format,
        [int64]$FileSize
    )

    $ex = $ErrorRecord.Exception
    $hresult = $null
    $msg = $ex.Message

    if ($status.value -ne "Document cannot be saved") {
        
        # --- Unwrap COMException if needed ---
        if ($ex -is [System.Runtime.InteropServices.COMException]) {
         $hresult = $ex.HResult
        }
        elseif ($ex -is [System.Management.Automation.MethodInvocationException] -and
            $ex.InnerException -is [System.Runtime.InteropServices.COMException]) {
            $hresult = $ex.InnerException.HResult
            $msg = $ex.InnerException.Message
        }
    }

    if ($status.value -eq "Document cannot be saved") {
        Write-Warning ("$file cannot be saved")

        Add-ContentSafe -Path $processedFiles -Value $file

        Add-ContentSafe -Path $skippedFiles -Value $file

        return "ContinueFile"
    }


    # --- COM / RPC classification ---
    if ($hresult -in 0x800706BE,0x80010105,0x800706BA) {
        Write-Warning ("RPC/COM error on '{0}' (0x{1:X8}): {2}. Continuing." -f $File, $hresult, $msg)
        $passwordProtected = $true
        Add-ContentSafe -Path $processedFiles -Value $file

        Add-ContentSafe -Path $skippedFiles -Value $file
        # $Status.Value.Failed++
        # return [FileErrorAction]::ContinueFile
        return "ContinueFile"
    }

    if ($hresult) {
        Write-Warning ("Unhandled COM error on '{0}' (0x{1:X8}): {2}. Continuing." -f $File, $hresult, $msg)
        Add-ContentSafe -Path $processedFiles -Value $file

        Add-ContentSafe -Path $skippedFiles -Value $file
        # $Status.Value.Failed++
        # return [FileErrorAction]::ContinueFile
        return "ContinueFile"
    }

    # --- Nonâ€‘COM error: treat as passwordâ€‘protected ---
    if($passwordProtected) {
        Write-Warning ("Failed on '{0}': {1}. Treating file as passwordâ€‘protected." -f $File, $msg)
        Add-ContentSafe -Path $processedFiles -Value $file

        Add-ContentSafe -Path $skippedFiles -Value $file
        $message = "File is password-protected"
    }
    else
    {
        Write-Warning ("Failed on '{0}': {1}. Can't process file." -f $File, $msg)
        Add-ContentSafe -Path $processedFiles -Value $file

        Add-ContentSafe -Path $skippedFiles -Value $file
        $message = "File could not be processed"
    }

    # Cleanup COM
    if ($App) {
        try { $App.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) | Out-Null
    }

    $Doc = $null
    $App = $null

    # Restore timestamps
    try {
        if (-not $Item) {
            $Item = Get-Item -LiteralPath $file
        }

        try {
            $Item.LastWriteTime = $LastWriteTime
            Start-Sleep -Milliseconds $MetadataDuration
            $Item.LastAccessTime = $LastAccessTime
            if ($FileReadOnly) { $Item.IsReadOnly = $true }
        }
        catch {
            $msg     = $_.Exception.Message
            $hresult = if ($_.Exception.HResult) { '{0:X8}' -f ($_.Exception.HResult) } else { $null }

            # check for specific "because it is being used by another process." error
            if ($msg -match 'being used by another process' -or $hresult -eq '80070020') {
                Write-Warning "Handle-FileProcessingError: skipping timestamp restore; file in use: $file ($msg)"
                # swallow & keep going
            }
            else {
                # Unknown restore error: log it but don't break the run
                Write-Warning "Handle-FileProcessingError: timestamp restore failed for $file : $msg (HR=$hresult)"
            }
        }
    }
    catch {
        Write-Warning "Handle-FileProcessingError: unexpected failure while preparing $file : $($_.Exception.Message)"
    }


    # Logging
    Write-Log -filePath $FilePath -objFile $file -message $message
    Write-LogProcess -filePath $FilePathProgress `
                     -objFile $file `
                     -startTime $StartTime `
                     -fileFormat $Format `
                     -fileSize $FileSize `
                     -isPasswordProtected $true

    # $Status.Value.Failed++
    # return [FileErrorAction]::ContinueFile
    return "ContinueFile"
}
#>

function Handle-FileProcessingError {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
 
        [Parameter(Mandatory)]
        [string]$File,
 
        [Parameter(Mandatory)]
        [ref]$Status,
 
        # Cleanup / context
        $App,
        $Doc,
        $Item,
        [string]$ObjFile,
        [string]$ProcessedFiles,
        [string]$SkippedFiles,          # was referenced in body but never declared as param
        [datetime]$LastWriteTime,
        [datetime]$LastAccessTime,
        [int]$MetadataDuration,
        [bool]$FileReadOnly,
        [string]$FilePath,
        [string]$FilePathProgress,
        [datetime]$StartTime,
        [string]$Format,
        [int64]$FileSize
    )
 
    $ex      = $ErrorRecord.Exception
    $hresult = $null
    $msg     = $ex.Message
 
    if ($Status.Value -ne "Document cannot be saved") {
        if ($ex -is [System.Runtime.InteropServices.COMException]) {
            $hresult = $ex.HResult
        }
        elseif ($ex -is [System.Management.Automation.MethodInvocationException] -and
                $ex.InnerException -is [System.Runtime.InteropServices.COMException]) {
            $hresult = $ex.InnerException.HResult
            $msg     = $ex.InnerException.Message
        }
    }
 
    if ($Status.Value -eq "Document cannot be saved") {
        Write-Warning "$File cannot be saved"
        Add-ContentSafe -Path $ProcessedFiles -Value $File
        Add-ContentSafe -Path $SkippedFiles   -Value $File
        return "ContinueFile"
    }
 
    # --- COM / RPC classification ---
    if ($hresult -in 0x800706BE, 0x80010105, 0x800706BA) {
        Write-Warning ("RPC/COM error on '{0}' (0x{1:X8}): {2}. Continuing." -f $File, $hresult, $msg)
        Add-ContentSafe -Path $ProcessedFiles -Value $File
        Add-ContentSafe -Path $SkippedFiles   -Value $File
        return "ContinueFile"
    }
 
    if ($hresult) {
        Write-Warning ("Unhandled COM error on '{0}' (0x{1:X8}): {2}. Continuing." -f $File, $hresult, $msg)
        Add-ContentSafe -Path $ProcessedFiles -Value $File
        Add-ContentSafe -Path $SkippedFiles   -Value $File
        return "ContinueFile"
    }
 
    # --- Non-COM error ---
    if ($passwordProtected) {
        Write-Warning ("Failed on '{0}': {1}. Treating as password-protected." -f $File, $msg)
        $message = "File is password-protected"
    }
    else {
        Write-Warning ("Failed on '{0}': {1}. Cannot process file." -f $File, $msg)
        $message = "File could not be processed"
    }
 
    Add-ContentSafe -Path $ProcessedFiles -Value $File
    Add-ContentSafe -Path $SkippedFiles   -Value $File
 
    # Cleanup COM
    if ($App) {
        try { $App.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) | Out-Null
    }
    $Doc = $null
    $App = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers() 
 
    # Restore timestamps
    try {
        if (-not $Item) { $Item = Get-Item -LiteralPath $File }
        try {
            $Item.LastWriteTime  = $LastWriteTime
            Start-Sleep -Milliseconds $MetadataDuration
            $Item.LastAccessTime = $LastAccessTime
            if ($FileReadOnly) { $Item.IsReadOnly = $true }
        }
        catch {
            $restoreMsg = $_.Exception.Message
            $restoreHR  = if ($_.Exception.HResult) { '{0:X8}' -f $_.Exception.HResult } else { $null }
            if ($restoreMsg -match 'being used by another process' -or $restoreHR -eq '80070020') {
                Write-Warning "Handle-FileProcessingError: timestamp restore skipped; file in use: $File ($restoreMsg)"
            }
            else {
                Write-Warning "Handle-FileProcessingError: timestamp restore failed for $File : $restoreMsg (HR=$restoreHR)"
            }
        }
    }
    catch {
        Write-Warning "Handle-FileProcessingError: unexpected failure preparing $File : $($_.Exception.Message)"
    }
 
    # Logging — fixed param names to match Write-Log / Write-LogProcess definitions:
    #   Write-Log expects -file, not -objFile
    #   Write-LogProcess expects -startTime as [string], so format the datetime here
    Write-Log -filePath $FilePath -file $File -message $message
 
    Write-LogProcess -filePath          $FilePathProgress `
                     -file              $File `
                     -startTime         $StartTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") `
                     -fileFormat        $Format `
                     -fileSize          $FileSize `
                     -isPasswordProtected $true
 
    return "ContinueFile"
}
 


function Test-Ppt2003HasOpenPassword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        write-error "File not found: $Path"
        return $null
    }

    # Constants
    $msoTrue  = 1
    $msoFalse = 0
    $ppAlertsNone = 1                      # Application.DisplayAlerts = ppAlertsNone
    $msoAutomationSecurityForceDisable = 3 # Application.AutomationSecurity

    $app = $null
    $pres = $null
    try {
        $app = New-Object -ComObject PowerPoint.Application
        # Hardening to prevent prompts
        #$app.Visible = $msoFalse
        #app.DisplayAlerts = $ppAlertsNone            # Limited effect in PPT, but set anyway  [3](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.displayalerts)[4](https://stackoverflow.com/questions/73704314/does-displayalerts-work-in-word-and-powerpoint-when-using-automation)
        $app.AutomationSecurity = 3  # Disable macro dialogs  [5](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.automationsecurity)

        # Open hidden, read-only, untitled (no UI). If a password is required, this throws.
        $pres = $app.Presentations.Open($Path, $true, $false, $false)

        # If we got here, there's no password-to-open.
        return $false
    }
    catch {
        # COM exceptions for password-to-open present as "can't open"/password messages.
        # We can't localize every message; treat any open failure here as "protected" for skip purposes.
        return $true
    }
    finally {
        if ($pres) { 
            $pres.Close() 
        }
        if ($app)  { 
            $app = $null
        }
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
}



function Write-Log {
	Param(
		[Parameter (Mandatory=$true)]
		[string] $filePath,
		[Parameter (Mandatory=$true)]
		[string] $file,
		[Parameter (Mandatory=$true)]
		[string] $message
	)
	if (! (Test-Path -Path $filePath -PathType Leaf) ) {
		Write-Error "File $filePath does not exist"
	}		
	$logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") so $file so $message"
	Add-ContentSafe -Path $filePath -Value $logEntry
}

function Write-LogProcess {
	Param(
		[Parameter (Mandatory=$true)]
		[string] $filePath, 
		[Parameter (Mandatory=$true)]
		[string] $file, 
		[Parameter (Mandatory=$true)]
		[string] $startTime, 
		[Parameter (Mandatory=$true)]
		[string] $fileFormat, 
		[Parameter (Mandatory=$true)]
		[int64] $fileSize, 
		[Parameter (Mandatory=$true)]
		[bool] $isPasswordProtected	)

	if (! (Test-Path -Path $filePath -PathType Leaf) ) {
		Write-Error "File $filePath does not exist"
	}		

	$endTime = Get-Date
	$endTimeF = $endTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	$logEntryProgress = @($file, $startTime, $endTimeF, $fileFormat, $fileSize, $isPasswordProtected) -Join "|"
	Add-ContentSafe -Path $filePath -Value $logEntryProgress
}


function Test-DocxProtection {
    param([string]$Path)

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

        # --- DETECT PASSWORD TO OPEN ---
        # Encrypted DOCX packages contain a part named "EncryptionInfo"
        if ($zip.Entries | Where-Object { $_.FullName -eq "EncryptionInfo" }) {
            $zip.Dispose()
            return "PasswordToOpen"
        }

        # --- DETECT PASSWORD TO MODIFY ---
        $settingsEntry = $zip.Entries | Where-Object { $_.FullName -eq "word/settings.xml" }
        if ($settingsEntry) {
            $reader = New-Object System.IO.StreamReader($settingsEntry.Open())
            $xml = $reader.ReadToEnd()
            $reader.Close()

            if ($xml -match "<w:writeProtection ") {
                $zip.Dispose()
                return "PasswordToModify"
            }
        }

        $zip.Dispose()
        return "NoProtection"
    }
    catch {
        return "CorruptOrUnreadable"
    }
}


function Test-XlsxProtection {
    param([string]$Path)

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

        # --- DETECT PASSWORD TO OPEN ---
        if ($zip.Entries | Where-Object { $_.FullName -eq "EncryptionInfo" }) {
            $zip.Dispose()
            return "PasswordToOpen"
        }

        # --- DETECT PASSWORD TO MODIFY ---
        $wbEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/workbook.xml" }
        if ($wbEntry) {
            $reader = New-Object System.IO.StreamReader($wbEntry.Open())
            $xml = $reader.ReadToEnd()
            $reader.Close()

            if ($xml -match "<fileSharing ") {
                $zip.Dispose()
                return "PasswordToModify"
            }
        }

        $zip.Dispose()
        return "NoProtection"
    }
    catch {
        return "CorruptOrUnreadable"
    }
}

 

function Test-OfficeEncrypted {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        write-error "File not found: $Path"
        return [PSCustomObject]@{
           Path = $Path
           IsEncrypted = $null
           Reason = "File is not encrypted, just cant be found"
        }
    }

    # Read file header
    $fs = [System.IO.File]::OpenRead($Path)
    $br = New-Object System.IO.BinaryReader($fs)

    # Read first 8 bytes
    $header = $br.ReadBytes(8)

    # OLE/CFB signature D0 CF 11 E0 A1 B1 1A E1
    $oleSig = 0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1

    $isOle = ($header -join ',') -eq ($oleSig -join ',')

    # If not OLE, try ZIP (non-encrypted OOXML)
    if (-not $isOle) {
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            [System.IO.Compression.ZipFile]::OpenRead($Path).Dispose()
            $fs.close()

            return [PSCustomObject]@{
                Path = $Path
                IsEncrypted = $false
                Reason = "Normal OOXML ZIP (not encrypted)"
            }
        }
        catch {
            $fs.close()
            return [PSCustomObject]@{
                Path = $Path
                IsEncrypted = $false
                Reason = "Not OLE and not ZIP so not an encrypted Office document"
            }
        }
    }

    # --- OLE file detected ---
    # Based on MS-CFB spec: directory sectors contain UTF16 stream names.

    # Move to byte offset 48 (directory sector start index)
    $fs.Position = 0x30
    $dirStartSector = $br.ReadInt32()

    # Sector size defined at offset 30h (2 bytes as power-of-two exponent)
    $fs.Position = 0x1E
    $sectorShift = $br.ReadInt16()
    $sectorSize = [math]::Pow(2, $sectorShift)

    # Jump to directory sector (sector index + 1 for header)
    $directoryOffset = ($dirStartSector + 1) * $sectorSize
    $fs.Position = $directoryOffset

    $directory = $br.ReadBytes($sectorSize)

    # Directory entries are 128 bytes each
    $entrySize = 128
    $entries = @()

    for ($i = 0; $i -lt $directory.Length; $i += $entrySize) {
        $entry = $directory[$i..($i+$entrySize-1)]
        # First 64 bytes = UTF16LE name (max 32 chars)
        $nameBytes = $entry[0..63]
        $name = ([System.Text.Encoding]::Unicode.GetString($nameBytes)).Trim([char]0)

        if ($name.Length -gt 0) {
            $entries += $name
        }
    }
    $fs.close()
    $hasEncryptedPackage = $entries -contains "EncryptedPackage"
    $hasEncryptionInfo   = $entries -contains "EncryptionInfo"

    if ($hasEncryptedPackage -and $hasEncryptionInfo) {
        return [PSCustomObject]@{
            Path = $Path
            IsEncrypted = $true
            Reason = "Encrypted (contains EncryptionInfo + EncryptedPackage streams)"
        }
    }
    elseif ($hasEncryptedPackage -or $hasEncryptionInfo) {
        return [PSCustomObject]@{
            Path = $Path
            IsEncrypted = $true
            Reason = "Partially encrypted (contains EncryptionInfo + EncryptedPackage streams)"
        }
    }
    else {
        return [PSCustomObject]@{
            Path = $Path
            IsEncrypted = $false
            Reason = "OLE file but missing encryption streams"
        }
    }
}

function Get-OfficeFormat {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return [PSCustomObject]@{
            File   = $Path
            Format = "NotFound"
            Type   = "Unknown"
        }
    }

    # Read first 8 bytes
    $fs = [System.IO.File]::OpenRead($Path)
    $buffer = New-Object byte[] 8
    $fs.Read($buffer, 0, 8) | Out-Null
    $fs.Close()

    # ZIP (PK 03 04) Office Open XML (DOCX/XLSX/PPTX)
    if ($buffer[0] -eq 0x50 -and $buffer[1] -eq 0x4B -and $buffer[2] -eq 0x03 -and $buffer[3] -eq 0x04) {
        return [PSCustomObject]@{
            File   = $Path
            Format = "OpenXML"
            Type   = "2007Plus (DOCX/XLSX/PPTX)"
        }
    }

    # OLE Compound File (D0 CF 11 E0 A1 B1 1A E1) Office Binary (DOC/XLS/PPT)
    if ($buffer[0] -eq 0xD0 -and $buffer[1] -eq 0xCF -and $buffer[2] -eq 0x11 -and 
        $buffer[3] -eq 0xE0 -and $buffer[4] -eq 0xA1 -and $buffer[5] -eq 0xB1 -and
        $buffer[6] -eq 0x1A -and $buffer[7] -eq 0xE1) {

        return [PSCustomObject]@{
            File   = $Path
            Format = "BinaryOLE"
            Type   = "97-2003 (DOC/XLS/PPT)"
        }
    }

    # Unknown or corrupted format
    return [PSCustomObject]@{
        File   = $Path
        Format = "Unknown"
        Type   = "Unrecognized or Corrupt"
    }
}

function Set-OpenXmlProperties {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [hashtable]$Properties
    )

    # Helper: write an XmlDocument to a ZipArchiveEntry stream with no BOM,
    # proper XML declaration, and no extra whitespace that can confuse Office.
    function Write-XmlToZipEntry {
        param(
            [System.IO.Compression.ZipArchiveEntry]$Entry,
            [System.Xml.XmlDocument]$Doc
        )

        $entryStream = $Entry.Open()
        try {
            $settings = New-Object System.Xml.XmlWriterSettings
            $settings.Encoding           = New-Object System.Text.UTF8Encoding $false  # no BOM
            $settings.Indent             = $false
            $settings.OmitXmlDeclaration = $false
            $settings.CloseOutput        = $false

            $xw = [System.Xml.XmlWriter]::Create($entryStream, $settings)
            $Doc.Save($xw)
            $xw.Flush()
            $xw.Close()
        }
        finally {
            $entryStream.Close()
        }
    }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $tempFile = "$FilePath.tmp"

        # Remove any leftover temp file from a previous failed run
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }

        $sourceZip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        $targetZip = [System.IO.Compression.ZipFile]::Open($tempFile, 'Create')

        # ---------------------------------------------------------------
        # 1. Copy every entry we are NOT rebuilding
        # ---------------------------------------------------------------
        $rebuildParts = @("docProps/custom.xml", "[Content_Types].xml", "_rels/.rels")

        foreach ($entry in $sourceZip.Entries) {
            if ($rebuildParts -contains $entry.FullName) { continue }

            $newEntry  = $targetZip.CreateEntry($entry.FullName)
            $inStream  = $entry.Open()
            $outStream = $newEntry.Open()
            $inStream.CopyTo($outStream)
            $inStream.Close()
            $outStream.Close()
        }

        # ---------------------------------------------------------------
        # 2. Build docProps/custom.xml from scratch
        # ---------------------------------------------------------------
        $customXml = New-Object System.Xml.XmlDocument
        $customXml.AppendChild($customXml.CreateXmlDeclaration("1.0", "UTF-8", "yes")) | Out-Null

        $cpNs = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        $vtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

        $propsRoot = $customXml.CreateElement("Properties", $cpNs)
        $propsRoot.SetAttribute("xmlns:vt", $vtNs)
        $customXml.AppendChild($propsRoot) | Out-Null

        $propId = 2
        foreach ($key in $Properties.Keys) {
            $prop = $customXml.CreateElement("property", $cpNs)
            $prop.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            $prop.SetAttribute("pid",   $propId.ToString())
            $prop.SetAttribute("name",  $key)

            $vtElem           = $customXml.CreateElement("vt:lpwstr", $vtNs)
            $vtElem.InnerText = [string]$Properties[$key]

            $prop.AppendChild($vtElem) | Out-Null
            $propsRoot.AppendChild($prop) | Out-Null
            $propId++
        }

        $customEntry = $targetZip.CreateEntry("docProps/custom.xml")
        Write-XmlToZipEntry -Entry $customEntry -Doc $customXml

        # ---------------------------------------------------------------
        # 3. Patch [Content_Types].xml  — add Override for custom.xml
        # ---------------------------------------------------------------
        $ctSourceEntry = $sourceZip.GetEntry("[Content_Types].xml")
        if (-not $ctSourceEntry) {
            throw "[Content_Types].xml not found in source archive."
        }

        $ctStream = $ctSourceEntry.Open()
        $ctXml    = New-Object System.Xml.XmlDocument
        $ctXml.Load($ctStream)
        $ctStream.Close()

        $ctNsUri = "http://schemas.openxmlformats.org/package/2006/content-types"
        $ctNsMgr = New-Object System.Xml.XmlNamespaceManager($ctXml.NameTable)
        $ctNsMgr.AddNamespace("ct", $ctNsUri)

        $existingOverride = $ctXml.SelectSingleNode(
            "//ct:Override[@PartName='/docProps/custom.xml']", $ctNsMgr
        )

        if (-not $existingOverride) {
            $override = $ctXml.CreateElement("Override", $ctNsUri)
            $override.SetAttribute("PartName",    "/docProps/custom.xml")
            $override.SetAttribute("ContentType",
                "application/vnd.openxmlformats-officedocument.custom-properties+xml")
            $ctXml.DocumentElement.AppendChild($override) | Out-Null
        }

        $ctEntry = $targetZip.CreateEntry("[Content_Types].xml")
        Write-XmlToZipEntry -Entry $ctEntry -Doc $ctXml

        # ---------------------------------------------------------------
        # 4. Patch _rels/.rels — add Relationship for custom.xml
        # ---------------------------------------------------------------
        $relsSourceEntry = $sourceZip.GetEntry("_rels/.rels")
        if (-not $relsSourceEntry) {
            throw "_rels/.rels not found in source archive."
        }

        $relsStream = $relsSourceEntry.Open()
        $relsXml    = New-Object System.Xml.XmlDocument
        $relsXml.Load($relsStream)
        $relsStream.Close()

        $relNsUri = "http://schemas.openxmlformats.org/package/2006/relationships"
        $relNsMgr = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
        $relNsMgr.AddNamespace("r", $relNsUri)

        $customPropRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"

        $existingRel = $relsXml.SelectSingleNode(
            "//r:Relationship[@Type='$customPropRelType']", $relNsMgr
        )

        if (-not $existingRel) {
            # Find the highest existing rId number so we don't collide
            $maxId = ($relsXml.SelectNodes("//r:Relationship", $relNsMgr) |
                ForEach-Object {
                    if ($_.Id -match '^rId(\d+)$') { [int]$Matches[1] }
                } | Measure-Object -Maximum).Maximum

            $nextId = if ($maxId) { $maxId + 1 } else { 1 }

            $rel = $relsXml.CreateElement("Relationship", $relNsUri)
            $rel.SetAttribute("Id",     "rId$nextId")
            $rel.SetAttribute("Type",   $customPropRelType)
            $rel.SetAttribute("Target", "docProps/custom.xml")
            $relsXml.DocumentElement.AppendChild($rel) | Out-Null
        }

        $relsEntry = $targetZip.CreateEntry("_rels/.rels")
        Write-XmlToZipEntry -Entry $relsEntry -Doc $relsXml

        # ---------------------------------------------------------------
        # 5. Swap temp over original
        # ---------------------------------------------------------------
        $sourceZip.Dispose()
        $targetZip.Dispose()

        Move-Item -Force $tempFile $FilePath
        return $true
    }
    catch {
        Write-Warning "Set-OpenXmlProperties failed for '$FilePath': $($_.Exception.Message)"
        try { $sourceZip.Dispose() } catch {}
        try { $targetZip.Dispose() } catch {}
        if (Test-Path "$FilePath.tmp") {
            Remove-Item "$FilePath.tmp" -Force -ErrorAction SilentlyContinue
        }
        return $false
    }
}
function Test-FileExists {
    param([string]$fileToTest)
    if (!(Test-Path -Path $fileToTest -PathType Leaf)) {
        Write-Error "$fileToTest does not exist."
        return $false
    }
    else {
        return $true
    }
}


# Function to check if the file is being used by another process and skipping them 
function Test-FileLocked {
    param([string]$fileToTest)

    try {
        $stream = [System.IO.File]::Open(
            $fileToTest,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $stream.Close()
        return $false
    }
    catch {
        return $true
    }
}

function Add-ContentSafe {
    param(
        [string]$Path,
        [string]$Value,
        [int]$MaxRetries = 5,
        [int]$DelayMs = 50
    )

    for ($i = 0; $i -lt $MaxRetries; $i++) {
        try {
            $fs = [System.IO.File]::Open(
                $Path,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::None
            )

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

#Function to loop through a collection of files, check their age and create/update custom document properties
<#
function Update-FileAgeProperties {
    param ([System.Collections.ArrayList]$Files,
            [String] $processedFiles) #pass in an existing collection object and list of processed files

    if (!(Test-FileExists -fileToTest $processedFiles)) {
        return
    }

    $debugFile = "$($env:LOCALAPPDATA)\temp\debug.txt"

    if (!(Test-FileExists -fileToTest $debugFile)) {
        New-Item -Path $debugFile -ItemType File -Force
    }

    

    $Propertylogfolderpath = "$targetDir\PropertyUpdateLogs"
	if (! (Test-Path $Propertylogfolderpath -PathType Container)) {
		New-Item -Path $Propertylogfolderpath -ItemType Directory -Force
	}

    $Propertystatusfolderpath = "$targetDir\PropertyUpdateStatus"
	if (! (Test-Path $Propertystatusfolderpath -PathType Container)) {
		New-Item -Path $Propertystatusfolderpath -ItemType Directory -Force
	}


    $comQueue     = New-Object System.Collections.ArrayList
    $openXmlQueue = New-Object System.Collections.ArrayList
    $pdfQueue     = New-Object System.Collections.ArrayList

    $jobs = @()
    # Define output log file
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filename = "$($timestamp)_AddPropertiesLog.txt"
    $filenameProgress = "$($timestamp)_AddPropertiesStatus.txt"
    $filepath = "$Propertylogfolderpath\$filename"
    $scannedFiles = "$targetDir\FilesScanned.txt"
    $filepathProgress = "$Propertystatusfolderpath\$filenameProgress"
    $metadataDuration = 100

    $logEntryProgress = @("Filename", "StartTime", "EndTime", "Format", "Filesize", "PasswordProtected") -Join "|"
    Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress

    # Create the log file
    New-Item -Path $filepath -ItemType File -Force

    # Loop through each file in collection parameter
    foreach ($file in $Files) {
        $status = $null

        if ($file -like "*Incentives Newsletter!*.doc") {
            Write-Output "$file so skipped due to constant (Exception from HRESULT: 0x800706BE) error"
            continue
        }

        write-Output $file

        if (!(Test-FileExists -fileToTest $file)) {

            Add-ContentSafe -Path $ProcessedFiles -Value $file
            $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file file not found"
    		Add-ContentSafe -Path $filepath -Value $logEntry
    		Write-Output "$file file locked/open so skipped"

            continue
            }

        if ((Test-FileLocked -fileToTest $file)) {
    		$logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file file locked/open so skipped"
    		Add-ContentSafe -Path $filepath -Value $logEntry
    		Write-Output "$file file locked/open so skipped"

    		Add-ContentSafe -Path $skippedFiles -Value $file
    		continue
            }

        $processed = $false
        $success = $false

        $item = Get-Item -LiteralPath $file

        # Get file time metadata
        $dtLastAccessedDoc = $item.LastAccessTime
        $dtCreated = $item.CreationTime
        $dtLastModified = $item.LastWriteTime


        
        if ($item.Extension -eq ".pdf") {
            $pdfQueue.Add($file) > $null
        }
        else {
            $format = (Get-OfficeFormat $file).Format
            Write-Output "$objFile -> $format"
            if ($format -eq "OpenXML") {
                $openXmlQueue.Add($file) > $null
            }
            elseif ($format -eq "BinaryOLE") {
                $comQueue.Add($file) > $null
            }
        }

        # Write to the log
        $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") $file preparing file for property update"
        Add-ContentSafe -Path $filepath -Value $logEntry

        # --- Batch trigger
        if ($openXmlQueue.Count -ge 3) {
            $openXmlQueueCopy = @($openXmlQueue)    
            $jobs += Process-openXmlBatch `
                -batch $openXmlQueueCopy `
                -metadataDuration $metadataDuration `
                -processedFiles $processedFiles `
                -skippedFiles $skippedFiles `
                -filepathProgress $filepathProgress
                -format $format
            $openXmlQueue.Clear()
        }

        if ($pdfQueue.Count -ge 3) {
            $pdfQueueCopy = @($pdfQueue)    
            #$jobs += Process-PdfBatch `
            #    -batch $pdfQueueCopy `
            #    -metadataDuration $metadataDuration `
            #    -processedFiles $processedFiles `
            #    -skippedFiles $skippedFiles `
            #    -filepathProgress $filepathProgress
            #    -format $format
            $pdfQueue.Clear()
        }

        if ($comQueue.Count -ge 2) {   # much smaller batch for COM
            $jobs += Process-ComBatch $comQueue
            $comQueue.Clear()
        }
    }

    if ($openXmlQueue.Count -gt 0) { $jobs += Process-OpenXmlBatch $openXmlQueue }
    if ($pdfQueue.Count -gt 0) {
        $pdfQueueCopy = @($pdfQueue)    
        #$jobs += Process-PdfBatch `
        #    -batch $pdfQueueCopy `
        #    -metadataDuration $metadataDuration `
        #    -processedFiles $processedFiles `
        #    -skippedFiles $skippedFiles `
        #    -filepathProgress $filepathProgress
        #    -format $format `
        #    -filepath $filepath
        $pdfQueue.Clear()
        }
    if ($comQueue.Count -gt 0)     { $jobs += Process-ComBatch $comQueue }

    # Wait for completion
    foreach ($job in $jobs) {
        $job.AsyncWaitHandle.WaitOne()

        try {
            $output = $job.Pipe.EndInvoke($job.Handle)
            $output
        }
        catch {
            Write-Host "RUNSPACE ERROR: $($_.Exception.Message)"
        }

        $job.Pipe.Dispose()
        $ps.EndInvoke($job)
        $ps.Dispose()
        }   

}
#>
function Update-FileAgeProperties {
    param (
        [System.Collections.ArrayList]$Files,
        [string]$processedFiles,
        [string]$skippedFiles      # added — was referenced but never received
    )
 
    if (!(Test-FileExists -fileToTest $processedFiles)) {
        return
    }
 
    $debugFile = "$($env:LOCALAPPDATA)\temp\debug.txt"
    if (!(Test-Path -Path $debugFile -PathType Leaf)) {
        New-Item -Path $debugFile -ItemType File -Force | Out-Null
    }
 
    $Propertylogfolderpath = "$targetDir\PropertyUpdateLogs"
    if (!(Test-Path $Propertylogfolderpath -PathType Container)) {
        New-Item -Path $Propertylogfolderpath -ItemType Directory -Force | Out-Null
    }
 
    $Propertystatusfolderpath = "$targetDir\PropertyUpdateStatus"
    if (!(Test-Path $Propertystatusfolderpath -PathType Container)) {
        New-Item -Path $Propertystatusfolderpath -ItemType Directory -Force | Out-Null
    }
    
    $comParallelItems = 2
    $openXmlParallelItems = 8
    $pdfParallelItems = 8

    $comQueue     = New-Object System.Collections.ArrayList
    $openXmlQueue = New-Object System.Collections.ArrayList
    $pdfQueue     = New-Object System.Collections.ArrayList
 
    $jobs = @()
 
    $timestamp        = Get-Date -Format "yyyyMMdd_HHmmss"
    $filepath         = "$Propertylogfolderpath\$($timestamp)_AddPropertiesLog.txt"
    $filepathProgress = "$Propertystatusfolderpath\$($timestamp)_AddPropertiesStatus.txt"
    $metadataDuration = 100
 
    $logEntryProgress = @("Filename","StartTime","EndTime","Format","Filesize","PasswordProtected") -Join "|"
    Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress
 
    New-Item -Path $filepath -ItemType File -Force | Out-Null
 
    foreach ($file in $Files) {
 
        if ($file -like "*Incentives Newsletter!*.doc") {
            Write-Output "$file so skipped due to constant (Exception from HRESULT: 0x800706BE) error"
            continue
        }
 
        Write-Output $file
 
        if (!(Test-FileExists -fileToTest $file)) {
            Add-ContentSafe -Path $processedFiles -Value $file
            Add-ContentSafe -Path $filepath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file file not found"
            Write-Output "$file file not found so skipped"
            continue
        }
 
        if (Test-FileLocked -fileToTest $file) {
            Add-ContentSafe -Path $filepath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file file locked/open so skipped"
            Write-Output "$file file locked/open so skipped"
            Add-ContentSafe -Path $skippedFiles -Value $file
            continue
        }
 
        $item = Get-Item -LiteralPath $file
 
        if ($item.Extension -eq ".pdf") {
            $pdfQueue.Add($file) | Out-Null
        }
        else {
            $format = (Get-OfficeFormat $file).Format
            #Write-Output "$file -> $format"
            if ($format -eq "OpenXML") {
                $openXmlQueue.Add($file) | Out-Null
            }
            elseif ($format -eq "BinaryOLE") {
                $comQueue.Add($file) | Out-Null
            }
        }
 
        Add-ContentSafe -Path $filepath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file queued for property update"
 
        # --- Batch triggers ---
        if ($openXmlQueue.Count -ge $openXmlParallelItems) {
            $openXmlQueueCopy = [System.Collections.ArrayList]@($openXmlQueue)
            $jobs += Process-OpenXmlBatch `
                -batch            $openXmlQueueCopy `
                -metadataDuration $metadataDuration `
                -processedFiles   $processedFiles `
                -skippedFiles     $skippedFiles `
                -filepathProgress $filepathProgress `
                -format           $format `
                -filePathLog      $filepath
            $openXmlQueue.Clear()
        }
 
        if ($pdfQueue.Count -ge $pdfParallelItems) {
            # PDF batch processing placeholder
            $pdfQueueCopy = [System.Collections.ArrayList]@($pdfQueue)
            $jobs += Process-PdfBatch `
                -batch            $pdfQueueCopy `
                -metadataDuration $metadataDuration `
                -processedFiles   $processedFiles `
                -skippedFiles     $skippedFiles `
                -filepathProgress $filepathProgress `
                -format           $format `
                -filePathLog      $filepath
            $pdfQueue.Clear()
        }
 
        if ($comQueue.Count -ge $comParallelItems) {
            $comQueueCopy = [System.Collections.ArrayList]@($comQueue)
            $jobs += Process-ComBatch `
                -batch            $comQueueCopy `
                -metadataDuration $metadataDuration `
                -processedFiles   $processedFiles `
                -skippedFiles     $skippedFiles `
                -filepathProgress $filepathProgress `
                -filePathLog      $filepath
            $comQueue.Clear()
        }
    }
 
    # --- Drain remaining queues ---
    if ($openXmlQueue.Count -gt 0) {
        $jobs += Process-OpenXmlBatch `
            -batch            $openXmlQueue `
            -metadataDuration $metadataDuration `
            -processedFiles   $processedFiles `
            -skippedFiles     $skippedFiles `
            -filepathProgress $filepathProgress `
            -format           $format `
            -filePathLog      $filepath
    }
 
    if ($pdfQueue.Count -gt 0) {
        # PDF batch processing placeholder
            $jobs += Process-PdfBatch `
                -batch            $pdfQueueCopy `
                -metadataDuration $metadataDuration `
                -processedFiles   $processedFiles `
                -skippedFiles     $skippedFiles `
                -filepathProgress $filepathProgress `
                -format           $format `
                -filePathLog      $filepath
            $pdfQueue.Clear()
    }
 
    if ($comQueue.Count -gt 0) {
        $jobs += Process-ComBatch `
            -batch            $comQueue `
            -metadataDuration $metadataDuration `
            -processedFiles   $processedFiles `
            -skippedFiles     $skippedFiles `
            -filepathProgress $filepathProgress `
            -filePathLog      $filepath
    }
 
    # --- Wait for all jobs and collect output ---
    foreach ($job in $jobs) {
        $job.AsyncWaitHandle.WaitOne()
        try {
            $output = $job.Pipe.EndInvoke($job.Handle)
            $output
        }
        catch {
            Write-Warning "Runspace error: $($_.Exception.Message)"
        }
        finally {
            $job.Pipe.Dispose()
        }
    }
}
 

function Process-PdfBatch{
    param
        ([System.Collections.ArrayList]$batch,
            [int]$metadataDuration,
            [string]$processedFiles,
            [string]$skippedFiles,
            [string]$filepathProgress,
            [string]$format,
            [string]$filePathLog,
            [int]$parallelItems
    )

    #Python dependencies for PDF updates
    $PythonPath = "C:\Program Files\Python313\python.exe"
    $ScriptPath = "C:\Temp\update_pdf_properties.py"
    #$PythonPath = "C:\Users\UDRTagging\AppData\Local\Programs\Python\Python313\python.exe"
    #$ScriptPath = "C:\Temp\update_pdf_properties.py"


    $pool = [runspacefactory]::CreateRunspacePool(1,$parallelItems)
    $pool.Open()

    $fnAddContentSafe   = "function Add-ContentSafe { ${function:Add-ContentSafe} }"

    $jobs = @()
    
    foreach ($fileToProcess in $batch) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        $ps.AddScript($fnAddContentSafe)    | Out-Null


        $ps.AddScript({
            param($metadataDuration `
                , $processedFiles       `
                , $skippedFiles         `
                , $filepathProgress     `
                , $format               `
                , $filePathLog `
                , $pythonPath `
                , $scriptPath `
                , $fileToProcess 
            )

            $item = Get-Item -LiteralPath $fileToProcess
            $dtLastAccessedDoc = $item.LastAccessTime
            $dtCreated         = $item.CreationTime
            $dtLastModified    = $item.LastWriteTime
            $fileReadOnly      = $item.IsReadOnly
            $startTime = Get-Date
            $startTimeF = $startTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
            $filesize = $item.Length

            $isError    = $false
            $processed  = $false



            try {
                $blProperty18Months = [bool]((New-TimeSpan -Start $dtLastAccessedDoc -End (Get-Date)).TotalDays -gt 540)
                $blProperty3Years = [bool]((New-TimeSpan -Start $dtCreated -End (Get-Date)).TotalDays -gt 1095)

                # Convert boolean values to text strings for Purview to read correctly
                $strProperty18Months = if($blProperty18Months) {"True"} else {"False"}
                $strProperty3Years =if($blProperty3Years) {"True"} else {"False"}


                $result = & $pythonPath $scriptPath $fileToProcess `
                    "OriginalPath=$fileToProcess" `
                    "LastAccessed18Months=$strProperty18Months" `
                    "Created3Years=$strProperty3Years"

                switch ($result) {
                    1 {$isPasswordProtected = $true}
                    2 {$isPasswordProtected = $true}
                    -1 {$isError = $true}
                    0 {$process = $true}
                    default {
                        $isError = $true 
                        $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $fileToProcess - unexpected return code from Python: $result"
                        Add-ContentSafe -Path $filePathLog -Value $logEntry
                    }
                }
            }
            catch {
                $isError = $true
                $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $fileToProcess - exception calling Python: $($_.Exception.Message)"
                Add-ContentSafe -Path $filePathLog -Value $logEntry
            }
            finally {
                $endTime = Get-Date
                $endTimeF = $endTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")

                try {
                    $item.LastWriteTime = $dtLastModified
                    Start-Sleep -MilliSeconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
                    $item.LastAccessTime = $dtLastAccessedDoc
                    if ($fileReadOnly) { $item.IsReadOnly = $true }
                }
                catch {
                    $restoreMsg = $_.Exception.Message
                    $restoreHR = if ($_.Exception.HResult) { '{0:X8}' -f ($_.Exception.HResult) } else { $null }
                    if ($restoreMsg -match 'being used by another process' -or $hresult -eq '80070020') {
                        Write-Warning "Timestamp restore skipped; file in use: $fileToProcess ($msg)"
                        Add-ContentSafe -Path $skippedFiles -Value $fileToProcess
                    } 
                    else {
                        Write-Warning "Timestamp restore failed for $fileToProcess : $restoreMsg (HR=$restoreHR)"
                    }
                }
                #$logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $fileToProcess - file currently open or locked, properties not set"
                #Add-ContentSafe -Path $filePathLog -Value $logEntry
                #Write to log that file has been updated

                $logEntryProgress = @($fileToProcess, $startTimeF, $endTimeF, $format, $filesize, $isPasswordProtected)  -Join "|"
                Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress
                Add-ContentSafe -Path $processedFiles -Value $fileToProcess

                if($isPasswordProtected -eq $true) {
                    $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $fileToProcess - file is encrypted or digitally signed"
                    Add-ContentSafe -Path $filePathLog -Value $logEntry
                }
                elseif ($isError) {
                    $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $fileToProcess properties NOT updated - see error above"
                    Add-ContentSafe -Path $filePathLog -Value $logEntry
                    Write-Output "$fileToProcess failed - see log"

                    Write-Output "$fileToProcess properties updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                }
                else {
                    $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $fileToProcess properties NOT updated - see error above"
                    Add-ContentSafe -Path $filePathLog -Value $logEntry

                    catch {
                        $msg     = $_.Exception.Message
                        $hresult = if ($_.Exception.HResult) { '{0:X8}' -f ($_.Exception.HResult) } else { $null }

                        if ($msg -match 'being used by another process' -or $hresult -eq '80070020') {
                            $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") $fileToProcess file currently open or locked, properties not set"
                            Add-ContentSafe -Path $filePathLog -Value $logEntry
                            Write-Warning "Timestamp restore skipped; file in use: $fileToProcess ($msg)"
                            Add-ContentSafe -Path $skippedFiles -Value $fileToProcess
                        } 
                        else {
                            $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ; $fileToProcess ; properties updated"
                            Add-ContentSafe -Path $filePathLog -Value $logEntry                        
                            Write-Output "$filetoProcess properties updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                        }
                    }                    
                }
            }
        })  | Out-Null
        $ps.AddArgument($metadataDuration)| Out-Null
        $ps.AddArgument($processedFiles)| Out-Null
        $ps.AddArgument($skippedFiles)| Out-Null
        $ps.AddArgument($filepathProgress)| Out-Null
        $ps.AddArgument($format)| Out-Null
        $ps.AddArgument($filePathLog)| Out-Null
        $ps.AddArgument($PythonPath)| Out-Null
        $ps.AddArgument($scriptPath)| Out-Null
        $ps.AddArgument($fileToProcess)| Out-Null

        
        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }
 
    return $jobs
}

<#
function Process-OpenXMLBatch{
    param
        ([System.Collections.ArrayList]$batch `
        , [int]$metadataDuration              `
        , [string]$processedFiles             `
        , [string]$skippedFiles               `
        , [string]$filepathProgress           `
        , [string]$format                     `
        , [string]$filePathLog
    )

    $pool = [runspacefactory]::CreateRunspacePool(1,3)
    $pool.Open()

    $jobs = @()
    
    foreach ($fileToProcess in $batch) {
        try {
            $stream = [System.IO.File]::Open($fileToProcess, 'Open', 'ReadWrite', 'None')
            $stream.Close()
        }
        catch {
            return
        }        
        
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        
        
        $ps.AddScript("function Set-OpenXmlProperties { ${function:Set-OpenXmlProperties} }")
        $ps.AddScript("function Test-OfficeEncrypted { ${function:Test-OfficeEncrypted} }")
        $ps.AddScript("function Add-ContentSafe { ${function:Add-ContentSafe} }")

        $ps.AddScript({
            param($metadataDuration `
                , $processedFiles       `
                , $skippedFiles         `
                , $filepathProgress     `
                , $format               `
                , $filePathLog `
                , $fileToProcess 
            )



            $isError = $false
            $processed = $false

            Add-ContentSafe -Path "$($env:LOCALAPPDATA)\temp\debug.txt" -Value "Processing $fileToProcess"

            $isPasswordProtected = (Test-OfficeEncrypted -Path $fileToProcess).IsEncrypted

            $item = Get-Item -LiteralPath $fileToProcess
            $dtLastAccessedDoc = $item.LastAccessTime
            $dtCreated         = $item.CreationTime
            $dtLastModified    = $item.LastWriteTime
            $fileReadOnly      = $item.IsReadOnly
            $startTime = Get-Date
            $startTimeF = $startTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
            $filesize = $item.Length



            try {
                $blProperty18Months = [bool]((New-TimeSpan -Start $dtLastAccessedDoc -End (Get-Date)).TotalDays -gt 540)
                $blProperty3Years = [bool]((New-TimeSpan -Start $dtCreated -End (Get-Date)).TotalDays -gt 1095)

                # Convert boolean values to text strings for Purview to read correctly
                $strProperty18Months = if($blProperty18Months) {"True"} else {"False"}
                $strProperty3Years =if($blProperty3Years) {"True"} else {"False"}
                
                $props = @{
                    OriginalPath = $fileToProcess
                    LastAccessed18Months = $strProperty18Months
                    Created3Years = $strProperty3Years
                }
                if (-not $isPasswordProtected) {
                    Set-OpenXmlProperties -FilePath $fileToProcess -Properties $props
                }
            }
            catch {
                Write-Output "Failed $fileToProcess"
            }
            finally {
                    $endTime = Get-Date
                    $endTimeF = $endTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    if($isPasswordProtected -eq $true) {
                        #Write to log that file has been updated
                        try {
                            $item.LastWriteTime = $dtLastModified
                            Start-Sleep -MilliSeconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
                            $item.LastAccessTime = $dtLastAccessedDoc
                            if ($fileReadOnly) { $item.IsReadOnly = $true }

                            $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") $fileToProcess file is password-protected"
                            Add-ContentSafe -Path $filePathLog -Value $logEntry
                            
                            $logEntryProgress = @($fileToProcess, $startTimeF, $endTimeF, $format, $filesize, $isPasswordProtected)  -Join "|"
                            Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress

                            #Write to processed list that file has been updated
                            Add-ContentSafe -Path $processedFiles -Value $fileToProcess
                            Write-Output "$fileToProcess file is password-protected"
                        }
                        catch {
                            $msg     = $_.Exception.Message
                            $hresult = if ($_.Exception.HResult) { '{0:X8}' -f ($_.Exception.HResult) } else { $null }
                            if ($fileReadOnly) { $item.IsReadOnly = $true }
                            if ($msg -match 'being used by another process' -or $hresult -eq '80070020') {
                                $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ; $fileToProcess ; file currently open or locked, properties not set"
                                Add-ContentSafe -Path $filePathLog -Value $logEntry
                                Write-Warning "Timestamp restore skipped; file in use: $fileToProcess ($msg)"
                                Add-ContentSafe -Path $skippedFiles -Value $fileToProcess
                            } else {
                                Write-Warning "Timestamp restore failed (unexpected) for $fileToProcess : $msg (HR=$hresult)"
                            }
                        }
                        
                    } 
                    elseif ($isPasswordProtected -eq $false -and $isError -eq $false) {
                        $processed = $true
                    }
                }

                if($processed -eq $true) {
                    # Update file timestamps
                    try {
                        $item.LastWriteTime = $dtLastModified
                        Start-Sleep -MilliSeconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
                        $item.LastAccessTime = $dtLastAccessedDoc
                        if ($fileReadOnly) { $item.IsReadOnly = $true }

                        $success = $true

                        #Write to log that file has been updated
                        $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ; $fileToProcess ; properties updated"
                        Add-ContentSafe -Path $filePathLog -Value $logEntry

                        #Write to data that file has been updated
                        $logEntryProgress = @($fileToProcess, $startTimeF, $endTimeF, $format, $filesize, $isPasswordProtected)  -Join "|"
                        Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress

                        #Write to processed list that file has been updated
                        Add-ContentSafe -Path $processedFiles -Value $fileToProcess

                        Write-Output "$filetoProcess properties updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                    }
                    catch {
                        $msg     = $_.Exception.Message
                        $hresult = if ($_.Exception.HResult) { '{0:X8}' -f ($_.Exception.HResult) } else { $null }
                        if ($fileReadOnly) { $item.IsReadOnly = $true }
                        if ($msg -match 'being used by another process' -or $hresult -eq '80070020') {
                            $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") so $fileToProcess so file currently open or locked, properties not set"
                            Add-ContentSafe -Path $filePathLog -Value $logEntry
                            Write-Warning "Timestamp restore skipped; file in use: $fileToProcess ($msg)"
                            Add-ContentSafe -Path $skippedFiles -Value $fileToProcess
                        } else {
                            Write-Warning "Timestamp restore failed (unexpected) for $fileToProcess : $msg (HR=$hresult)"
                        }
                    }
                }
        }).
        AddArgument($metadataDuration).
        AddArgument($processedFiles).
        AddArgument($skippedFiles).
        AddArgument($filepathProgress).
        AddArgument($format).
        AddArgument($filePathLog).
        AddArgument($fileToProcess)



        $jobs += [pscustomobject]@{
            Pipe = $ps
            Handle = $ps.BeginInvoke()

        }

    }

    return $jobs

}
#>

function Process-OpenXmlBatch {
    param(
        [System.Collections.ArrayList]$batch,
        [int]$metadataDuration,
        [string]$processedFiles,
        [string]$skippedFiles,
        [string]$filepathProgress,
        [string]$format,
        [string]$filePathLog,
        [int]$parallelItems
    )
 
    $pool = [runspacefactory]::CreateRunspacePool(1, $parallelItems)
    $pool.Open()
 
    # Capture function definitions once, outside the loop
    $fnSetOpenXml       = "function Set-OpenXmlProperties { ${function:Set-OpenXmlProperties} }"
    $fnTestEncrypted    = "function Test-OfficeEncrypted { ${function:Test-OfficeEncrypted} }"
    $fnAddContentSafe   = "function Add-ContentSafe { ${function:Add-ContentSafe} }"
 
    $jobs = @()
 
    foreach ($fileToProcess in $batch) {
 
        # Skip files that are locked before even spinning up a runspace
        try {
            $stream = [System.IO.File]::Open($fileToProcess, 'Open', 'ReadWrite', 'None')
            $stream.Close()
        }
        catch {
            Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess locked before runspace — skipped"
            Add-ContentSafe -Path $skippedFiles -Value $fileToProcess
            continue
        }
 
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
 
        # Inject dependencies into the runspace as named functions
        $ps.AddScript('Add-Type -AssemblyName System.IO.Compression.FileSystem') | Out-Null
        $ps.AddScript($fnAddContentSafe)    | Out-Null
        $ps.AddScript($fnTestEncrypted)     | Out-Null
        $ps.AddScript($fnSetOpenXml)        | Out-Null
 
        $ps.AddScript({
            param(
                $metadataDuration,
                $processedFiles,
                $skippedFiles,
                $filepathProgress,
                $format,
                $filePathLog,
                $fileToProcess
            )
 
            $isError    = $false
            $processed  = $false
 
            $isPasswordProtected = (Test-OfficeEncrypted -Path $fileToProcess).IsEncrypted
 
            $item              = Get-Item -LiteralPath $fileToProcess
            $dtLastAccessedDoc = $item.LastAccessTime
            $dtCreated         = $item.CreationTime
            $dtLastModified    = $item.LastWriteTime
            $fileReadOnly      = $item.IsReadOnly
            $startTime         = Get-Date
            $startTimeF        = $startTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
            $filesize          = $item.Length
 
            try {
                if (-not $isPasswordProtected) {
                    $blProperty18Months  = [bool]((New-TimeSpan -Start $dtLastAccessedDoc -End (Get-Date)).TotalDays -gt 540)
                    $blProperty3Years    = [bool]((New-TimeSpan -Start $dtCreated         -End (Get-Date)).TotalDays -gt 1095)
                    $strProperty18Months = if ($blProperty18Months) { "True" } else { "False" }
                    $strProperty3Years   = if ($blProperty3Years)   { "True" } else { "False" }
 
                    $props = @{
                        OriginalPath         = $fileToProcess
                        LastAccessed18Months = $strProperty18Months
                        Created3Years        = $strProperty3Years
                    }
 
                    $setResult = Set-OpenXmlProperties -FilePath $fileToProcess -Properties $props
                    if (-not $setResult) {
                        $isError = $true
                        Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess Set-OpenXmlProperties returned false"
                    }
                }
            }
            catch {
                $isError = $true
                Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess exception during property update: $($_.Exception.Message)"
                Write-Output "Failed $fileToProcess : $($_.Exception.Message)"
            }
            finally {
                $endTime  = Get-Date
                $endTimeF = $endTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
 
                # Restore timestamps and log result regardless of outcome
                try {
                    $item.LastWriteTime  = $dtLastModified
                    Start-Sleep -Milliseconds $metadataDuration
                    $item.LastAccessTime = $dtLastAccessedDoc
                    if ($fileReadOnly) { $item.IsReadOnly = $true }
                }
                catch {
                    $msg     = $_.Exception.Message
                    $hresult = if ($_.Exception.HResult) { '{0:X8}' -f $_.Exception.HResult } else { $null }
                    if ($msg -match 'being used by another process' -or $hresult -eq '80070020') {
                        Write-Warning "Timestamp restore skipped; file in use: $fileToProcess ($msg)"
                        Add-ContentSafe -Path $skippedFiles -Value $fileToProcess
                    }
                    else {
                        Write-Warning "Timestamp restore failed for $fileToProcess : $msg (HR=$hresult)"
                    }
                }
 
                $logEntryProgress = @($fileToProcess, $startTimeF, $endTimeF, $format, $filesize, $isPasswordProtected) -Join "|"
                Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress
                Add-ContentSafe -Path $processedFiles   -Value $fileToProcess
 
                if ($isPasswordProtected) {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess file is password-protected"
                    Write-Output "$fileToProcess skipped — password-protected"
                }
                elseif ($isError) {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess properties NOT updated — see error above"
                    Write-Output "$fileToProcess failed — see log"
                }
                else {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess properties updated"
                    Write-Output "$fileToProcess properties updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                }
            }
 
        }) | Out-Null
 
        $ps.AddArgument($metadataDuration) | Out-Null
        $ps.AddArgument($processedFiles)   | Out-Null
        $ps.AddArgument($skippedFiles)     | Out-Null
        $ps.AddArgument($filepathProgress) | Out-Null
        $ps.AddArgument($format)           | Out-Null
        $ps.AddArgument($filePathLog)      | Out-Null
        $ps.AddArgument($fileToProcess)    | Out-Null
 
        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }
 
    return $jobs
}

<#
function Process-COMBatch{
    param
        ([System.Collections.ArrayList]$batch `
        , [int]$metadataDuration              `
        , [string]$processedFiles             `
        , [string]$skippedFiles               `
        , [string]$filepathProgress           `
        , [string]$format                     `
        , [string]$filePathLog
    )


    $pool = [runspacefactory]::CreateRunspacePool(1,2)
    $pool.Open()

    # Capture function definitions once, outside the loop
    $fnSetOfficeDoc       = "function Set-OfficeDocCustomProperty { ${function:Set-OfficeDocCustomProperty} }"
    $fnTestEncrypted    = "function Test-OfficeEncrypted { ${function:Test-OfficeEncrypted} }"
    $fnTestEncryptedPpt2003    = "function Test-Ppt2003HasOpenPassword { ${function:Test-Ppt2003HasOpenPassword} }"
    $fnAddContentSafe   = "function Add-ContentSafe { ${function:Add-ContentSafe} }"
    $fnHandleError   = "function Handle-FileProcessingError { ${function:Handle-FileProcessingError} }"


    $jobs = @()
    
    foreach ($filePath in $batch) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        $ps.AddScript($fnHandleError)    | Out-Null
        $ps.AddScript($fnAddContentSafe)    | Out-Null
        $ps.AddScript($fnTestEncrypted)     | Out-Null
        $ps.AddScript($fnTestEncryptedPpt2003)     | Out-Null
        $ps.AddScript($fnSetOfficeDoc)        | Out-Null


        $ps.AddScript({
            param($metadataDuration `
                , $processedFiles       `
                , $skippedFiles         `
                , $filepathProgress     `
                , $format               `
                , $filePathLog `
                , $file 
            )

            $item = Get-Item -LiteralPath $file

            if ($item.Extension = ".ppt") {
                $isPasswordProtected = Test-Ppt2003HasOpenPassword -Path $filePath
            }
            else {
                $isPasswordProtected = (Test-OfficeEncrypted -Path $filePath).IsEncrypted
            }

            $item = Get-Item -LiteralPath $file
            $dtLastAccessedDoc = $item.LastAccessTime
            $dtCreated         = $item.CreationTime
            $dtLastModified    = $item.LastWriteTime
            $fileReadOnly      = $item.IsReadOnly
            $startTime = Get-Date
            $startTimeF = $startTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
            $filesize = $item.Length


            try{
                switch -regex ($item.Extension) {
				".docx|.docm|.doc" {
					if($isPasswordProtected -eq $false) {
						try {
                            # Write-Host "Trying to open Word document"
							$app = New-Object -ComObject Word.Application
                            if($app -eq $null) {
                                Write-Error "Application COM object failed to initialise"
                                $action = Handle-FileProcessingError `
                                    -ErrorRecord $_ `
                                    -File $file `
                                    -Status ([ref]$status) `
                                    -App $app `
                                    -Doc $doc `
                                    -Item $item `
                                    -ObjFile $file `
                                    -ProcessedFiles $processedFiles `
                                    -LastWriteTime $dtLastModified `
                                    -LastAccessTime $dtLastAccessedDoc `
                                    -MetadataDuration $metadataDuration `
                                    -FileReadOnly $fileReadOnly `
                                    -FilePath $filePath `
                                    -FilePathProgress $filePathProgress `
                                    -StartTime $startTime `
                                    -Format $format `
                                    -FileSize $filesize

                                # if ($action -eq [FileErrorAction]::ContinueFile) {
                                if ($action -eq "ContinueFile") {
                                    Start-Sleep -MilliSeconds $metadataDuration
                                    continue
                                    }
                                }
                            }                            

                        catch {
                            $action = Handle-FileProcessingError `
                                -ErrorRecord $_ `
                                -File $file `
                                -Status ([ref]$status) `
                                -App $app `
                                -Doc $doc `
                                -Item $item `
                                -ObjFile $file `
                                -ProcessedFiles $processedFiles `
                                -LastWriteTime $dtLastModified `
                                -LastAccessTime $dtLastAccessedDoc `
                                -MetadataDuration $metadataDuration `
                                -FileReadOnly $fileReadOnly `
                                -FilePath $filePath `
                                -FilePathProgress $filePathProgress `
                                -StartTime $startTime `
                                -Format $format `
                                -FileSize $filesize

                            # if ($action -eq [FileErrorAction]::ContinueFile) {
                            if ($action -eq "ContinueFile") {
                                Start-Sleep -MilliSeconds $metadataDuration
                                continue
                            }
                        }


					}
				}

			    ".xlsx|.xlsm|.xls|.xlsb" {
				   	
					if (-not (Get-Variable -Name status -Scope Local -ErrorAction SilentlyContinue)) {
    					$status = $null
					}

				    if ($isPasswordProtected -eq $false) {
					    try {
                            # Write-Host "Trying to open Excel document"
						    $app = New-Object -ComObject Excel.Application
                            if($app -eq $null) {
                                Write-Error "Application COM object failed to initialise"
                                $action = Handle-FileProcessingError `
                                    -ErrorRecord $_ `
                                    -File $file `
                                    -Status ([ref]$status) `
                                    -App $app `
                                    -Doc $doc `
                                    -Item $item `
                                    -ObjFile $file `
                                    -ProcessedFiles $processedFiles `
                                    -LastWriteTime $dtLastModified `
                                    -LastAccessTime $dtLastAccessedDoc `
                                    -MetadataDuration $metadataDuration `
                                    -FileReadOnly $fileReadOnly `
                                    -FilePath $filePath `
                                    -FilePathProgress $filePathProgress `
                                    -StartTime $startTime `
                                    -Format $format `
                                    -FileSize $filesize

                                # if ($action -eq [FileErrorAction]::ContinueFile) {
                                if ($action -eq "ContinueFile") {
                                    Start-Sleep -MilliSeconds $metadataDuration
                                    continue
                                }
                            }

                        }

                        catch {
                            $action = Handle-FileProcessingError `
                                -ErrorRecord $_ `
                                -File $file `
                                -Status ([ref]$status) `
                                -App $app `
                                -Doc $doc `
                                -Item $item `
                                -ObjFile $file `
                                -ProcessedFiles $processedFiles `
                                -LastWriteTime $dtLastModified `
                                -LastAccessTime $dtLastAccessedDoc `
                                -MetadataDuration $metadataDuration `
                                -FileReadOnly $fileReadOnly `
                                -FilePath $filePath `
                                -FilePathProgress $filePathProgress `
                                -StartTime $startTime `
                                -Format $format `
                                -FileSize $filesize

                            # if ($action -eq [FileErrorAction]::ContinueFile) {
                            if ($action -eq "ContinueFile") {
                                Start-Sleep -MilliSeconds $metadataDuration
                                continue
                            }
                        }
                    }
                }

			    ".pptx|.pptm|.ppt" {
				    if($isPasswordProtected -eq $false) {
					    $app = New-Object -ComObject PowerPoint.Application
					    try {
                            # Write-Host "Trying to open PowerPoint document"
						    $doc = $app.Presentations.Open($file, $false, $false, $false)
                            if($app -eq $null) {
                                Write-Error "Application COM object failed to initialise"
                                $action = Handle-FileProcessingError `
                                    -ErrorRecord $_ `
                                    -File $file `
                                    -Status ([ref]$status) `
                                    -App $app `
                                    -Doc $doc `
                                    -Item $item `
                                    -ObjFile $file `
                                    -ProcessedFiles $processedFiles `
                                    -LastWriteTime $dtLastModified `
                                    -LastAccessTime $dtLastAccessedDoc `
                                    -MetadataDuration $metadataDuration `
                                    -FileReadOnly $fileReadOnly `
                                    -FilePath $filePath `
                                    -FilePathProgress $filePathProgress `
                                    -StartTime $startTime `
                                    -Format $format `
                                    -FileSize $filesize

                                # if ($action -eq [FileErrorAction]::ContinueFile) {
                                if ($action -eq "ContinueFile") {
                                    Start-Sleep -MilliSeconds $metadataDuration
                                    continue
                                }
                            }
                            else {
                                # Write-Host "1"
						        $doc.Saved = $false
                                # Write-Host "2"
                                $format = ".ppt"
						        $officeApp = $true
                                # Write-Host "Successfully opened PowerPoint document"
                            }
					    }

                        catch {
                            $action = Handle-FileProcessingError `
                                -ErrorRecord $_ `
                                -File $file `
                                -Status ([ref]$status) `
                                -App $app `
                                -Doc $doc `
                                -Item $item `
                                -ObjFile $file `
                                -ProcessedFiles $processedFiles `
                                -LastWriteTime $dtLastModified `
                                -LastAccessTime $dtLastAccessedDoc `
                                -MetadataDuration $metadataDuration `
                                -FileReadOnly $fileReadOnly `
                                -FilePath $filePath `
                                -FilePathProgress $filePathProgress `
                                -StartTime $startTime `
                                -Format $format `
                                -FileSize $filesize

                            # if ($action -eq [FileErrorAction]::ContinueFile) {
                            if ($action -eq "ContinueFile") {
                                Start-Sleep -MilliSeconds $metadataDuration
                                continue
                            }
                        }

				    }
			    }

                default { Write-Host "No match found for extension: $($item.Extension)" }
                }
            }
            catch {
                    $action = Handle-FileProcessingError `
                                -ErrorRecord $_ `
                                -File $objFile `
                                -Status ([ref]$status) `
                                -App $app `
                                -Doc $doc `
                                -Item $item `
                                -ObjFile $objFile `
                                -ProcessedFiles $processedFiles `
                                -LastWriteTime $dtLastModified `
                                -LastAccessTime $dtLastAccessedDoc `
                                -MetadataDuration $metadataDuration `
                                -FileReadOnly $fileReadOnly `
                                -FilePath $filePath `
                                -FilePathProgress $filePathProgress `
                                -StartTime $startTime `
                                -Format $format `
                                -FileSize $filesize

                    $errortext = $($_.Exception.Message)
                    $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") so $objFile so file not updatable, properties not set.Error: $errortext"
                    if($format -eq ".xls" -and $app -ne $null) {
                        $app.EnableEvents = $false
                    }
                    Add-ContentSafe -Path $filepath -Value $logEntry
                    #Write-Output "$objFile so file not updatable, properties not set"
                    #Write-Output "Detailed Error: $($_.Exception)"
                    if($doc -ne $null) {
                        $doc.Close()
                    }
                    if($app -ne $null) {
                        $app.Quit()
                    }
                    if($app -ne $null) {
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
                    }
                    $doc = $null
                    $app = $null
                    # Update file timestamps
                    if($item -eq $null) {
                        $item = Get-Item -LiteralPath $objFile
                        }
                    $item.LastWriteTime = $dtLastModified
                    Start-Sleep -MilliSeconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
                    $item.LastAccessTime = $dtLastAccessedDoc
                    if($fileReadOnly -eq $true) {
                        $item.IsReadOnly = $true
                    }
                    # Start-Sleep -MilliSeconds $metadataDuration

                    # if ($action -eq [FileErrorAction]::ContinueFile) {
                    if ($action -eq "ContinueFile") {
                        Start-Sleep -MilliSeconds $metadataDuration
                        continue
                    }
                }              
                    
            try {
                $blProperty18Months = [bool]((New-TimeSpan -Start $dtLastAccessedDoc -End (Get-Date)).TotalDays -gt 540)
                $blProperty3Years = [bool]((New-TimeSpan -Start $dtCreated -End (Get-Date)).TotalDays -gt 1095)

                # Convert boolean values to text strings for Purview to read correctly
                $strProperty18Months = if($blProperty18Months) {"True"} else {"False"}
                $strProperty3Years =if($blProperty3Years) {"True"} else {"False"}
                
                $propertyExistsOriginalPath = Set-OfficeDocCustomProperty "OriginalPath" $file $doc
                $propertyExistsLastAccessed18Months = Set-OfficeDocCustomProperty "LastAccessed18Months" $strProperty18Months $doc
                $propertyExistsCreated3Years = Set-OfficeDocCustomProperty "Created3Years" $strProperty3Years $doc

                #$doc.Save()
                    
                #while ($doc.Saved -eq $false) {
                    #   start-sleep -milliseconds 100  
                #}  

                try {
                    $doc.Save()
                    $doc.Close()
                } 
                catch {
                    $status = "Document cannot be saved"
                    $action = Handle-FileProcessingError `
                            -ErrorRecord $_ `
                            -File $file `
                            -Status ([ref]$status) `
                            -App $app `
                            -Doc $doc `
                            -Item $item `
                            -ObjFile $file `
                            -ProcessedFiles $processedFiles `
                            -LastWriteTime $dtLastModified `
                            -LastAccessTime $dtLastAccessedDoc `
                            -MetadataDuration $metadataDuration `
                            -FileReadOnly $fileReadOnly `
                            -FilePath $filePath `
                            -FilePathProgress $filePathProgress `
                            -StartTime $startTime `
                            -Format $format `
                            -FileSize $filesize

                    if ($action -eq "ContinueFile") {
                            Start-Sleep -MilliSeconds $metadataDuration
                            continue
                    }

                    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') so $file so file cannot be saved so skipped"
    	            Add-ContentSafe -Path $filepath -Value $logEntry

                    Add-ContentSafe -Path $processedFiles -Value $file

   		            Add-ContentSafe -Path $skippedFiles -Value $file
   		            continue
                }
                       
                if($format -eq ".xls" -and $app -ne $null) {
                    $app.EnableEvents = $false
                }
                if ($app -ne $Null) {
                    $app.Quit()
                }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
                $app = $null
                $processed = $true
                
            }
            catch {
                Write-Output "Failed $filePath"
            }
            finally {
                    if($isPasswordProtected -eq $true) {
                        #Write to log that file has been updated
                        try {
                            $item.LastWriteTime = $dtLastModified
                            Start-Sleep -MilliSeconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
                            $item.LastAccessTime = $dtLastAccessedDoc
                            if ($fileReadOnly) { $item.IsReadOnly = $true }

                            $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") $file file is password-protected"
                            Add-ContentSafe -Path $filepath -Value $logEntry
                            
                            $logEntryProgress = @($file, $startTimeF, $endTimeF, $format, $filesize, $isPasswordProtected)  -Join "|"
                            Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress

                            #Write to processed list that file has been updated
                            Add-ContentSafe -Path $processedFiles -Value $file
                            Write-Output "$file file is password-protected"
                        }
                        catch {
                            $msg     = $_.Exception.Message
                            $hresult = if ($_.Exception.HResult) { '{0:X8}' -f ($_.Exception.HResult) } else { $null }

                            if ($msg -match 'being used by another process' -or $hresult -eq '80070020') {
                                $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") so $file so file currently open or locked, properties not set"
                                Add-ContentSafe -Path $filepath -Value $logEntry
                                Write-Warning "Timestamp restore skipped; file in use: $file ($msg)"
                                Add-ContentSafe -Path $skippedFiles -Value $file
                            } else {
                                Write-Warning "Timestamp restore failed (unexpected) for $file : $msg (HR=$hresult)"
                            }
                        }
                        
                    } elseif ($isPasswordProtected -eq $false -and $isError -eq $false) {
                        $processed = $true
                    }

                if($processed -eq $true) {
                    # Update file timestamps
                    try {
                        $item.LastWriteTime = $dtLastModified
                        Start-Sleep -MilliSeconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
                        $item.LastAccessTime = $dtLastAccessedDoc
                        if ($fileReadOnly) { $item.IsReadOnly = $true }

                        $success = $true
                        $endTime = Get-Date
                        $endTimeF = $endTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")

                        #Write to log that file has been updated
                        $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") so $file so properties updated"
                        Add-ContentSafe -Path $filepath -Value $logEntry

                        #Write to data that file has been updated
                        $logEntryProgress = @($file, $startTimeF, $endTimeF, $format, $filesize, $isPasswordProtected)  -Join "|"
                        Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress

                        #Write to processed list that file has been updated
                        Add-ContentSafe -Path $processedFiles -Value $file

                        Write-Output "$file so properties updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                    }
                    catch {
                        $msg     = $_.Exception.Message
                        $hresult = if ($_.Exception.HResult) { '{0:X8}' -f ($_.Exception.HResult) } else { $null }

                        if ($msg -match 'being used by another process' -or $hresult -eq '80070020') {
                            $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") so $file so file currently open or locked, properties not set"
                            Add-ContentSafe -Path $filepath -Value $logEntry
                            Write-Warning "Timestamp restore skipped; file in use: $file ($msg)"
                            Add-ContentSafe -Path $skippedFiles -Value $file
                        } else {
                            Write-Warning "Timestamp restore failed (unexpected) for $file : $msg (HR=$hresult)"
                        }
                    }
                }
            }
       })
        $ps.AddArgument($metadataDuration) | Out-Null
        $ps.AddArgument($processedFiles)   | Out-Null
        $ps.AddArgument($skippedFiles)     | Out-Null
        $ps.AddArgument($filepathProgress) | Out-Null
        $ps.AddArgument($format)           | Out-Null
        $ps.AddArgument($filePathLog)      | Out-Null
        $ps.AddArgument($fileToProcess)    | Out-Null
 
        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }
 
    return $jobs
}
#>

function Process-COMBatch {
    param(
        [System.Collections.ArrayList]$batch,
        [int]$metadataDuration,
        [string]$processedFiles,
        [string]$skippedFiles,
        [string]$filepathProgress,
        [string]$format,
        [string]$filePathLog,
        [int]$parallelItems
    )
 
    $pool = [runspacefactory]::CreateRunspacePool(1, $parallelItems)
    $pool.Open()
 
    # Capture all required function definitions once, outside the loop
    $fnSetOfficeDoc         = "function Set-OfficeDocCustomProperty { ${function:Set-OfficeDocCustomProperty} }"
    $fnTestEncrypted        = "function Test-OfficeEncrypted { ${function:Test-OfficeEncrypted} }"
    $fnTestEncryptedPpt2003 = "function Test-Ppt2003HasOpenPassword { ${function:Test-Ppt2003HasOpenPassword} }"
    $fnAddContentSafe       = "function Add-ContentSafe { ${function:Add-ContentSafe} }"
    $fnHandleError          = "function Handle-FileProcessingError { ${function:Handle-FileProcessingError} }"
    # Write-Log and Write-LogProcess are called by Handle-FileProcessingError — must also be injected
    $fnWriteLog             = "function Write-Log { ${function:Write-Log} }"
    $fnWriteLogProcess      = "function Write-LogProcess { ${function:Write-LogProcess} }"
 
    $jobs = @()
 
    foreach ($filePath in $batch) {
 
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
 
        $ps.AddScript($fnAddContentSafe)       | Out-Null
        $ps.AddScript($fnWriteLog)             | Out-Null
        $ps.AddScript($fnWriteLogProcess)      | Out-Null
        $ps.AddScript($fnHandleError)          | Out-Null
        $ps.AddScript($fnTestEncrypted)        | Out-Null
        $ps.AddScript($fnTestEncryptedPpt2003) | Out-Null
        $ps.AddScript($fnSetOfficeDoc)         | Out-Null
 
        $ps.AddScript({
            param(
                $metadataDuration,
                $processedFiles,
                $skippedFiles,
                $filepathProgress,
                $format,
                $filePathLog,
                $file
            )
 
            $isError           = $false   # initialise before any branch can reference it
            $processed         = $false
            $isPasswordProtected = $false
            $status            = $null
            $app               = $null
            $doc               = $null
 
            # --- Encryption check ---
            $item = Get-Item -LiteralPath $file   # get item once before the check
 
            if ($item.Extension -eq ".ppt") {     # -eq not = 
                $isPasswordProtected = Test-Ppt2003HasOpenPassword -Path $file
            }
            else {
                $isPasswordProtected = (Test-OfficeEncrypted -Path $file).IsEncrypted
            }
 
            $dtLastAccessedDoc = $item.LastAccessTime
            $dtCreated         = $item.CreationTime
            $dtLastModified    = $item.LastWriteTime
            $fileReadOnly      = $item.IsReadOnly
            $startTime         = Get-Date
            $startTimeF        = $startTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
            $filesize          = $item.Length
 
            try {
                switch -regex ($item.Extension) {
 
                    # ----------------------------------------------------------
                    # Word
                    # ----------------------------------------------------------
                    "\.docx|\.docm|\.doc" {
                        if ($isPasswordProtected -eq $false) {
                            try {
                                $app = New-Object -ComObject Word.Application
                                $app.Visible = $false
                                $app.DisplayAlerts = 0   # wdAlertsNone
 
                                if ($app -eq $null) {
                                    Write-Error "Word COM object failed to initialise"
                                    $action = Handle-FileProcessingError `
                                        -ErrorRecord      $_ `
                                        -File             $file `
                                        -Status           ([ref]$status) `
                                        -App              $app `
                                        -Doc              $doc `
                                        -Item             $item `
                                        -ObjFile          $file `
                                        -ProcessedFiles   $processedFiles `
                                        -SkippedFiles     $skippedFiles `
                                        -LastWriteTime    $dtLastModified `
                                        -LastAccessTime   $dtLastAccessedDoc `
                                        -MetadataDuration $metadataDuration `
                                        -FileReadOnly     $fileReadOnly `
                                        -FilePath         $filePathLog `
                                        -FilePathProgress $filePathProgress `
                                        -StartTime        $startTime `
                                        -Format           $format `
                                        -FileSize         $filesize
                                    if ($action -eq "ContinueFile") { continue }
                                }
                                else {
                                    $doc = $app.Documents.Open($file, $false, $false)
                                }
                            }
                            catch {
                                $action = Handle-FileProcessingError `
                                    -ErrorRecord      $_ `
                                    -File             $file `
                                    -Status           ([ref]$status) `
                                    -App              $app `
                                    -Doc              $doc `
                                    -Item             $item `
                                    -ObjFile          $file `
                                    -ProcessedFiles   $processedFiles `
                                    -SkippedFiles     $skippedFiles `
                                    -LastWriteTime    $dtLastModified `
                                    -LastAccessTime   $dtLastAccessedDoc `
                                    -MetadataDuration $metadataDuration `
                                    -FileReadOnly     $fileReadOnly `
                                    -FilePath         $filePathLog `
                                    -FilePathProgress $filePathProgress `
                                    -StartTime        $startTime `
                                    -Format           $format `
                                    -FileSize         $filesize
                                if ($action -eq "ContinueFile") { continue }
                            }
                        }
                    }
 
                    # ----------------------------------------------------------
                    # Excel
                    # ----------------------------------------------------------
                    "\.xlsx|\.xlsm|\.xls|\.xlsb" {
                        if ($isPasswordProtected -eq $false) {
                            try {
                                $app = New-Object -ComObject Excel.Application
                                $app.Visible        = $false
                                $app.DisplayAlerts  = $false
                                $app.EnableEvents   = $false
                                $app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
 
                                if ($app -eq $null) {
                                    Write-Error "Excel COM object failed to initialise"
                                    $action = Handle-FileProcessingError `
                                        -ErrorRecord      $_ `
                                        -File             $file `
                                        -Status           ([ref]$status) `
                                        -App              $app `
                                        -Doc              $doc `
                                        -Item             $item `
                                        -ObjFile          $file `
                                        -ProcessedFiles   $processedFiles `
                                        -SkippedFiles     $skippedFiles `
                                        -LastWriteTime    $dtLastModified `
                                        -LastAccessTime   $dtLastAccessedDoc `
                                        -MetadataDuration $metadataDuration `
                                        -FileReadOnly     $fileReadOnly `
                                        -FilePath         $filePathLog `
                                        -FilePathProgress $filePathProgress `
                                        -StartTime        $startTime `
                                        -Format           $format `
                                        -FileSize         $filesize
                                    if ($action -eq "ContinueFile") { continue }
                                }
                                else {
                                    $doc = $app.Workbooks.Open($file, 0, $false)
                                }
                            }
                            catch {
                                $action = Handle-FileProcessingError `
                                    -ErrorRecord      $_ `
                                    -File             $file `
                                    -Status           ([ref]$status) `
                                    -App              $app `
                                    -Doc              $doc `
                                    -Item             $item `
                                    -ObjFile          $file `
                                    -ProcessedFiles   $processedFiles `
                                    -SkippedFiles     $skippedFiles `
                                    -LastWriteTime    $dtLastModified `
                                    -LastAccessTime   $dtLastAccessedDoc `
                                    -MetadataDuration $metadataDuration `
                                    -FileReadOnly     $fileReadOnly `
                                    -FilePath         $filePathLog `
                                    -FilePathProgress $filePathProgress `
                                    -StartTime        $startTime `
                                    -Format           $format `
                                    -FileSize         $filesize
                                if ($action -eq "ContinueFile") { continue }
                            }
                        }
                    }
 
                    # ----------------------------------------------------------
                    # PowerPoint
                    # ----------------------------------------------------------
                    "\.pptx|\.pptm|\.ppt" {
                        if ($isPasswordProtected -eq $false) {
                            try {
                                $app = New-Object -ComObject PowerPoint.Application
                                $app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
 
                                if ($app -eq $null) {
                                    Write-Error "PowerPoint COM object failed to initialise"
                                    $action = Handle-FileProcessingError `
                                        -ErrorRecord      $_ `
                                        -File             $file `
                                        -Status           ([ref]$status) `
                                        -App              $app `
                                        -Doc              $doc `
                                        -Item             $item `
                                        -ObjFile          $file `
                                        -ProcessedFiles   $processedFiles `
                                        -SkippedFiles     $skippedFiles `
                                        -LastWriteTime    $dtLastModified `
                                        -LastAccessTime   $dtLastAccessedDoc `
                                        -MetadataDuration $metadataDuration `
                                        -FileReadOnly     $fileReadOnly `
                                        -FilePath         $filePathLog `
                                        -FilePathProgress $filePathProgress `
                                        -StartTime        $startTime `
                                        -Format           $format `
                                        -FileSize         $filesize
                                    if ($action -eq "ContinueFile") { continue }
                                }
                                else {
                                    $doc = $app.Presentations.Open($file, $false, $false, $false)
                                    $doc.Saved = $false
                                }
                            }
                            catch {
                                $action = Handle-FileProcessingError `
                                    -ErrorRecord      $_ `
                                    -File             $file `
                                    -Status           ([ref]$status) `
                                    -App              $app `
                                    -Doc              $doc `
                                    -Item             $item `
                                    -ObjFile          $file `
                                    -ProcessedFiles   $processedFiles `
                                    -SkippedFiles     $skippedFiles `
                                    -LastWriteTime    $dtLastModified `
                                    -LastAccessTime   $dtLastAccessedDoc `
                                    -MetadataDuration $metadataDuration `
                                    -FileReadOnly     $fileReadOnly `
                                    -FilePath         $filePathLog `
                                    -FilePathProgress $filePathProgress `
                                    -StartTime        $startTime `
                                    -Format           $format `
                                    -FileSize         $filesize
                                if ($action -eq "ContinueFile") { continue }
                            }
                        }
                    }
 
                    default {
                        Write-Warning "No COM handler for extension: $($item.Extension)"
                    }
                }
 
                # --- Write properties (doc must be open and non-null to reach here) ---
                if ($isPasswordProtected -eq $false -and $doc -ne $null) {
 
                    $blProperty18Months  = [bool]((New-TimeSpan -Start $dtLastAccessedDoc -End (Get-Date)).TotalDays -gt 540)
                    $blProperty3Years    = [bool]((New-TimeSpan -Start $dtCreated         -End (Get-Date)).TotalDays -gt 1095)
                    $strProperty18Months = if ($blProperty18Months) { "True" } else { "False" }
                    $strProperty3Years   = if ($blProperty3Years)   { "True" } else { "False" }
 
                    Set-OfficeDocCustomProperty "OriginalPath"         $file                 $doc | Out-Null
                    Set-OfficeDocCustomProperty "LastAccessed18Months" $strProperty18Months  $doc | Out-Null
                    Set-OfficeDocCustomProperty "Created3Years"        $strProperty3Years    $doc | Out-Null
 
                    try {
                        $doc.Save()
                        $doc.Close()
                    }
                    catch {
                        $status = "Document cannot be saved"
                        $action = Handle-FileProcessingError `
                            -ErrorRecord      $_ `
                            -File             $file `
                            -Status           ([ref]$status) `
                            -App              $app `
                            -Doc              $doc `
                            -Item             $item `
                            -ObjFile          $file `
                            -ProcessedFiles   $processedFiles `
                            -SkippedFiles     $skippedFiles `
                            -LastWriteTime    $dtLastModified `
                            -LastAccessTime   $dtLastAccessedDoc `
                            -MetadataDuration $metadataDuration `
                            -FileReadOnly     $fileReadOnly `
                            -FilePath         $filePathLog `
                            -FilePathProgress $filePathProgress `
                            -StartTime        $startTime `
                            -Format           $format `
                            -FileSize         $filesize
                        if ($action -eq "ContinueFile") { continue }
                    }
 
                    if ($item.Extension -eq ".xls" -and $app -ne $null) {
                        $app.EnableEvents = $false
                    }
                    if ($app -ne $null) {
                        $app.Quit()
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
                        $app = $null
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers() 
                    }
 
                    $processed = $true
                }
            }
            catch {
                $isError = $true
                $action = Handle-FileProcessingError `
                    -ErrorRecord      $_ `
                    -File             $file `
                    -Status           ([ref]$status) `
                    -App              $app `
                    -Doc              $doc `
                    -Item             $item `
                    -ObjFile          $file `
                    -ProcessedFiles   $processedFiles `
                    -SkippedFiles     $skippedFiles `
                    -LastWriteTime    $dtLastModified `
                    -LastAccessTime   $dtLastAccessedDoc `
                    -MetadataDuration $metadataDuration `
                    -FileReadOnly     $fileReadOnly `
                    -FilePath         $filePathLog `
                    -FilePathProgress $filePathProgress `
                    -StartTime        $startTime `
                    -Format           $format `
                    -FileSize         $filesize
 
                Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file exception: $($_.Exception.Message)"
 
                if ($app -ne $null) {
                    try { $app.Quit() } catch {}
                    try { 
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
                        $app = $null
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers() 
                    } 
                    catch {}
                    $app = $null


                }
 
                if ($action -eq "ContinueFile") { continue }
            }
            finally {
                $endTime  = Get-Date
                $endTimeF = $endTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
 
                # Restore timestamps unconditionally
                try {
                    $item.LastWriteTime  = $dtLastModified
                    Start-Sleep -Milliseconds $metadataDuration
                    $item.LastAccessTime = $dtLastAccessedDoc
                    if ($fileReadOnly) { $item.IsReadOnly = $true }
                }
                catch {
                    $restoreMsg = $_.Exception.Message
                    $restoreHR  = if ($_.Exception.HResult) { '{0:X8}' -f $_.Exception.HResult } else { $null }
                    if ($restoreMsg -match 'being used by another process' -or $restoreHR -eq '80070020') {
                        Write-Warning "Timestamp restore skipped; file in use: $file ($restoreMsg)"
                        Add-ContentSafe -Path $skippedFiles -Value $file
                    }
                    else {
                        Write-Warning "Timestamp restore failed for $file : $restoreMsg (HR=$restoreHR)"
                    }
                }
 
                # Log outcome
                $logEntryProgress = @($file, $startTimeF, $endTimeF, $format, $filesize, $isPasswordProtected) -Join "|"
                Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress
                Add-ContentSafe -Path $processedFiles   -Value $file
 
                if ($isPasswordProtected) {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file skipped — password-protected"
                    Write-Output "$file skipped — password-protected"
                }
                elseif ($isError) {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file properties NOT updated — see error above"
                    Write-Output "$file failed — see log"
                }
                elseif ($processed) {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file properties updated"
                    Write-Output "$file properties updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                }
            }
 
        }) | Out-Null
 
        $ps.AddArgument($metadataDuration)  | Out-Null
        $ps.AddArgument($processedFiles)    | Out-Null
        $ps.AddArgument($skippedFiles)      | Out-Null
        $ps.AddArgument($filepathProgress)  | Out-Null
        $ps.AddArgument($format)            | Out-Null
        $ps.AddArgument($filePathLog)       | Out-Null
        $ps.AddArgument($filePath)          | Out-Null   # the actual file — was $fileToProcess (undefined)
 
        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }
 
    return $jobs
}

function Get-ApplicableFiles {
    [CmdletBinding()]
    [OutputType([System.Collections.ArrayList])]
    param (
        #[System.Collections.ArrayList]$Files,
        [string]$FolderName
    )

    # Write-Host "Folder Name --- line 487 inside fucntion -- is: $FolderName"

    # Validate folder
    if (-not (Test-Path -LiteralPath $FolderName -PathType Container)) {
        Write-Error "Folder '$FolderName' does not exist."
        return [System.Collections.ArrayList]::new()
    }


    # Allowed extensions
    $officeExtensions = @(
        ".doc", ".docx", ".docm",
        ".xls", ".xlsx", ".xlsm", ".xlsb",
        ".ppt", ".pptx", ".pptm",
        ".pdf"
    )

    $result = [System.Collections.ArrayList]::new()

    try {
        # Process files in current folder
        Get-ChildItem -LiteralPath $FolderName -File -ErrorAction Stop |
        Where-Object {
            #$_.LastAccessTime -lt (Get-Date).AddDays(-540) -and
            #$_.CreationTime   -lt (Get-Date).AddDays(-1095) -and
            ($officeExtensions -contains $_.Extension.ToLowerInvariant()) -and
            $_.Name.Substring(0,1) -ne '~' -and
            $_.Length -gt 0
        } |
        ForEach-Object {
            [void]$result.Add($_.FullName)
        }

        # Recurse into subfolders
        Get-ChildItem -LiteralPath $FolderName -Directory -ErrorAction Stop |
        ForEach-Object {
            Write-Verbose "Recursing into subfolder: $($_.FullName)"

            $childResults = Get-ApplicableFiles `
                -FolderName $_.FullName #`
                #-LastAccessedMonths $lastAccessedMonthsAbs `
                #-CreatedMonths $createdMonthsAbs

            foreach ($child in $childResults) {
                [void]$result.Add($child)
            }
        }
    }
    catch {
        Write-Error "Error processing folder '$FolderName': $($_.Exception.Message)"
    }

    return $result


}


function Execute_Tagging() {
	clear
	$newRun = $false
 
	# Define the folder path
	$FolderName = $DrivePath
	Write-Host "Folder Name is: $FolderName"
	# $FolderName = "C:\temp\Labelling"
    $targetFolder = ($($FolderName.TrimStart('\').Replace('\','_'))).Replace('C:','_')
	# Define the file collection location
	#$targetDir = "C:\Temp\Unstructured\$($FolderName.TrimStart('\').Replace('\','_'))"
    $targetDir = "$($env:LOCALAPPDATA)\Temp\Unstructured\$targetFolder"
	if (!(Test-Path $targetDir -PathType Container)) {
		New-Item -ItemType Directory -Path $targetDir
		$newRun = $true
	}
 
	$targetFiles = "$targetDir\FilesToScan.txt"
	if (!(Test-Path $targetFiles -PathType Leaf) ) {
		New-Item -Path $targetFiles -ItemType File -Force
		$newRun = $true
	}
	elseif ((Get-Item $targetFiles).Length -eq 0) {
		$newRun = $true
	}
 
	$scannedFiles = "$targetDir\FilesScanned.txt"
	if (! (Test-Path $scannedFiles -PathType Leaf)) {
		New-Item -Path $scannedFiles -ItemType File -Force
	}
    
    $skippedFiles = "$targetDir\FilesSkipped.txt"
    if (!(Test-Path $skippedFiles -PathType Leaf)) {
        New-Item -Path $skippedFiles -ItemType File -Force | Out-Null
    }

	$filesToScan =[System.Collections.ArrayList]::new()
	if ($newRun -eq $false) {
		$targetFilesList = (Get-Content -Path $targetFiles).Trim()
		if((Get-Item $scannedFiles).Length -ne 0) {
			$scannedFilesList = (Get-Content -Path $scannedFiles).Trim()
			foreach ($targetFile in $targetFilesList) {
				if($scannedFilesList.contains($targetFile)) {
				Write-Host "$targetFile has already been scanned"
				}
				else {
					$filesToScan.Add($targetFile)
				}
			}
		}
		else {
			$filesToScan = $targetFilesList
		}
 
		if(($filesToScan).Count -ne 0) {
			$filesToScanUnique = $filesToScan | sort -Unique
			$continue = $true
		}
	}
	else {
		Write-Host "Retrieving applicable files with Get-ApplicableFiles ... "
		# Get the applicable files
		if($filesToScanUnique -eq $null) {
			$filesToScanUnique=[System.Collections.ArrayList]::new()
		}
		$filesToScan = Get-ApplicableFiles -FolderName $FolderName
		if($filesToScan.Count -ne 0) {
			$filesToScanUnique = $filesToScan | sort -Unique
			foreach ($file in $filesToScanUnique) {
				Add-ContentSafe -Path $TargetFiles -Value $file
			}
			$continue = $true
		}
	}
 
	if ($continue -eq $true) {
		Write-Host "Processing files with Update-FileAgeProperties ... "
		# Execute the update process on retrieved files
		Update-FileAgeProperties -Files $filesToScanUnique -ProcessedFiles $scannedFiles
		$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
		$filenameToScan = "$($timestamp)_FilesToScan.txt"
		$filenameScanned = "$($timestamp)_FilesScanned.txt"
 
		Rename-Item -Path $targetFiles -NewName $filenameToScan
		Rename-Item -Path $scannedFiles -NewName $filenameScanned
	}   
	else {
		Write-Host "No files to process"
	}
 
 
	# Invoke garbage collection to clear processes
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
}




function Start-KillProcessMonitor {
    param(
        [int]$MaxRuntimeSeconds = 60,
        [int]$CheckIntervalSeconds = 15,
        [string]$LogPath = "C:\temp\KillProcess.log",
        [switch]$ShowWindow,   # show a console window
        [switch]$NoExit        # keep it open (for debugging)
    )

    # --- Build the monitor script that runs in the external PowerShell process ---
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
            Write-Host "Monitor error: $($_.Exception.Message)" -ForegroundColor Red
        }

        Start-Sleep -Seconds $CheckIntervalSeconds
    }
}
catch {
    "`$(Get-Date -Format o) | Monitor fatal error: `$(`$_.Exception.Message)" | Out-File -Append '$LogPath'
    Write-Host "Monitor fatal error: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Press Enter to close the monitor..."
}
finally {
    "`$(Get-Date -Format o) | External Kill-Process exiting" | Out-File -Append '$LogPath'
}
"@

    # --- Resolve a console host (prefer Windows PowerShell console for ISE) ---
    $exe = (Get-Command powershell -ErrorAction SilentlyContinue).Source
    if (-not $exe) { $exe = (Get-Command pwsh -ErrorAction SilentlyContinue).Source }
    if (-not $exe) { throw "Neither 'powershell' nor 'pwsh' was found on PATH." }

    # --- Build argument list (avoid using $args because it is an automatic variable) ---
    $procArgs = @('-NoProfile','-ExecutionPolicy','Bypass')

    # ShowWindow often implies interactive debugging; add -NoExit if requested
    if ($NoExit -or $ShowWindow) { $procArgs += '-NoExit' }

    # Prefer EncodedCommand to avoid quoting/length issues
    $bytes     = [System.Text.Encoding]::Unicode.GetBytes($monitorScript)
    $b64       = [Convert]::ToBase64String($bytes)
    $procArgs += @('-EncodedCommand', $b64)

    # --- Launch (window visible if ShowWindow) ---
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
    # For inspection if needed:
    Write-Verbose ("Args: {0}" -f ($procArgs -join ' '))

    return $proc
}


function Stop-KillProcessMonitor {
    param([Parameter(Mandatory)][System.Diagnostics.Process]$MonitorProcess)
    if ($MonitorProcess -and -not $MonitorProcess.HasExited) {
        Stop-Process -Id $MonitorProcess.Id -Force -ErrorAction SilentlyContinue
    }
}



function Invoke-ExecuteTaggingSafely {

    # Start the external monitor process (hidden)
    #$monitorProc = Start-KillProcessMonitor -MaxRuntimeSeconds 60 -CheckIntervalSeconds 15 -LogPath "C:\temp\KillProcess.log" -ShowWindow -NoExit
    $monitorProc = Start-KillProcessMonitor -MaxRuntimeSeconds 60 -CheckIntervalSeconds 15 -LogPath "$($env:LOCALAPPDATA)\Temp\KillProcess.log" -ShowWindow -NoExit

    $taggingFailed = $false

    try {
        Execute_Tagging
    }
    catch {
        $taggingFailed = $true
        Write-Warning ("Execute_Tagging failed (continuing): {0}" -f $_.Exception.Message)
        # Do NOT rethrow if you want the script to continue
        write-host $_.Exception.Message
        if ($_.Exception.Message -match "The RPC server is unavailable" -or $_.Exception.HResult -eq 0x800706BA) {
            Write-Warning "RPC failure"
            $global:RpcFailureDetected = $true
        }
    }

    finally {
        if ($monitorProc) {

            # Wait for any Office processes still alive after Execute_Tagging returns.
            # COM automation in runspaces can leave WINWORD/EXCEL/POWERPNT briefly
            # alive after the PowerShell job completes, so we drain them here
            # before killing the monitor that would have cleaned them up.
            $officeProcesses  = @('WINWORD', 'EXCEL', 'POWERPNT')
            $drainTimeoutSecs = 1200    # give up after this long regardless
            $pollIntervalMs   = 2000
            $elapsed          = 0

            Write-Host "Waiting for Office processes to exit before stopping monitor..."

            while ($elapsed -lt ($drainTimeoutSecs * 1000)) {
                $remaining = $officeProcesses | Where-Object {
                    Get-Process -Name $_ -ErrorAction SilentlyContinue
                }

                if (-not $remaining) {
                    Write-Host "All Office processes exited."
                    break
                }

                Write-Host "Still running: $($remaining -join ', ') — waiting..."
                Start-Sleep -Milliseconds $pollIntervalMs
                $elapsed += $pollIntervalMs
            }

            if ($elapsed -ge ($drainTimeoutSecs * 1000)) {
                Write-Warning "Drain timeout reached ($drainTimeoutSecs s) — Office processes may still be running. Stopping monitor anyway."
            }

            Stop-KillProcessMonitor -MonitorProcess $monitorProc
        }
    }    

}

    # Normalize exit code if you're in a pipeline that treats non-zero as failure
    if (-not $global:RpcFailureDetected) {
        $global:LASTEXITCODE = 0
    }


$global:RpcFailureDetected = $false
Invoke-ExecuteTaggingSafely -Verbose

if ($global:RpcFailureDetected) {
    Write-Host "##[error]Restart required due to RPC failure"
    exit 42
}

<#
$fileToProcess = "C:\Users\keval\OneDrive\Documents\Book1.xlsx"
Add-Content "$($env:LOCALAPPDATA)\temp\debug.txt" "BEFORE XML: $fileToProcess"

$check = Set-OpenXmlProperties -FilePath $fileToProcess -Properties @{
    Test = "True"
}

Add-Content -Path "$($env:LOCALAPPDATA)\temp\debug.txt" "AFTER XML: $fileToProcess $($check)"
#>
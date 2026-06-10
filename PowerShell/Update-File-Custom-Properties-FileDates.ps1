# Hints taken with thanks from:
# https://stackoverflow.com/questions/51248195/check-if word-file-is-password-protected-in-powershell
# https://stackoverflow.com/questions/53147328/word bypass-password-protected-files

Function IsOfficeFilePasswordProtected {   
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$officeFile
    )
 
	if (!(Test-Path -Path $officeFile -PathType Leaf) ) {
		Write-Error "File $officeFile does not exist"
		return $null
	}

    $extension = (Get-Item $officeFile).Extension
	if ($extension.Length -eq 5) {
		$hasPassword = [bool](Test-OfficeEncrypted -officeFile $officeFile).IsEncrypted
	} 
	else {
        switch -Exact ($extension.Substring(1,2)){
            "pp" {
            	$hasPassword = Test-Ppt2003HasOpenPassword -officeFile $officeFile
            }
            default {
		        $header = Get-Content $officeFile -Encoding Unicode -Total 1
		        if (!($header -ne $null) -or !($header -notmatch "Microsoft Enhanced Cryptographic Provider")) {
			        $hasPassword = $true
		        } 
            }
        }
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

	return $hasPassword
	
}

Function Test-OfficeEncrypted {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$officeFile
    )
 
    # Read file header
    $fs = [System.IO.File]::OpenRead($officeFile)
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
            [System.IO.Compression.ZipFile]::OpenRead($officeFile).Dispose()
            $fs.Close()
 
            return [PSCustomObject]@{
                Path = $officeFile
                IsEncrypted = $false
                Reason = "Normal OOXML ZIP (not encrypted)"
            }
        }
        catch {
            $fs.Close()
            return [PSCustomObject]@{
                Path = $officeFile
                IsEncrypted = $false
                Reason = "Not OLE and not ZIP and not an encrypted Office document"
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
    $fs.Close()
    $hasEncryptedPackage = $entries -contains "EncryptedPackage"
    $hasEncryptionInfo   = $entries -contains "EncryptionInfo"
 
    if ($hasEncryptedPackage -and $hasEncryptionInfo) {
        return [PSCustomObject]@{
            Path = $officeFile
            IsEncrypted = $true
            Reason = "Encrypted (contains EncryptionInfo + EncryptedPackage streams)"
        }
    }
    elseif ($hasEncryptedPackage -or $hasEncryptionInfo) {
        return [PSCustomObject]@{
            Path = $officeFile
            IsEncrypted = $true
            Reason = "Partially encrypted (contains EncryptionInfo + EncryptedPackage streams)"
        }
    }
    else {
        return [PSCustomObject]@{
            Path = $officeFile
            IsEncrypted = $false
            Reason = "OLE file but missing encryption streams"
        }
    }
}

Function Test-Ppt2003HasOpenPassword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$officeFile
    )
 
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
        $pres = $app.Presentations.Open($officeFile, $true, $false, $false)

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

    # --- Unwrap COMException if needed ---
    if ($ex -is [System.Runtime.InteropServices.COMException]) {
        $hresult = $ex.HResult
    }
    elseif ($ex -is [System.Management.Automation.MethodInvocationException] -and
            $ex.InnerException -is [System.Runtime.InteropServices.COMException]) {
        $hresult = $ex.InnerException.HResult
        $msg = $ex.InnerException.Message
    }

    # --- COM / RPC classification ---
    if ($hresult -in 0x800706BE,0x80010105,0x800706BA) {
        Write-Warning ("RPC/COM error on '{0}' (0x{1:X8}): {2}. Continuing." -f $File, $hresult, $msg)
        $passwordProtected = $true
        $Status.Value.Failed++
        return [FileErrorAction]::ContinueFile
    }

    if ($hresult) {
        Write-Warning ("Unhandled COM error on '{0}' (0x{1:X8}): {2}. Continuing." -f $File, $hresult, $msg)
        $Status.Value.Failed++
        return [FileErrorAction]::ContinueFile
    }

    # --- COM error: treat as password-protected ---
    if($passwordProtected) {
        Write-Warning ("Failed on '{0}': {1}. Treating file as password-protected." -f $File, $msg)
        $message = "File is password-protected"
    }
    else
    {
        Write-Warning ("Failed on '{0}': {1}. Can't process file." -f $File, $msg)
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
    if (-not $Item) {
        $Item = Get-Item -LiteralPath $ObjFile
    }

    Restore-FileMetadata -Item $Item -LastWriteTime $LastWriteTime -LastAccessTime $LastAccessTime `
        -MetadataDuration $MetadataDuration -RestoreReadOnly:$FileReadOnly `
        -ObjFile $ObjFile -LogPath $FilePath | Out-Null


    # Logging
    Write-Log -filePath $FilePath -objFile $ObjFile -message $message
    Write-LogProcess -filePath $FilePathProgress `
                     -objFile $ObjFile `
                     -startTime $StartTime `
                     -fileFormat $Format `
                     -fileSize $FileSize `
                     -isPasswordProtected $true

    $Status.Value.Failed++
    return [FileErrorAction]::ContinueFile
}


function Write-Log {
	Param(
		[Parameter (Mandatory=$true)]
		[string] $filePath,
		[Parameter (Mandatory=$true)]
		[string] $objFile,
		[Parameter (Mandatory=$true)]
		[string] $message
	)
	if (! (Test-Path -Path $filePath -PathType Leaf) ) {
		Write-Error "File $filePath does not exist"
	}		
	$logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $objFile - $message"
	Add-Content -Path $filePath -Value $logEntry
}

function Write-LogProcess {
	Param(
		[Parameter (Mandatory=$true)]
		[string] $filePath, 
		[Parameter (Mandatory=$true)]
		[string] $objFile, 
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
	$logEntryProgress = @($objFile, $startTime, $endTimeF, $fileFormat, $fileSize, $isPasswordProtected) -Join "|"
	Add-Content -Path $filePath -Value $logEntryProgress
}
	
function IsFile-Open {
	[OutputType([boolean])]
	Param(
		[Parameter (Mandatory=$true) ]
		[string] $FileName
    )
	try {
        $isReadOnly = [bool](Get-Item $FileName).IsReadOnly
		if( $isReadOnly -eq $false) {
            $stream = [System.IO.File]::Open( $FileName,[System.IO.FileMode]::Open,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)
            $stream.Close()
            return $false
        }
    }
    catch {
        return $true
    }

}
 
#Taken with thanks from https://www.rlvision.com/blog/read-write-ms-office-custom-properties-with-powershell
#Function to set the Custom Document Property of an MS Office file (passed in as COM object parameter
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
			[System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | out-null
		}
		catch [system.exception] {
			$propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
			[System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
			[System.__ComObject].InvokeMember("add", $binding:: InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
		}
		return $true
	}
	catch {
		return $false
	}
}

# Task 1 (Keval, 2026-04-21): mirror of Set-OfficeDocCustomProperty that returns the value of
# a custom document property, or $null if the property is not set.
# Item() is fetched via the same reflection pattern as the setter (so the review stays
# consistent), but Value is read via direct PowerShell property access - InvokeMember on
# a late-bound COM property getter can silently return the wrong thing.
function Get-OfficeDocCustomProperty {
	[OutputType([string])]
	Param(
		[Parameter (Mandatory=$true) ]
		[string] $PropertyName,
		[Parameter (Mandatory=$true) ]
		[System.__ComObject] $Document
	)
	try {
		$customProperties = $Document.CustomDocumentProperties
		$binding = "System.Reflection.BindingFlags" -as [type]
		try {
			$propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
			if ($null -eq $propertyObject) { return $null }
			$value = $propertyObject.Value
			if ($null -eq $value) { return $null }
			return [string]$value
		}
		catch {
			# Item() throws when the named property does not exist on the document
			return $null
		}
	}
	catch {
		return $null
	}
}

#Function to loop through a collection of files, check their age and create/update custom document properties
# Restore LastWriteTime / LastAccessTime / IsReadOnly on a file, retrying briefly
# when Office COM or a network share has not yet released its lock. A failure here
# must not abort the whole tagging run, so the final exception is logged and
# swallowed.
function Restore-FileMetadata {
	param(
		[Parameter(Mandatory)]$Item,
		[Parameter(Mandatory)][datetime]$LastWriteTime,
		[Parameter(Mandatory)][datetime]$LastAccessTime,
		[int]$MetadataDuration = 100,
		[bool]$RestoreReadOnly = $false,
		[string]$ObjFile,
		[string]$LogPath,
		[int]$MaxAttempts = 5,
		[int]$BaseDelayMs = 500
	)

	$lastException = $null
	for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
		try {
			$Item.LastWriteTime = $LastWriteTime
			Start-Sleep -Milliseconds $MetadataDuration
			$Item.LastAccessTime = $LastAccessTime
			if ($RestoreReadOnly) { $Item.IsReadOnly = $true }
			return $true
		}
		catch {
			$lastException = $_
			Start-Sleep -Milliseconds ($BaseDelayMs * $attempt)
		}
	}

	$msg = "could not restore file metadata after $MaxAttempts attempts: $($lastException.Exception.Message)"
	if ($LogPath -and $ObjFile -and (Test-Path -LiteralPath $LogPath -PathType Leaf)) {
		try { Write-Log -filePath $LogPath -objFile $ObjFile -message $msg } catch {}
	}
	Write-Warning ("{0} - {1}" -f $ObjFile, $msg)
	return $false
}


function Update-FileAgeProperties{
	param (
		[System.Collections.ArrayList] $Files,
		[String] $processedFiles

	) #pass in an existing collection object and list of processed files

	#Ensure the output file exists

	if (!(Test-Path -Path $processedFiles -PathType Leaf)) {
		Write-Error "File $ProcessedFiles does not exist."
		return
	}
    
    $pdfRun = $true

	#Python dependencies for PDF updates
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $PythonPath = $pythonCmd.Source
        $pythonVersion = & $pythonCmd.Source --version 2>&1
        
        if ($pythonVersion -match '(\d+)\.(\d+)') {
            $majorMinor = "$($matches[1]).$($matches[2])"
        }

    } elseif (Get-Command py -ErrorAction SilentlyContinue) {
        $PythonPath =  (& py -0p)
        $majorMinor = ($pythonPath -split "`r?`n") |
            Where-Object { $_ -match '^\s*-V:(\d+\.\d+)\s*\*' } |
            ForEach-Object { $matches[1] } |
            Select-Object -First 1

    } else {
        Write-Host "Python is not installed or not in PATH. PDF tagging cannot be completed"
        $pdfRun = $false
    }

    $venvDir = Join-Path $PSScriptRoot '.venv'  # or use (Get-Location) if not in a script
    $venvPython = Join-Path $venvDir 'Scripts\python.exe'  # Windows
    $venvActivate = Join-Path $venvDir 'Scripts\python.exe'
    # Get the path to a file in the same folder as this script
    $ScriptPath = Join-Path $PSScriptRoot 'update_pdf_properties.py'

    if (Test-Path $ScriptPath) {
        Write-Host "Found file at $ScriptPath"
    } else {
        Write-Host "Python is not stored in expected location. PDF tagging cannot be completed"
        $pdfRun = $false
    }

    if ($pdfRun) {
        try {
            # Create venv next to the PowerShell + Python scripts if it isn't already there.
            if (-not (Test-Path $venvPython)) {
                py -$majorMinor -m venv $venvDir
            }

            # Install dependencies INTO the venv (not the system Python). Using the venv's
            # own interpreter guarantees pip writes to .venv\Lib\site-packages.
            # pip itself is provisioned automatically by `python -m venv`, so we don't
            # try to upgrade it here (PowerShell parses `--upgrade` as one of its own
            # parameters in some host versions).
            if (Test-Path $venvPython) {
                & $venvPython -m pip install pypdf cryptography | Out-Null
            }
            else {
                throw "venv interpreter not found at $venvPython after creation"
            }
        }
        catch {
            Write-Error "Failed to install Python virtual environment. PDF tagging cannot be completed. $($_.Exception.Message)"
            $pdfRun = $false
        }
    }

    
    # Adjust name/path as needed (e.g., ".venv" or "venv")
    # For PowerShell 7 on Linux/macOS, use 'bin/python' instead

    if (Test-Path $venvPython) {
        Write-Host "Venv exists at: $venvDir"
        # Use the venv's isolated Python for actual PDF tagging. This sidesteps the
        # Windows App Execution Alias stub (WindowsApps\python.exe) that Get-Command
        # can return and that produces 'can't open file ... WindowsApps\python.exe
        # [Errno 22]' when used as the interpreter for update_pdf_properties.py.
        $PythonPath = $venvPython
    } else {
        Write-Host "Venv not found (expected $venvPython) - PDF tagging disabled."
        $pdfRun = $false
    }

	# Define output log file
	$filePathBase = "$Env:LOCALAPPDATA\Temp\Unstructured"
    if (!(Test-Path $filePathBase -PathType Container)) {
        New-Item -Path $filePathBase -ItemType Container -Force 
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$filename = "$($timestamp)_AddPropertiesLog.txt"
	$filenameProgress = "$($timestamp)_AddPropertiesStatus.txt"
	$filepath = "$filePathBase\$filename"
	$filepathProgress = "$filePathBase\$filenameProgress"
    $metadataDuration = 100

	$logEntryProgress = @("Filename", "StartTime", "EndTime", "Format", "Filesize", "PasswordProtected") -Join "|"
	Add-Content -Path $filepathProgress -Value $logEntryProgress
	# Create the log file
	New-Item -Path $filepath -ItemType File -Force
	# Per-run counters passed by [ref] into Handle-FileProcessingError. Must exist
	# before the loop or `[ref]$status` throws "cannot be applied to a variable
	# that does not exist" the first time a file fails to open.
	$status = [pscustomobject]@{ Failed = 0; Succeeded = 0; Skipped = 0 }

	# Loop through each file in collection parameter
	foreach ($objFile in $Files) {
		$processed = $false
		$success = $false
		$item = Get-Item $objFile
		$startTime = Get-Date
		$startTimeF = $startTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		$filesize = $item.Length
		# Get file time metadata
		$dtLastAccessedDoc = $item.LastAccessTime
		$dtCreated = $item.CreationTime
		$dtLastModified = $item.LastWriteTime
		# Write to the log
		Write-Log -filePath $filePath -objFile $objFile -message "preparing file for property update"
		#Test if the file is already open
		$isFileOpen = IsFile-Open -Filename $objFile
		# If the file is open, skip over and update the log
		if ($isFileOpen) {
			Write-Log -filePath $filePath -objFile $objFile -message "file currently open, properties not set"
			Write-Output "$objFile - file currently open, properties not set"
		}

		# If the file is closed, open the file
		else {
			$officeApp = $false
			$pdfApp = $false
			$fileReadOnly = $false
		}
		if((Get-Item $objFile).IsReadOnly -eq $true) {
			(Get-Item $objFile).IsReadOnly = $false
			$fileReadOnly = $true
		}

		if($item.Extension -ne ".pdf") {
    		$isPasswordProtected = IsOfficeFilePasswordProtected -officeFile $objFile
        }
		try{
			switch -regex ($item.Extension) {
				".docx|.docm|.doc" {
					if($isPasswordProtected -eq $false) {
						try {
							$app = New-Object -ComObject Word.Application
							$app.Visible = $false
							$doc = $app.Documents.Open($objFile)
							$doc.Saved = $false
							$format = ".doc"
							$officeApp = $true
						}

                        catch {
                            $action = Handle-FileProcessingError `
                                -ErrorRecord $_ `
                                -File $file `
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

                            if ($action -eq "ContinueFile") {
                                continue
                            }
                        }

					}
				}

			    ".xlsx|.xlsm|.xls|.xlsb" {
				    if ($isPasswordProtected -eq $false) {
					    try {
						    $app = New-Object -ComObject Excel.Application
						    $app.Visible = $false
						    $app.DisplayAlerts = $false
						    $app.EnableEvents = $false
						    $doc = $app.Workbooks. Open($objFile, $false)
						    $doc.CheckCompatibility = $False
						    $doc.Saved = $false
						    $format = ".xls"
						    $officeApp = $true
					    }

                        catch {
                            $action = Handle-FileProcessingError `
                                -ErrorRecord $_ `
                                -File $file `
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

                            if ($action -eq "ContinueFile") {
                                continue
                            }
                        }

				    }
			    }
			    ".pptx|.pptm|.ppt" {
				    if($isPasswordProtected -eq $false) {
					    $app = New-Object -ComObject PowerPoint.Application
					    try {
						    $doc = $app.Presentations.Open($objFile, $false, $false, $false)
						    $doc.Saved = $false
						    $format = ".ppt"
						    $officeApp = $true
					    }

                        catch {
                            $action = Handle-FileProcessingError `
                                -ErrorRecord $_ `
                                -File $file `
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

                            if ($action -eq "ContinueFile") {
                                continue
                            }
                        }

				    }
			    }

			    ".pdf$" {
				    $format = ".pdf"
				    $pdfApp = $True
			    }

			    default { Write-Host "No match found for extension: $($item.Extension)" }
		    }
        }

		# If the file raises an error when it is opened, skip over and update the log
		catch {
			$errortext = $($_.Exception.Message)
			Write-Log -filePath $filePath -objFile $objFile -message "file not updatable, properties not set. Error: $errortext"
			if ($format -eq ".xls" -and $app -ne $null){
				$app.EnableEvents = $false
			}
			Write-Output "$objFile - file not updatable, properties not set"
			Write-Output "Detailed Error: $($_.Exception)"

			# Guard COM teardown: if Office already died (e.g. the file failed to open
			# or the process was killed), Close()/Quit() throw "RPC server unavailable",
			# which would otherwise escape the loop and abort the whole run.
			if ($doc -ne $null) {
				try { $doc.Close() } catch { }
			}

			if ($app -ne $null) {
				try { $app.Quit() } catch { }
				try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
			}
			$doc = $null
			$app = $null

			# Update file timestamps
			if($item -eq $null) {
    			$item = Get-Item $objFile
            }
			Restore-FileMetadata -Item $item -LastWriteTime $dtLastModified -LastAccessTime $dtLastAccessedDoc `
				-MetadataDuration $metadataDuration -RestoreReadOnly:$fileReadOnly `
				-ObjFile $objFile -LogPath $filePath | Out-Null
		}

		# If the file has been opened
		if ($doc -ne $null -or $pdfApp -eq $true){
			# Task 2: determine bucket from parent-folder location, read (or initialise)
			# the RetentionStartDate custom property, compute OutOfRetention, then write
			# the four properties Purview DLP consumes.
			$bucket = Get-RetentionBucket -FilePath $objFile -Buckets $script:RetentionBuckets

			if (-not $bucket) {
				# File sits outside any known retention bucket - leave it untouched but
				# still close the app cleanly so the timestamp restore below works.
				Write-Log -filePath $filePath -objFile $objFile -message "file not in a recognised retention bucket; skipping retention tagging"
				Write-Output "$objFile - skipped (no matching retention bucket)"
				if ($officeApp -eq $true -and $doc -ne $null) {
					try { $doc.Saved = $true } catch {}
					try { $doc.Close() } catch {}
					if ($format -eq ".xls" -and $app -ne $null) { $app.EnableEvents = $false }
					if ($app -ne $null) {
						try { $app.Quit() } catch {}
						[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
						$app = $null
					}
				}
			}
			else {
				# Read an existing RetentionStartDate, if one is already on the file.
				# PDFs cannot yet be read by the Python helper (TODO below) - they will
				# be initialised on every pass until update_pdf_properties.py gains a read mode.
				$retentionStartRaw = $null
				if ($officeApp -eq $true) {
					$retentionStartRaw = Get-OfficeDocCustomProperty -PropertyName "RetentionStartDate" -Document $doc
				}

				$retentionStart = $null
				if ($retentionStartRaw) {
					try {
						$retentionStart = [datetime]::ParseExact($retentionStartRaw, 'yyyy-MM-dd', $null)
					}
					catch {
						Write-Log -filePath $filePath -objFile $objFile -message "RetentionStartDate '$retentionStartRaw' unparseable; resetting to today"
						$retentionStart = (Get-Date).Date
					}
				}
				else {
					# No existing RetentionStartDate. Try the HR leavers list first
					# (option 3 from the 2026-05-01 catch-up); if no match, fall
					# back to today (Keval's "time of run as proxy" - acceptable
					# for files that arrive after first run).
					$leaverMatch = $null
					if ($script:LeaversIndex) {
						$leaverMatch = Get-LeaverForFile -FilePath $objFile -LeaversIndex $script:LeaversIndex
					}
					if ($leaverMatch) {
						$retentionStart = $leaverMatch.LeaveDate
						Write-Log -filePath $filePath -objFile $objFile `
							-message ("matched leaver {0} (ID '{1}'); using leave date {2:yyyy-MM-dd}" -f $leaverMatch.EmployeeName, $leaverMatch.EmployeeID, $leaverMatch.LeaveDate)
					}
					else {
						$retentionStart = (Get-Date).Date
					}
				}

				$retentionStartStr    = $retentionStart.ToString('yyyy-MM-dd')
				$isOutOfRetention     = ($retentionStart.AddYears($bucket.Years) -lt (Get-Date).Date)
				$outOfRetentionStr    = if ($isOutOfRetention) { "True" } else { "False" }
				$retentionCategoryStr = $bucket.Name

				if ($officeApp -eq $true) {
					[void](Set-OfficeDocCustomProperty "OriginalPath"       $objFile              $doc)
					[void](Set-OfficeDocCustomProperty "RetentionStartDate" $retentionStartStr    $doc)
					[void](Set-OfficeDocCustomProperty "OutOfRetention"     $outOfRetentionStr    $doc)
					[void](Set-OfficeDocCustomProperty "RetentionCategory"  $retentionCategoryStr $doc)
					$doc.Save()
					$doc.Close()
					if ($format -eq ".xls" -and $app -ne $null) {
						$app.EnableEvents = $false
					}
					$app.Quit()
					[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
					$app = $null
					$processed = $true
				}
				elseif ($pdfApp -eq $true -and $pdfRun -eq $true) {
					# TODO: update_pdf_properties.py needs a read mode so PDFs can re-use an
					# existing RetentionStartDate across runs. Until then, every pass writes
					# today's date as the start - which is fine for the first ever run but
					# means PDFs already in retention-land will have their clock reset.
					$pdfProps = [ordered]@{
						"OriginalPath"       = $objFile
						"RetentionStartDate" = $retentionStartStr
						"OutOfRetention"     = $outOfRetentionStr
						"RetentionCategory"  = $retentionCategoryStr
					}
					foreach ($name in $pdfProps.Keys) {
						$value = $pdfProps[$name]
						switch (& $PythonPath $ScriptPath $name $value $objFile) {
							1  { $isPasswordProtected = $true }
							2  { $isPasswordProtected = $true }
							-1 { $isError = $true }
							default {
								$isPasswordProtected = $false
								$isError = $false
							}
						}
					}
				}
			}

			if ($isPasswordProtected -eq $true) {
				#Write to log that file has been updated
				Restore-FileMetadata -Item $item -LastWriteTime $dtLastModified -LastAccessTime $dtLastAccessedDoc `
					-MetadataDuration $metadataDuration -RestoreReadOnly:$fileReadOnly `
					-ObjFile $objFile -LogPath $filePath | Out-Null

				Write-Log -filePath $filePath -objFile $objFile -message "file is password-protected"
				Write-LogProcess -filePath $filePathProgress -objFile $objFile -startTime $startTime -fileFormat $format -fileSize $filesize -isPasswordProtected $isPasswordProtected

				#Write to processed list that file has been updated
				Add-Content -Path $processedFiles -Value $objFile
				Write-Output "$objFile - file is password-protected"
			} 
			elseif ($isPasswordProtected -eq $false -and $isError -eq $false) {
					$processed = $true
			}

			if ($processed -eq $true) {
				# Update file timestamps
				Restore-FileMetadata -Item $item -LastWriteTime $dtLastModified -LastAccessTime $dtLastAccessedDoc `
					-MetadataDuration $metadataDuration -RestoreReadOnly:$fileReadOnly `
					-ObjFile $objFile -LogPath $filePath | Out-Null
				$success = $true
				$endTime = Get-Date
				$endTimeF = $endTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				#Write to log that file has been updated
				Write-Log -filePath $filePath -objFile $objFile -message "properties updated"
				#Write to data that file has been updated
				Write-LogProcess -filePath $filePathProgress -objFile $objFile -startTime $startTime -fileFormat $format -fileSize $filesize -isPasswordProtected $isPasswordProtected
				#Write to processed list that file has been updated
				Add-Content -Path $processedFiles -Value $objFile
				Write-Output "$objFile - properties updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
			}
		}
	}
}



# Retention buckets. Folder names under the labelling root map to retention periods in years.
# Get-RetentionBucket walks up each file's parents and returns the first match.
$script:RetentionBuckets = [ordered]@{
    'PAST EMPLOYEES - FILES' = 7
}

# Walk up a file's parent folders until we find a folder name that matches one of the
# configured retention buckets. Returns an object with Name (the bucket folder name) and
# Years (retention period). Returns $null if the file sits outside any known bucket.
function Get-RetentionBucket {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string] $FilePath,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Buckets
    )

    $dir = Split-Path -Path $FilePath -Parent
    while ($dir) {
        $name = Split-Path -Path $dir -Leaf
        if ($Buckets.Contains($name)) {
            return [pscustomobject]@{
                Name  = $name
                Years = [int]$Buckets[$name]
            }
        }
        $parent = Split-Path -Path $dir -Parent
        if ([string]::IsNullOrEmpty($parent) -or $parent -eq $dir) { break }
        $dir = $parent
    }
    return $null
}

# Task 2 rewrite. Replaces the old age-threshold-based filter. A file is "applicable"
# if it (a) has a supported Office/PDF extension AND (b) sits under a recognised
# retention bucket folder (6/7/10 year). The actual in-or-out-of-retention decision
# happens per-file inside Update-FileAgeProperties, not here.
function Get-ApplicableFiles {
    [CmdletBinding()]
    [OutputType([System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory)]
        [string] $FolderName,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary] $Buckets
    )

    if (-not (Test-Path -LiteralPath $FolderName -PathType Container)) {
        Write-Error "Folder '$FolderName' does not exist."
        return [System.Collections.ArrayList]::new()
    }

    $officeExtensions = @(
        ".doc", ".docx", ".docm",
        ".xls", ".xlsx", ".xlsm", ".xlsb",
        ".ppt", ".pptx", ".pptm",
        ".pdf"
    )

    $result = [System.Collections.ArrayList]::new()

    try {
        Get-ChildItem -LiteralPath $FolderName -File -Recurse -ErrorAction Stop |
        Where-Object {
            # Skip Office owner/lock files (~$name.docx) - they aren't real documents,
            # they fail to open via COM, and that failure triggers the error cascade.
            (-not $_.Name.StartsWith('~$')) -and
            ($officeExtensions -contains $_.Extension.ToLowerInvariant()) -and
            (Get-RetentionBucket -FilePath $_.FullName -Buckets $Buckets)
        } |
        ForEach-Object {
            [void]$result.Add($_.FullName)
        }
    }
    catch {
        Write-Error "Error processing folder '$FolderName': $($_.Exception.Message)"
    }

    return $result
}


# ---- Task 3 (Arven, 2026-05-05; spec from catch-up call 2026-05-01) ----
# HR leavers-list match. When HR provides Leavers.csv with EmployeeName,
# EmployeeID, LeaveDate, we resolve files to a leaver and use the leave date
# as RetentionStartDate instead of stamping today. Two-hashtable lookup
# (by ID + by name) per Keval's suggestion on the call. Falls back to
# today when no CSV is loaded or no row matches, so the existing test-bed
# pipeline is unaffected if HR hasn't shipped the list yet.

# Tokenise a name/path candidate into lowercase alphanumeric chunks. Used
# for both ID lookup (token equality) and name match (subset comparison).
function ConvertTo-LeaverTokens {
    [CmdletBinding()]
    [OutputType([string[]])]
    param([string] $Text)
    if (-not $Text) { return @() }
    return @(($Text -split '[^A-Za-z0-9]+') |
             Where-Object { $_.Length -gt 0 } |
             ForEach-Object { $_.ToLowerInvariant() })
}

# Read HR's leavers CSV and build two indices: one keyed by lower-cased
# Employee ID, one keyed by sorted-and-joined name tokens. Returns $null
# if the CSV does not exist - callers should treat that as "leavers match
# disabled" and fall through to the today-as-start behaviour.
function Import-LeaversList {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)][string] $CsvPath
    )

    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        Write-Warning "Leavers CSV '$CsvPath' not found - leaver matching disabled."
        return $null
    }

    # Expected columns: EmployeeName, EmployeeID, LeaveDate. Date format
    # ideally yyyy-MM-dd; we accept several common formats since HR may
    # export from different systems (Excel, Sage, etc.).
    $rows = Import-Csv -LiteralPath $CsvPath
    $byId   = @{}
    $byName = @{}
    $ambiguous = New-Object System.Collections.Generic.HashSet[string]
    $acceptedFormats = @('yyyy-MM-dd','dd/MM/yyyy','MM/dd/yyyy','d MMM yyyy','dd-MM-yyyy','yyyy/MM/dd')

    foreach ($row in $rows) {
        $id      = if ($row.EmployeeID)   { "$($row.EmployeeID)".Trim()   } else { '' }
        $name    = if ($row.EmployeeName) { "$($row.EmployeeName)".Trim() } else { '' }
        $rawDate = if ($row.LeaveDate)    { "$($row.LeaveDate)".Trim()    } else { '' }

        if (-not $id -and -not $name) { continue }

        $leaveDate = $null
        foreach ($fmt in $acceptedFormats) {
            try { $leaveDate = [datetime]::ParseExact($rawDate, $fmt, $null); break } catch {}
        }
        if (-not $leaveDate) {
            try { $leaveDate = [datetime]::Parse($rawDate) } catch {}
        }
        if (-not $leaveDate) {
            Write-Warning "Leaver row skipped (unparseable LeaveDate '$rawDate'): name='$name' id='$id'"
            continue
        }

        $nameTokens = @(ConvertTo-LeaverTokens $name | Sort-Object -Unique)
        $record = [pscustomobject]@{
            EmployeeID   = $id
            EmployeeName = $name
            LeaveDate    = $leaveDate.Date
            NameTokens   = $nameTokens
        }

        if ($id) {
            $key = $id.ToLowerInvariant()
            if ($byId.ContainsKey($key)) {
                Write-Warning "Duplicate EmployeeID '$id' in leavers CSV - last entry wins."
            }
            $byId[$key] = $record
        }

        # Single-token names (e.g. just "Smith") are too risky for path
        # matching - they collide with file/folder words. Require >=2
        # distinct tokens before registering for name lookup.
        if ($nameTokens.Count -ge 2) {
            $nameKey = ($nameTokens -join '')
            if ($byName.ContainsKey($nameKey) -and $byName[$nameKey] -ne $record) {
                # Two leavers with the same normalised name - mark
                # ambiguous and refuse to match by name for that key.
                [void]$ambiguous.Add($nameKey)
            } else {
                $byName[$nameKey] = $record
            }
        }
    }

    foreach ($k in @($ambiguous)) { $byName.Remove($k) | Out-Null }

    return @{
        ById   = $byId
        ByName = $byName
    }
}

# Match a file to a leaver: ID first (high precision), then name tokens
# (medium precision). Returns $null when no confident match exists. Looks
# at the filename stem and every parent folder name up to the root, so
# layouts like \Leavers\General-6yr\EMP12345 - John Smith\file.docx work.
function Get-LeaverForFile {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]    $FilePath,
        [Parameter(Mandatory)][hashtable] $LeaversIndex
    )

    # Collect candidate strings: filename stem + each parent folder name.
    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add([System.IO.Path]::GetFileNameWithoutExtension($FilePath))
    $dir = Split-Path -Path $FilePath -Parent
    while ($dir) {
        $candidates.Add((Split-Path -Path $dir -Leaf))
        $parent = Split-Path -Path $dir -Parent
        if ([string]::IsNullOrEmpty($parent) -or $parent -eq $dir) { break }
        $dir = $parent
    }

    # 1. ID match. Tokenise each candidate and look for the ID as a whole
    # token so "EMP12345" doesn't accidentally match "EMP123450".
    if ($LeaversIndex.ById.Count -gt 0) {
        foreach ($candidate in $candidates) {
            foreach ($tok in (ConvertTo-LeaverTokens $candidate)) {
                if ($LeaversIndex.ById.ContainsKey($tok)) {
                    return $LeaversIndex.ById[$tok]
                }
            }
        }
    }

    # 2. Name match. A leaver matches if every token of their name is
    # present (in any order) in the candidate's token set.
    if ($LeaversIndex.ByName.Count -gt 0) {
        foreach ($candidate in $candidates) {
            $candidateTokens = @(ConvertTo-LeaverTokens $candidate)
            if ($candidateTokens.Count -eq 0) { continue }
            $candidateSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$candidateTokens, [System.StringComparer]::OrdinalIgnoreCase)
            foreach ($key in $LeaversIndex.ByName.Keys) {
                $record = $LeaversIndex.ByName[$key]
                $allPresent = $true
                foreach ($t in $record.NameTokens) {
                    if (-not $candidateSet.Contains($t)) { $allPresent = $false; break }
                }
                if ($allPresent) { return $record }
            }
        }
    }

    return $null
}


function Execute_Tagging() {
	clear
	$newRun = $false
 
 	$filePathBase = "$Env:LOCALAPPDATA\Temp"
    if (!(Test-Path $filePathBase -PathType Container)) {
        New-Item -Path $filePathBase -ItemType Container -Force 
    }

	# Labelling root - the top-level HR share the scanner walks.
	$FolderName = "\\UK-GH-PURVIEW01\HRDataOld\HR Data"
	Write-Host "Folder Name: $FolderName"

	if (-not (Test-Path -LiteralPath $FolderName -PathType Container)) {
		Write-Error "Labelling root '$FolderName' does not exist or is not reachable from this host."
		return
	}

	# Optional HR leavers list. Drop a Leavers.csv (EmployeeName, EmployeeID,
	# LeaveDate) at the path below to enable per-leaver retention dates.
	# Pipeline still works without it - Update-FileAgeProperties just falls
	# back to today's date as RetentionStartDate when no match is found.
	$script:LeaversIndex = $null
	$leaversCsv = "$filePathBase\Unstructured\Leavers.csv"
	if (Test-Path -LiteralPath $leaversCsv -PathType Leaf) {
		$script:LeaversIndex = Import-LeaversList -CsvPath $leaversCsv
		if ($script:LeaversIndex) {
			Write-Host ("Leavers list loaded from {0}: {1} by ID, {2} by name" -f $leaversCsv, $script:LeaversIndex.ById.Count, $script:LeaversIndex.ByName.Count)
		}
	}
	else {
		Write-Host "No Leavers.csv found at $leaversCsv - leaver matching disabled (today will be used as retention start)."
	}
 
	# Define the file collection location
	$targetDir = "$filePathBase\Unstructured"
 
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
		$filesToScan = Get-ApplicableFiles -FolderName $FolderName -Buckets $script:RetentionBuckets
		if($filesToScan.Count -ne 0) {
			$filesToScanUnique = $filesToScan | sort -Unique
			foreach ($file in $filesToScanUnique) {
				Add-Content -Path $TargetFiles -Value $file
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


function Kill-Process {
    param(
        [int]$MaxRuntimeSeconds = 60,
        [int]$CheckIntervalSeconds = 60,
        [int]$HardStopAfterMinutes = 0,
        [string]$LogPath = "C:\Users\KB9\AppData\Local\Temp\KillProcess.log"
    )

    # Isolate job preferences so it can NEVER terminate the parent
    $ErrorActionPreference  = 'SilentlyContinue'
    $WarningPreference      = 'Continue'
    $ProgressPreference     = 'SilentlyContinue'

    $officeProcesses = @("WINWORD", "EXCEL", "POWERPNT")
    $startTime = Get-Date

    "$(Get-Date -Format o) | Kill-Process started" | Out-File -Append $LogPath

    while ($true) {
        try {
            $now = Get-Date

            if ($HardStopAfterMinutes -gt 0 -and
                ($now - $startTime).TotalMinutes -gt $HardStopAfterMinutes) {
                "$(Get-Date -Format o) | Hard stop timeout reached" | Out-File -Append $LogPath
                break
            }

            foreach ($procName in $officeProcesses) {
                Get-Process -Name $procName | ForEach-Object {
                    $runtime = $now - $_.StartTime
                    if ($runtime.TotalSeconds -gt $MaxRuntimeSeconds) {
                        "$(Get-Date -Format o) | Stopping $($_.ProcessName) PID=$($_.Id) Runtime=$([math]::Round($runtime.TotalMinutes,2)) min" |
                            Out-File -Append $LogPath
                        Stop-Process -Id $_.Id -Force
                    }
                }
            }
        }
        catch {
            # Swallow any unexpected errors from the monitor
            "$(Get-Date -Format o) | Kill-Process internal error: $($_.Exception.Message)" | Out-File -Append $LogPath
        }

        Start-Sleep -Seconds $CheckIntervalSeconds
    }

    "$(Get-Date -Format o) | Kill-Process stopped" | Out-File -Append $LogPath
}


function Start-KillProcessMonitor {
    param(
        [int]$MaxRuntimeSeconds = 60,
        [int]$CheckIntervalSeconds = 60,
        [string]$LogPath = "C:\Users\KB9\AppData\Local\Temp\KillProcess.log",
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
            Read-Host "Press Enter to continue the monitor loop..."
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

    # Kill-monitor log lives next to the regular run logs, so we are not
    # dependent on C:\temp existing on the VDI. Created if absent.
    $killLogDir = Join-Path $Env:LOCALAPPDATA 'Temp\Unstructured'
    if (-not (Test-Path -LiteralPath $killLogDir -PathType Container)) {
        New-Item -Path $killLogDir -ItemType Directory -Force | Out-Null
    }
    $killLogPath = Join-Path $killLogDir 'KillProcess.log'

    # Start the external monitor process (hidden)
    $monitorProc = Start-KillProcessMonitor -MaxRuntimeSeconds 60 -CheckIntervalSeconds 60 -LogPath $killLogPath -ShowWindow -NoExit

    $taggingFailed = $false

    try {
        Execute_Tagging
    }
    catch {
        $taggingFailed = $true
        Write-Warning ("Execute_Tagging failed (continuing): {0}" -f $_.Exception.Message)
        # Do NOT rethrow if you want the script to continue
    }
    finally {
        if ($monitorProc) {
            Stop-KillProcessMonitor -MonitorProcess $monitorProc
        }
    }

    # Normalize exit code if you're in a pipeline that treats non-zero as failure
    $global:LASTEXITCODE = 0
}

Invoke-ExecuteTaggingSafely

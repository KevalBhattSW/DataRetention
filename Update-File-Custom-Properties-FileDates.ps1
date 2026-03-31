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

    $Item.LastWriteTime = $LastWriteTime
    Start-Sleep -Milliseconds $MetadataDuration
    $Item.LastAccessTime = $LastAccessTime
    if ($FileReadOnly) {
        $Item.IsReadOnly = $true
    }


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

#Function to loop through a collection of files, check their age and create/update custom document properties
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
    $ScriptPath = Join-Path $PSScriptRoot '\update_pdf_properties.py'

    if (Test-Path $ScriptPath) {
        Write-Host "Found file at $ScriptPath"
    } else {
        Write-Host "Python is not stored in expected location. PDF tagging cannot be completed"
        $pdfRun = $false
    }

    if($pdfRun){
        try{
            # Create venv next to your PowerShell + Python scripts
            py -$majorMinor -m venv $venvDir

            try {
                $out = py -$majorMinor -c 'import sys; print(sys.prefix!=sys.base_prefix)' 2>$null
                if ($out -match 'True') { $activeBySys = $true }
            } catch { }

            if(!$activeBySys) {
                # Activate it in PowerShell:
                $venvActivate

                # Install dependencies inside the venv:
                py -$majorMinor -m pip install --upgrade pip
                py -$majorMinor -m pip install pypdf
                }
            }
        catch {
            Write-Error "Failed to install Python virtual environment. PDF tagging cannot be completed"
            $pdfRun = $false
        }
    }

    
    # Adjust name/path as needed (e.g., ".venv" or "venv")
    # For PowerShell 7 on Linux/macOS, use 'bin/python' instead

    if (Test-Path $venvPython) {
        Write-Host "Venv exists at: $venvDir"
    } else {
        Write-Host "Venv not found (expected $venvPython)"
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

			if ($doc -ne $null) {
				$doc.Close()
			}
			
			if ($app -ne $null) {
				$app.Quit()
			}

			if ($app -ne $null) {
				[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null
			}
			$doc = $null
			$app = $null

			# Update file timestamps
			if($item -eq $null) {
    			$item = Get-Item $objFile
            }
			$item.LastWriteTime = $dtLastModified
			Start-Sleep -Milliseconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
			$item.LastAccessTime = $dtLastAccessedDoc
			if ($fileReadOnly -eq $true) {
				$item.IsReadOnly = $true
			}
		}

		# If the file has been opened
		if ($doc -ne $null -or $pdfApp -eq $true){
			# Test if the file was created over three years ago and accessed over 18 months ago - already tested in file collation
			$blPropertyLastAccessed = $true
			$blPropertyCreated = $true
			
			#Convert boolean values to text strings for Purview to read correctly
			$strPropertyLastAccessed = if($blPropertyLastAccessed) {"True"} else {"False"}
			$strPropertyCreated = if($blPropertyCreated) {"True"} else {"False"}

			if($officeApp -eq $True) {
				$propertyExistsOriginalPath = Set-OfficeDocCustomProperty "OriginalPath" $objFile $doc
				$propertyExistsLastAccessed = Set-OfficeDocCustomProperty "LastAccessedThreshold" $strPropertyLastAccessed $doc
				$propertyExistsCreated = Set-OfficeDocCustomProperty "CreatedThreshold" $strPropertyCreated $doc
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
			elseif($pdfApp -eq $True -and $pdfRun -eq $True) {
				$PropertyName = "OriginalPath"
				$PropertyValue = $objFile
				switch (& $PythonPath $ScriptPath $PropertyName $PropertyValue $objFile) {
					1 {$isPasswordProtected = $true}
					2 {$isPasswordProtected = $true}
					-1 {$isError = $true}
					default {
						$isPasswordProtected = $false
						$isPasswordProtected = $false
						$isError = $false
					}
				}
				
				$PropertyName = "LastAccessedThreshold"
				$PropertyValue = $strPropertyLastAccessed
				switch (& $PythonPath $ScriptPath $PropertyName $PropertyValue $objFile) {
					1 {$isPasswordProtected = $true}
					2 {$isPasswordProtected = $true}
					-1 {$isError = $true}
					default {
						$isPasswordProtected = $false
						$isPasswordProtected = $false
						$isError = $false
					}
				}
				$PropertyName = "CreatedThreshold"
				$PropertyValue = $strPropertyCreated
				switch (& $PythonPath $ScriptPath $PropertyName $PropertyValue $objFile) {
					1 {$isPasswordProtected = $true}
					2 {$isPasswordProtected = $true}
					-1 {$isError = $true}
					default {
						$isPasswordProtected = $false
						$isPasswordProtected = $false
						$isError = $false
					}
				}
			}

			if ($isPasswordProtected -eq $true) {
				#Write to log that file has been updated
				$item.LastWriteTime = $dtLastModified
				Start-Sleep -Milliseconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
				$item.LastAccessTime = $dtLastAccessedDoc
				if ($fileReadOnly -eq $true) {
					$item.IsReadOnly = $true
				}
		
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
				$item.LastWriteTime = $dtLastModified
				Start-Sleep -Milliseconds $metadataDuration # If we don't pause here, the dates do not get updated correctly
				$item.LastAccessTime = $dtLastAccessedDoc
				if($fileReadOnly -eq $true) {
					$item.IsReadOnly = $true
				}
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



function Get-ApplicableFiles {
    [CmdletBinding()]
    [OutputType([System.Collections.ArrayList])]
    param (
        [Parameter(Mandatory)]
        [string] $FolderName,

        [Parameter(Mandatory)]
        [int] $LastAccessedMonths,

        [Parameter(Mandatory)]
        [int] $CreatedMonths
    )

    # Validate folder
    if (-not (Test-Path -LiteralPath $FolderName -PathType Container)) {
        Write-Error "Folder '$FolderName' does not exist."
        return [System.Collections.ArrayList]::new()
    }

    # Normalize months
    $lastAccessedMonthsAbs = [math]::Abs($LastAccessedMonths)
    $createdMonthsAbs     = [math]::Abs($CreatedMonths)

    # Threshold dates
    $lastAccessThreshold = (Get-Date).AddMonths(-$lastAccessedMonthsAbs)
    $creationThreshold   = (Get-Date).AddMonths(-$createdMonthsAbs)

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
            $_.LastAccessTime -lt $lastAccessThreshold -and
            $_.CreationTime   -lt $creationThreshold -and
            ($officeExtensions -contains $_.Extension.ToLowerInvariant())
        } |
        ForEach-Object {
            [void]$result.Add($_.FullName)
        }

        # Recurse into subfolders
        Get-ChildItem -LiteralPath $FolderName -Directory -ErrorAction Stop |
        ForEach-Object {
            Write-Verbose "Recursing into subfolder: $($_.FullName)"

            $childResults = Get-ApplicableFiles `
                -FolderName $_.FullName `
                -LastAccessedMonths $lastAccessedMonthsAbs `
                -CreatedMonths $createdMonthsAbs

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
 
 	$filePathBase = "$Env:LOCALAPPDATA\Temp"
    if (!(Test-Path $filePathBase -PathType Container)) {
        New-Item -Path $filePathBase -ItemType Container -Force 
    }

	# Define the folder path
	$FolderName = "$filePathBase\Unstructured\labelling"
	#$FolderName = (Get-Content -Path "$filePathBase\searchpath.txt").Trim()
	Write-Host "Folder Nameis: $FolderName"
 
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
		$filesToScan = Get-ApplicableFiles -FolderName $FolderName -LastAccessedMonths 0 -CreatedMonths 1
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

    # Start the external monitor process (hidden)
    $monitorProc = Start-KillProcessMonitor -MaxRuntimeSeconds 60 -CheckIntervalSeconds 60 -LogPath "C:\temp\KillProcess.log" -ShowWindow -NoExit

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
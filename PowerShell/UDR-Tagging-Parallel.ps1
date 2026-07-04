param(
    [Parameter(Mandatory=$true)]
    [string]$DrivePath,
    [Parameter(Mandatory=$true)]
    [string]$ScriptPath
)



if(-not(Test-Path -LiteralPath $DrivePath -PathType Container)) {
    Write-Error "DrivePath '$drivePath' does not exist or is not accessible. Aborting."
    Exit 1
}

function Test-LegacyOfficeProtectionDiagnostic {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $log = "$($env:LOCALAPPDATA)\Temp\LegacyProtectionDebug.txt"

    $result = [PSCustomObject]@{
        Path               = $Path
        IsPasswordToOpen   = $false
        IsPasswordToModify = $false
        IsProtected        = $false
        Reason             = "NoProtection"
    }

    Add-Content $log "=== Testing: $Path ==="

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Reason = "FileNotFound"
        Add-Content $log "RESULT: FileNotFound"
        return $result
    }

    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $br = New-Object System.IO.BinaryReader($fs)

        $header = $br.ReadBytes(8)
        $oleSig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
        if (($header -join ',') -ne ($oleSig -join ',')) {
            $result.Reason = "NotOLE"
            Add-Content $log "RESULT: NotOLE - header bytes: $($header -join ',')"
            $fs.Close()
            return $result
        }
        Add-Content $log "OLE signature confirmed"

        $fs.Position = 0x1E
        $sectorShift = $br.ReadInt16()
        $sectorSize  = [int][math]::Pow(2, $sectorShift)
        Add-Content $log "SectorShift=$sectorShift  SectorSize=$sectorSize"

        $fs.Position = 0x38
        $miniSectorCutoff = $br.ReadInt32()
        Add-Content $log "MiniSectorCutoff=$miniSectorCutoff"

        $fs.Position = 0x30
        $dirStartSector = $br.ReadInt32()
        Add-Content $log "DirStartSector=$dirStartSector"

        $fs.Position = 0x2C
        $fatSectorCount = $br.ReadInt32()
        Add-Content $log "FatSectorCount=$fatSectorCount"

        $fatSectorList = @()
        $fs.Position = 0x4C
        for ($i = 0; $i -lt [math]::Min(109, $fatSectorCount); $i++) {
            $sec = $br.ReadInt32()
            if ($sec -ge 0) { $fatSectorList += $sec }
        }
        Add-Content $log "FAT sectors found: $($fatSectorList.Count) -> $($fatSectorList -join ',')"

        $fatData = New-Object System.Collections.Generic.List[int]
        foreach ($fatSec in $fatSectorList) {
            $fs.Position = ($fatSec + 1) * $sectorSize
            $entriesPerSector = $sectorSize / 4
            for ($i = 0; $i -lt $entriesPerSector; $i++) {
                $fatData.Add($br.ReadInt32())
            }
        }
        Add-Content $log "FAT entries loaded: $($fatData.Count)"

        $streamNames = @{}
        $sector  = $dirStartSector
        $visited = @{}

        while ($sector -ge 0 -and $sector -ne -2 -and $sector -ne -1 -and -not $visited.ContainsKey($sector)) {
            $visited[$sector] = $true
            $fs.Position = ($sector + 1) * $sectorSize
            $dirBytes    = $br.ReadBytes($sectorSize)
            $entryCount  = $sectorSize / 128

            for ($e = 0; $e -lt $entryCount; $e++) {
                $offset    = $e * 128
                $nameLen   = [System.BitConverter]::ToInt16($dirBytes, $offset + 64)
                $entryType = $dirBytes[$offset + 66]

                if ($entryType -gt 0 -and $nameLen -gt 0) {
                    $nameBytes  = $dirBytes[$offset..($offset + $nameLen - 1)]
                    $name       = [System.Text.Encoding]::Unicode.GetString($nameBytes).TrimEnd([char]0)
                    $streamSize = [int64][System.BitConverter]::ToUInt32($dirBytes, $offset + 120)
                    $startSec   = [System.BitConverter]::ToInt32($dirBytes, $offset + 116)
                    $streamNames[$name] = [PSCustomObject]@{ Size = $streamSize; Start = $startSec; Type = $entryType }
                    Add-Content $log "  Stream: '$name'  Type=$entryType  Start=$startSec  Size=$streamSize"
                }
            }

            if ($sector -lt $fatData.Count) {
                $sector = $fatData[$sector]
            } else {
                break
            }
        }

        $fs.Close()

        Add-Content $log "All stream names: $($streamNames.Keys -join ', ')"

        if ($streamNames.ContainsKey("EncryptedPackage") -or $streamNames.ContainsKey("EncryptionInfo")) {
            $result.IsPasswordToOpen = $true
            $result.IsProtected      = $true
            $result.Reason           = "EncryptedToOpen"
            Add-Content $log "RESULT: EncryptedToOpen"
            return $result
        }

        $extension = [System.IO.Path]::GetExtension($Path).ToLower()
        Add-Content $log "Extension: $extension  Checking write-reservation..."

        switch ($extension) {

            ".doc" {
                if ($streamNames.ContainsKey("WordDocument")) {
                    $si  = $streamNames["WordDocument"]
                    Add-Content $log "WordDocument stream: Start=$($si.Start) Size=$($si.Size)"
                    $raw = Read-OleStream -Path $Path -StartSector $si.Start -Size $si.Size -SectorSize $sectorSize -Fat $fatData -MiniCutoff $miniSectorCutoff
                    Add-Content $log "WordDocument raw bytes returned: $(if ($null -eq $raw) { 'NULL' } else { $raw.Length })"
                    if ($raw -ne $null -and $raw.Length -ge 12) {
                        $fibFlags = [System.BitConverter]::ToUInt16($raw, 10)
                        Add-Content $log "FIB flags word (offset 10): 0x$('{0:X4}' -f $fibFlags)  fWriteReservation bit (0x0080): $(($fibFlags -band 0x0080) -ne 0)"
                        if ($fibFlags -band 0x0080) {
                            $result.IsPasswordToModify = $true
                            $result.IsProtected        = $true
                            $result.Reason             = "WriteReservedDoc"
                        }
                    } else {
                        Add-Content $log "WARNING: WordDocument stream too short or null"
                    }
                } else {
                    Add-Content $log "WARNING: WordDocument stream not found in directory"
                }
            }

            ".xls" {
                $streamName = if ($streamNames.ContainsKey("Workbook")) { "Workbook" } else { "Book" }
                Add-Content $log "Using XLS stream name: '$streamName'  Found=$($streamNames.ContainsKey($streamName))"
                if ($streamNames.ContainsKey($streamName)) {
                    $si  = $streamNames[$streamName]
                    Add-Content $log "Stream: Start=$($si.Start) Size=$($si.Size)"
                    $raw = Read-OleStream -Path $Path -StartSector $si.Start -Size $si.Size -SectorSize $sectorSize -Fat $fatData -MiniCutoff $miniSectorCutoff
                    Add-Content $log "Raw bytes returned: $(if ($null -eq $raw) { 'NULL' } else { $raw.Length })"
                    if ($raw -ne $null) {
                        $i = 0
                        while ($i -lt ($raw.Length - 4)) {
                            $recType = [System.BitConverter]::ToUInt16($raw, $i)
                            $recLen  = [System.BitConverter]::ToUInt16($raw, $i + 2)
                            if ($recType -eq 0x002F) {
                                Add-Content $log "Found FilePass record (0x002F) at offset $i — PasswordToOpen"
                                $result.IsPasswordToOpen = $true; $result.IsProtected = $true; $result.Reason = "EncryptedToOpen"
                                break
                            }
                            if ($recType -eq 0x005C -and $recLen -ge 4) {
                                $pwHash = [System.BitConverter]::ToUInt16($raw, $i + 4 + 2)
                                Add-Content $log "Found FileSharing record (0x005C) at offset $i — pwHash=0x$('{0:X4}' -f $pwHash)"
                                if ($pwHash -ne 0) {
                                    $result.IsPasswordToModify = $true; $result.IsProtected = $true; $result.Reason = "WriteReservedXls"
                                }
                            }
                            $i += 4 + $recLen
                            if ($recLen -eq 0 -and $recType -eq 0) { break }
                        }
                    }
                }
            }

            ".ppt" {
                $validFirstRecTypes = @(0x03E8, 0x0FF0, 0x0FF3, 0x0FF4, 0x07E5, 0x0FA0)

                # Check HeaderToken in Current User mini-stream
                $rootEntry = $streamNames["Root Entry"]
                if ($rootEntry -and $rootEntry.Start -ge 0 -and $rootEntry.Start -ne -2) {
                    $miniBytes = Read-OleStream -Path $Path -StartSector $rootEntry.Start `
                                                -Size $rootEntry.Size -SectorSize $sectorSize `
                                                -Fat $fatData -MiniCutoff $miniSectorCutoff
                    if ($miniBytes -ne $null -and $miniBytes.Length -ge 16) {
                        $headerToken = [System.BitConverter]::ToUInt32($miniBytes, 12)
                        Add-Content $log "Current User HeaderToken: 0x$('{0:X8}' -f $headerToken)"
                        if ($headerToken -eq 0xF3D1C4DF) {
                            $result.IsPasswordToOpen = $true
                            $result.IsProtected      = $true
                            $result.Reason           = "EncryptedToOpen_PPT_CurrentUser"
                            Add-Content $log "RESULT: Encrypted via HeaderToken"
                        }
                    }
                }

                # Check PowerPoint Document stream first record type
                if ($streamNames.ContainsKey("PowerPoint Document")) {
                    $si  = $streamNames["PowerPoint Document"]
                    Add-Content $log "PPT stream: Start=$($si.Start) Size=$($si.Size)"
                    $rawFirst = Read-OleStream -Path $Path -StartSector $si.Start -Size ([math]::Min($si.Size, 8)) `
                                               -SectorSize $sectorSize -Fat $fatData -MiniCutoff $miniSectorCutoff
                    Add-Content $log "PPT first 8 bytes: $(if ($null -eq $rawFirst) { 'NULL' } else { ($rawFirst | ForEach-Object { '{0:X2}' -f $_ }) -join ' ' })"
                    if ($rawFirst -ne $null -and $rawFirst.Length -ge 8) {
                        $firstRecType = [System.BitConverter]::ToUInt16($rawFirst, 2)
                        Add-Content $log "PPT first recType: 0x$('{0:X4}' -f $firstRecType)"
                        if ($validFirstRecTypes -notcontains $firstRecType) {
                            $result.IsPasswordToOpen = $true
                            $result.IsProtected      = $true
                            $result.Reason           = "EncryptedToOpen_PPT_StreamGarbled"
                            Add-Content $log "RESULT: Encrypted - first record type 0x$('{0:X4}' -f $firstRecType) is not a valid PPT record"
                        }
                    }
                }

                # Password-to-modify: scan for WriteAccessAtom (0x03EF) only if not encrypted
                if (-not $result.IsPasswordToOpen -and $streamNames.ContainsKey("PowerPoint Document")) {
                    $si  = $streamNames["PowerPoint Document"]
                    $rawFull = Read-OleStream -Path $Path -StartSector $si.Start -Size $si.Size `
                                              -SectorSize $sectorSize -Fat $fatData -MiniCutoff $miniSectorCutoff
                    Add-Content $log "PPT full stream bytes: $(if ($null -eq $rawFull) { 'NULL' } else { $rawFull.Length })"
                    if ($rawFull -ne $null) {
                        $i             = 0
                        $maxIterations = 10000
                        $iteration     = 0
                        while ($i -lt ($rawFull.Length - 8) -and $iteration -lt $maxIterations) {
                            $iteration++
                            $recVer  = $rawFull[$i] -band 0x0F
                            $recType = [System.BitConverter]::ToUInt16($rawFull, $i + 2)
                            $recLen  = [System.BitConverter]::ToUInt32($rawFull, $i + 4)

                            Add-Content $log "  Iter=$iteration  Offset=$i  recVer=$recVer  recType=0x$('{0:X4}' -f $recType)  recLen=$recLen"

                            if ($recType -eq 0x03EF) {
                                $flagByte = if (($i + 8) -lt $rawFull.Length) { $rawFull[$i + 8] } else { 0 }
                                Add-Content $log "Found WriteAccessAtom at offset $i - flagByte=$flagByte"
                                if ($flagByte -eq 0x01) {
                                    $result.IsPasswordToModify = $true
                                    $result.IsProtected        = $true
                                    $result.Reason             = "WriteReservedPpt"
                                }
                                break
                            }

                            if ($recVer -eq 0x0F) {
                                $i += 8
                            } else {
                                if ($recLen -gt ($rawFull.Length - $i - 8)) { 
                                    Add-Content $log "  Overrun guard triggered at offset $i"
                                    break 
                                }
                                $i += 8 + [int]$recLen
                            }
                        }
                        Add-Content $log "PPT write-reservation scan complete: iterations=$iteration  finalOffset=$i"
                    }
                }
            }
        }
    }
    catch {
        $result.Reason = "ReadError: $($_.Exception.Message)"
        Add-Content $log "EXCEPTION: $($_.Exception.Message)"
        Add-Content $log $_.ScriptStackTrace
    }

    Add-Content $log "FINAL RESULT: IsPasswordToOpen=$($result.IsPasswordToOpen)  IsPasswordToModify=$($result.IsPasswordToModify)  Reason=$($result.Reason)"
    Add-Content $log ""
    return $result
}

function Test-LegacyOfficeProtection {
    <#
    .SYNOPSIS
        Detects password-to-open AND password-to-modify in legacy OLE binary Office files
        (.doc, .xls, .ppt) without opening the file in any Office application.
    .OUTPUTS
        [PSCustomObject] with:
          .IsPasswordToOpen   [bool]
          .IsPasswordToModify [bool]
          .IsProtected        [bool]  (either of the above)
          .Reason             [string]
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $result = [PSCustomObject]@{
        Path              = $Path
        IsPasswordToOpen   = $false
        IsPasswordToModify = $false
        IsProtected        = $false
        Reason             = "NoProtection"
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Reason = "FileNotFound"
        return $result
    }

    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $br = New-Object System.IO.BinaryReader($fs)

        # Verify OLE signature
        $header = $br.ReadBytes(8)
        $oleSig = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
        if (($header -join ',') -ne ($oleSig -join ',')) {
            $result.Reason = "NotOLE"
            $fs.Close()
            return $result
        }

        # Sector size (power-of-two exponent at offset 0x1E)
        $fs.Position = 0x1E
        $sectorShift = $br.ReadInt16()
        $sectorSize  = [int][math]::Pow(2, $sectorShift)

        # Mini-stream cutoff at offset 0x38
        $fs.Position = 0x38
        $miniSectorCutoff = $br.ReadInt32()

        # Directory sector start at offset 0x30
        $fs.Position = 0x30
        $dirStartSector = $br.ReadInt32()

        # FAT sector count and first FAT sector location
        $fs.Position = 0x2C
        $fatSectorCount = $br.ReadInt32()
        $fs.Position = 0x4C
        $firstFatSector = $br.ReadInt32()

        # Build FAT chain (read all FAT sectors)
        $fat = New-Object int[] ($fatSectorCount * ($sectorSize / 4))
        $fatIndex = 0
        $fs.Position = 0x4C
        # Read DIFAT from header (first 109 FAT sector locations at offset 0x4C)
        $fatSectorList = @()
        $fs.Position = 0x4C
        for ($i = 0; $i -lt [math]::Min(109, $fatSectorCount); $i++) {
            $sec = $br.ReadInt32()
            if ($sec -ge 0) { $fatSectorList += $sec }
        }

        $fatData = New-Object System.Collections.Generic.List[int]
        foreach ($fatSec in $fatSectorList) {
            $fs.Position = ($fatSec + 1) * $sectorSize
            $entriesPerSector = $sectorSize / 4
            for ($i = 0; $i -lt $entriesPerSector; $i++) {
                $fatData.Add($br.ReadInt32())
            }
        }

        # Walk directory sector chain and collect stream names
        $streamNames = @{}
        $sector = $dirStartSector
        $visited = @{}

        while ($sector -ge 0 -and $sector -ne -2 -and $sector -ne -1 -and -not $visited.ContainsKey($sector)) {
            $visited[$sector] = $true
            $fs.Position = ($sector + 1) * $sectorSize

            $dirBytes = $br.ReadBytes($sectorSize)
            $entryCount = $sectorSize / 128

            for ($e = 0; $e -lt $entryCount; $e++) {
                $offset     = $e * 128
                $nameLen    = [System.BitConverter]::ToInt16($dirBytes, $offset + 64)  # byte length of name
                $entryType  = $dirBytes[$offset + 66]   # 0=empty, 1=storage, 2=stream, 5=root

                if ($entryType -gt 0 -and $nameLen -gt 0) {
                    $nameBytes = $dirBytes[$offset..($offset + $nameLen - 1)]
                    $name      = [System.Text.Encoding]::Unicode.GetString($nameBytes).TrimEnd([char]0)
                    $streamSize = [int64][System.BitConverter]::ToUInt32($dirBytes, $offset + 120)
                    $startSec   = [System.BitConverter]::ToInt32($dirBytes, $offset + 116)
                    $streamNames[$name] = [PSCustomObject]@{ Size = $streamSize; Start = $startSec; Type = $entryType }
                }
            }

            # Follow FAT chain to next directory sector
            if ($sector -lt $fatData.Count) {
                $sector = $fatData[$sector]
            } else {
                break
            }
        }

        $fs.Close()

        # ---------------------------------------------------------------
        # PASSWORD-TO-OPEN detection
        # ---------------------------------------------------------------
        # If EncryptedPackage + EncryptionInfo exist, the whole file is encrypted
        if ($streamNames.ContainsKey("EncryptedPackage") -or $streamNames.ContainsKey("EncryptionInfo")) {
            $result.IsPasswordToOpen  = $true
            $result.IsProtected       = $true
            $result.Reason            = "EncryptedToOpen"
            return $result
        }

        # ---------------------------------------------------------------
        # PASSWORD-TO-MODIFY detection (WriteReservation / WriteProtection)
        # ---------------------------------------------------------------
        # Word .doc: the "1Table" or "0Table" stream contains a FIB with wri flag,
        # but the simplest reliable signal is the WorkBook / WordDocument stream
        # containing a write-reservation password hash.
        #
        # The most portable approach: look for the FilePass record in the
        # Document Summary stream, OR check known stream names for write protection.
        #
        # For .xls: Workbook stream starts with BOF record; FilePass record (0x002F)
        # near the start means open-password. WRITEACCESS record can contain
        # write-reservation. Simpler: check for "WriteAccess" password via
        # the FILEPASS record type 0x05 or via the FileSharing info in the
        # WorkBook stream header area.
        #
        # Practical approach that covers all three apps without full parser:
        # Read the known primary stream and scan for the write-reservation signature bytes.

        $extension = [System.IO.Path]::GetExtension($Path).ToLower()

        switch ($extension) {

            ".doc" {
                # WordDocument stream: write reservation is flagged by
                # FIB.fWriteReservation bit. FIB starts at offset 0 of WordDocument stream.
                # Flags word is at FIB offset 0x0A (word). Bit 0x0080 = fWriteReservation.
                $streamName = "WordDocument"
                if ($streamNames.ContainsKey($streamName)) {
                    $streamInfo = $streamNames[$streamName]
                    $rawBytes   = Read-OleStream -Path $Path -StartSector $streamInfo.Start `
                                                 -Size $streamInfo.Size -SectorSize $sectorSize `
                                                 -Fat $fatData -MiniCutoff $miniSectorCutoff
                    if ($rawBytes -ne $null -and $rawBytes.Length -ge 12) {
                        $fibFlags = [System.BitConverter]::ToUInt16($rawBytes, 10)
                        if ($fibFlags -band 0x0080) {
                            $result.IsPasswordToModify = $true
                            $result.IsProtected        = $true
                            $result.Reason             = "WriteReservedDoc"
                        }
                    }
                    else {
                        Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Path WARNING: WordDocument stream returned $(if ($null -eq $rawBytes) { 'null' } else { $rawBytes.Length }) bytes - protection check skipped"
                    }
                }
            }

            ".xls" {
                # Workbook stream: scan for FilePass (0x002F) record — means open password.
                # Write reservation: look for FILESHARING record (0x005C).
                # FILESHARING record: fReadOnlyRecommended (2 bytes) + password hash (2 bytes).
                # If password hash != 0x0000 at offset +2, write-reserved.
                $streamName = "Workbook"
                if (-not $streamNames.ContainsKey($streamName)) { $streamName = "Book" }
                if ($streamNames.ContainsKey($streamName)) {
                    $streamInfo = $streamNames[$streamName]
                    $rawBytes   = Read-OleStream -Path $Path -StartSector $streamInfo.Start `
                                                 -Size $streamInfo.Size -SectorSize $sectorSize `
                                                 -Fat $fatData -MiniCutoff $miniSectorCutoff
                    if ($rawBytes -ne $null) {
                        $i = 0
                        while ($i -lt ($rawBytes.Length - 4)) {
                            $recType = [System.BitConverter]::ToUInt16($rawBytes, $i)
                            $recLen  = [System.BitConverter]::ToUInt16($rawBytes, $i + 2)

                            if ($recType -eq 0x002F) {
                                # FilePass = open password
                                $result.IsPasswordToOpen = $true
                                $result.IsProtected      = $true
                                $result.Reason           = "EncryptedToOpen"
                                break
                            }
                            if ($recType -eq 0x005C -and $recLen -ge 4) {
                                # FileSharing record
                                $pwHash = [System.BitConverter]::ToUInt16($rawBytes, $i + 4 + 2)
                                if ($pwHash -ne 0) {
                                    $result.IsPasswordToModify = $true
                                    $result.IsProtected        = $true
                                    $result.Reason             = "WriteReservedXls"
                                }
                            }

                            $i += 4 + $recLen
                            if ($recLen -eq 0) { break }  # safety: avoid infinite loop
                        }
                    }
                }
            }

            ".ppt" {
                $validFirstRecTypes = @(0x03E8, 0x0FF0, 0x0FF3, 0x0FF4, 0x07E5, 0x0FA0)

                # Check HeaderToken in Current User mini-stream
                $rootEntry = $streamNames["Root Entry"]
                if ($rootEntry -and $rootEntry.Start -ge 0 -and $rootEntry.Start -ne -2) {
                    $miniBytes = Read-OleStream -Path $Path -StartSector $rootEntry.Start `
                                                -Size $rootEntry.Size -SectorSize $sectorSize `
                                                -Fat $fatData -MiniCutoff $miniSectorCutoff
                    if ($miniBytes -ne $null -and $miniBytes.Length -ge 16) {
                        $headerToken = [System.BitConverter]::ToUInt32($miniBytes, 12)
                        if ($headerToken -eq 0xF3D1C4DF) {
                            $result.IsPasswordToOpen = $true
                            $result.IsProtected      = $true
                            $result.Reason           = "EncryptedToOpen_PPT_CurrentUser"
                        }
                    }
                }

                # Check PowerPoint Document stream first record type
                if ($streamNames.ContainsKey("PowerPoint Document")) {
                    $si  = $streamNames["PowerPoint Document"]
                    $rawFirst = Read-OleStream -Path $Path -StartSector $si.Start -Size ([math]::Min($si.Size, 8)) `
                                               -SectorSize $sectorSize -Fat $fatData -MiniCutoff $miniSectorCutoff
                    if ($rawFirst -ne $null -and $rawFirst.Length -ge 8) {
                        $firstRecType = [System.BitConverter]::ToUInt16($rawFirst, 2)
                        if ($validFirstRecTypes -notcontains $firstRecType) {
                            $result.IsPasswordToOpen = $true
                            $result.IsProtected      = $true
                            $result.Reason           = "EncryptedToOpen_PPT_StreamGarbled"
                        }
                    }
                }

                # Password-to-modify: scan for WriteAccessAtom (0x03EF) only if not encrypted
                if (-not $result.IsPasswordToOpen -and $streamNames.ContainsKey("PowerPoint Document")) {
                    $si  = $streamNames["PowerPoint Document"]
                    $rawFull = Read-OleStream -Path $Path -StartSector $si.Start -Size $si.Size `
                                              -SectorSize $sectorSize -Fat $fatData -MiniCutoff $miniSectorCutoff
                    if ($rawFull -ne $null) {
                        $i             = 0
                        $maxIterations = 10000
                        $iteration     = 0
                        while ($i -lt ($rawFull.Length - 8) -and $iteration -lt $maxIterations) {
                            $iteration++
                            $recVer  = $rawFull[$i] -band 0x0F
                            $recType = [System.BitConverter]::ToUInt16($rawFull, $i + 2)
                            $recLen  = [System.BitConverter]::ToUInt32($rawFull, $i + 4)


                            if ($recType -eq 0x03EF) {
                                $flagByte = if (($i + 8) -lt $rawFull.Length) { $rawFull[$i + 8] } else { 0 }
                                if ($flagByte -eq 0x01) {
                                    $result.IsPasswordToModify = $true
                                    $result.IsProtected        = $true
                                    $result.Reason             = "WriteReservedPpt"
                                }
                                break
                            }

                            if ($recVer -eq 0x0F) {
                                $i += 8
                            } else {
                                if ($recLen -gt ($rawFull.Length - $i - 8)) { 
                                    Add-Content $log "  Overrun guard triggered at offset $i"
                                    break 
                                }
                                $i += 8 + [int]$recLen
                            }
                        }
                    }
                }
            }
        }
    }
    catch {
        $result.Reason = "ReadError: $($_.Exception.Message)"
    }

    return $result
}


function Read-OleStream {
    <#
    .SYNOPSIS
        Reads raw bytes from an OLE stream given its start sector and size.
        Handles both normal sectors and mini-stream sectors.
    #>
    param(
        [string]$Path,
        [int]$StartSector,
        [int64]$Size,
        [int]$SectorSize,
        [System.Collections.Generic.List[int]]$Fat,
        [int]$MiniCutoff
    )

    # Mini-stream not supported here (streams < MiniCutoff live in root's data).
    # For the primary document streams (WordDocument, Workbook, PowerPoint Document)
    # these are always full-sector streams, so this covers all three cases.
    # Mini-stream support can be added later if needed for edge cases.

    try {
        $fs     = [System.IO.File]::OpenRead($Path)
        $br     = New-Object System.IO.BinaryReader($fs)
        $buffer = New-Object System.Collections.Generic.List[byte]

        $sector    = $StartSector
        $remaining = $Size
        $visited   = @{}

        while ($sector -ge 0 -and $sector -ne -2 -and $sector -ne -1 -and -not $visited.ContainsKey($sector)) {
            $visited[$sector] = $true
            $fs.Position      = ($sector + 1) * $SectorSize

            $toRead = [math]::Min($SectorSize, $remaining)
            $bytes  = $br.ReadBytes($toRead)
            $buffer.AddRange($bytes)
            $remaining -= $toRead

            if ($sector -lt $Fat.Count) {
                $sector = $Fat[$sector]
            } else {
                break
            }
        }

        $fs.Close()
        return $buffer.ToArray()
    }
    catch {
        return $null
    }
}

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
 
    # Logging - fixed param names to match Write-Log / Write-LogProcess definitions:
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
        return [PSCustomObject]@{ File = $Path; Format = "NotFound"; Type = "Unknown" }
    }

    $fs = [System.IO.File]::OpenRead($Path)
    $buffer = New-Object byte[] 8
    $fs.Read($buffer, 0, 8) | Out-Null
    $fs.Close()

    # ZIP signature = OpenXML (unencrypted)
    if ($buffer[0] -eq 0x50 -and $buffer[1] -eq 0x4B -and
        $buffer[2] -eq 0x03 -and $buffer[3] -eq 0x04) {
        return [PSCustomObject]@{ File = $Path; Format = "OpenXML"; Type = "2007Plus (DOCX/XLSX/PPTX)" }
    }

    # OLE signature
    if ($buffer[0] -eq 0xD0 -and $buffer[1] -eq 0xCF -and $buffer[2] -eq 0x11 -and
        $buffer[3] -eq 0xE0 -and $buffer[4] -eq 0xA1 -and $buffer[5] -eq 0xB1 -and
        $buffer[6] -eq 0x1A -and $buffer[7] -eq 0xE1) {

        # Could be a legacy binary file, OR an encrypted OpenXML file.
        # Encrypted OpenXML is wrapped in OLE with EncryptedPackage + EncryptionInfo streams.
        # Check the extension to disambiguate — encrypted DOCX/XLSX/PPTX should still
        # be treated as OpenXML so they get routed to Process-OpenXmlBatch.
        $ext = [System.IO.Path]::GetExtension($Path).ToLower()
        $openXmlExtensions = @(".docx", ".docm", ".xlsx", ".xlsm", ".xlsb", ".pptx", ".pptm")

        if ($openXmlExtensions -contains $ext) {
            return [PSCustomObject]@{
                File   = $Path
                Format = "OpenXML"
                Type   = "2007Plus Encrypted (OLE-wrapped DOCX/XLSX/PPTX)"
            }
        }

        return [PSCustomObject]@{ File = $Path; Format = "BinaryOLE"; Type = "97-2003 (DOC/XLS/PPT)" }
    }

    return [PSCustomObject]@{ File = $Path; Format = "Unknown"; Type = "Unrecognized or Corrupt" }
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
        # 3. Patch [Content_Types].xml  - add Override for custom.xml
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
        # 4. Patch _rels/.rels - add Relationship for custom.xml
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
function Update-FileAgeProperties {
    param (
        [System.Collections.ArrayList]$Files,
        [string]$processedFiles,
        [string]$skippedFiles,
        [int]$openXmlParallelItems = 10,
        [int]$pdfParallelItems     = 10,
        [int]$comParallelItems     = 5,

        # Batch trigger thresholds   how many files accumulate in a queue
        # before that queue is dispatched as a batch. This is NOT a
        # concurrency control; the *ParallelItems values above still cap
        # how many runspaces execute simultaneously within a batch via
        # CreateRunspacePool(1, $parallelItems). Raising these thresholds
        # just means fewer, larger dispatch/drain cycles   less pool
        # open/close overhead   while parallelItems continues to throttle
        # how many files are actually being worked on at once.
        [int]$openXmlBatchTrigger = 100,   # was 3
        [int]$pdfBatchTrigger     = 100,   # was 3
        [int]$comBatchTrigger     = 30    # was 2
    )

    if (!(Test-FileExists -fileToTest $processedFiles)) {
        return
    }

    $debugFile = "$($env:LOCALAPPDATA)\$ScriptPath\debug.txt"
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

    $comQueue     = New-Object System.Collections.ArrayList
    $openXmlQueue = New-Object System.Collections.ArrayList
    $pdfQueue     = New-Object System.Collections.ArrayList

    # NOTE: $jobs is gone entirely   nothing accumulates across the run any more.
    # Each batch is dispatched and drained (waited on, disposed, pool closed)
    # immediately, in-line, before the foreach loop moves to the next file.

    $timestamp        = Get-Date -Format "yyyyMMdd_HHmmss"
    $filepath         = "$Propertylogfolderpath\$($timestamp)_AddPropertiesLog.txt"
    $filepathProgress = "$Propertystatusfolderpath\$($timestamp)_AddPropertiesStatus.txt"
    $metadataDuration = 100

    $logEntryProgress = @("Filename","StartTime","EndTime","Format","Filesize","PasswordProtected") -Join "|"
    Add-ContentSafe -Path $filepathProgress -Value $logEntryProgress

    New-Item -Path $filepath -ItemType File -Force | Out-Null

    $batchCounter = 0

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
            Write-Output "$file -> $format"
            if ($format -eq "OpenXML") {
                $openXmlQueue.Add($file) | Out-Null
            }
            elseif ($format -eq "BinaryOLE") {
                $comQueue.Add($file) | Out-Null
            }
        }

        Add-ContentSafe -Path $filepath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file queued for property update"

        # --- Batch triggers: dispatch, wait, dispose, close pool   all before continuing ---

        if ($openXmlQueue.Count -ge $openXmlBatchTrigger) {
            $openXmlQueueCopy = [System.Collections.ArrayList]@($openXmlQueue)
            $batchResult = Process-OpenXmlBatch `
                -batch            $openXmlQueueCopy `
                -metadataDuration $metadataDuration `
                -processedFiles   $processedFiles `
                -skippedFiles     $skippedFiles `
                -filepathProgress $filepathProgress `
                -format           $format `
                -filePathLog      $filepath `
                -parallelItems    $openXmlParallelItems

            Wait-AndCollectJobs -BatchResult $batchResult | Out-Null
            $openXmlQueue.Clear()
            $batchCounter++
        }

        if ($pdfQueue.Count -ge $pdfBatchTrigger) {
            $pdfQueueCopy = [System.Collections.ArrayList]@($pdfQueue)
            $batchResult = Process-PdfBatch `
                -batch            $pdfQueueCopy `
                -metadataDuration $metadataDuration `
                -processedFiles   $processedFiles `
                -skippedFiles     $skippedFiles `
                -filepathProgress $filepathProgress `
                -format           $format `
                -filePathLog      $filepath `
                -parallelItems    $pdfParallelItems

            Wait-AndCollectJobs -BatchResult $batchResult | Out-Null
            $pdfQueue.Clear()
            $batchCounter++
        }

        if ($comQueue.Count -ge $comBatchTrigger) {
            $comQueueCopy = [System.Collections.ArrayList]@($comQueue)
            $batchResult = Process-ComBatch `
                -batch            $comQueueCopy `
                -metadataDuration $metadataDuration `
                -processedFiles   $processedFiles `
                -skippedFiles     $skippedFiles `
                -filepathProgress $filepathProgress `
                -filePathLog      $filepath `
                -parallelItems    $comParallelItems

            Wait-AndCollectJobs -BatchResult $batchResult | Out-Null
            $comQueue.Clear()
            $batchCounter++
        }

        # --- Periodic GC nudge every 20 batches ---
        # Not done per-file (too costly); done at a batch-count interval so COM
        # and runspace-related memory gets reclaimed promptly across a long run
        # without constantly interrupting throughput.
        if ($batchCounter -gt 0 -and ($batchCounter % 20) -eq 0) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            $batchCounter = 0   # reset so the modulus check doesn't refire every iteration once past a multiple of 20
        }
    }

    # --- Drain any remaining partial batches ---

    if ($openXmlQueue.Count -gt 0) {
        $batchResult = Process-OpenXmlBatch `
            -batch            $openXmlQueue `
            -metadataDuration $metadataDuration `
            -processedFiles   $processedFiles `
            -skippedFiles     $skippedFiles `
            -filepathProgress $filepathProgress `
            -format           $format `
            -filePathLog      $filepath `
            -parallelItems    $openXmlParallelItems

        Wait-AndCollectJobs -BatchResult $batchResult | Out-Null
    }

    if ($pdfQueue.Count -gt 0) {
        $batchResult = Process-PdfBatch `
            -batch            $pdfQueue `
            -metadataDuration $metadataDuration `
            -processedFiles   $processedFiles `
            -skippedFiles     $skippedFiles `
            -filepathProgress $filepathProgress `
            -format           $format `
            -filePathLog      $filepath `
            -parallelItems    $pdfParallelItems

        Wait-AndCollectJobs -BatchResult $batchResult | Out-Null
    }

    if ($comQueue.Count -gt 0) {
        $batchResult = Process-ComBatch `
            -batch            $comQueue `
            -metadataDuration $metadataDuration `
            -processedFiles   $processedFiles `
            -skippedFiles     $skippedFiles `
            -filepathProgress $filepathProgress `
            -filePathLog      $filepath `
            -parallelItems    $comParallelItems

        Wait-AndCollectJobs -BatchResult $batchResult | Out-Null
    }

    # Final GC pass at the very end of the run
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}


# ---------------------------------------------------------------------------
# Wait-AndCollectJobs
# Shared helper   waits on every job in a batch result, collects output,
# disposes each Pipe, then closes and disposes the runspace pool itself.
# This is what makes the incremental draining above actually free memory:
# without explicitly closing $BatchResult.Pool, the pool's threads and
# buffers stay allocated even after every job in it has completed.
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
    if($batch -eq $null -or $batch.Count -eq 0) {
        Write-Warning "Process-PdfBatch called with null or empty batch - returning early"
        return @()
    }    
    #Python dependencies for PDF updates
    #$PythonPath = "C:\Program Files\Python313\python.exe"
    #$PythonScriptPath = "C:\Temp\update_pdf_properties.py"
    $PythonPath = "C:\Users\UDRTagging\AppData\Local\Python\pythoncore-3.14-64\python.exe"
    $PythonScriptPath = "$ScriptPath\update_pdf_properties_new.py"


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
                , $PythonScriptPath `
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


                $result = & $pythonPath $PythonScriptPath $fileToProcess `
                    "OriginalPath=$fileToProcess" `
                    "LastAccessed18Months=$strProperty18Months" `
                    "Created3Years=$strProperty3Years"

                switch ($result) {
                    1 {$isPasswordProtected = $true}
                    2 {$isPasswordProtected = $true}
                    -1 {$isError = $true}
                    0 {$processed = $true}
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
                        Write-Warning "Timestamp restore skipped; file in use: $fileToProcess ($restoreMsg)"
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
                    $logEntry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $fileToProcess properties NOT updated - $restoreMsg"
                    Add-ContentSafe -Path $filePathLog -Value $logEntry
                    Write-Output "$fileToProcess failed - see log"

                }
                else {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess properties updated"
                    Write-Output "$fileToProcess properties updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

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
        $ps.AddArgument($PythonScriptPath)| Out-Null
        $ps.AddArgument($fileToProcess)| Out-Null

        
        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }
 
    return [pscustomobject]@{
        Jobs = $jobs
        Pool = $pool
    }
}



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
    if($batch -eq $null -or $batch.Count -eq 0) {
        Write-Warning "Process-XmlBatch called with null or empty batch - returning early"
        return @()
    }    
 
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
            Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess locked before runspace - skipped"
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
                    Write-Output "$fileToProcess skipped - password-protected"
                }
                elseif ($isError) {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $fileToProcess properties NOT updated - see error above"
                    Write-Output "$fileToProcess failed - see log"
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
 
    return [pscustomobject]@{
        Jobs = $jobs
        Pool = $pool
    }
}



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
    if($batch -eq $null -or $batch.Count -eq 0) {
        Write-Warning "Process-ComBatch called with null or empty batch - returning early"
        return @()
    }    
 
    $pool = [runspacefactory]::CreateRunspacePool(1, $parallelItems)
    $pool.Open()
 
    # Capture all required function definitions once, outside the loop
    $fnReadOleStream        = "function Read-OleStream { ${function:Read-OleStream} }"
    $fnTestLegacyProtection = "function Test-LegacyOfficeProtection { ${function:Test-LegacyOfficeProtection} }"
    $fnSetOfficeDoc         = "function Set-OfficeDocCustomProperty { ${function:Set-OfficeDocCustomProperty} }"
    $fnTestEncrypted        = "function Test-OfficeEncrypted { ${function:Test-OfficeEncrypted} }"
    $fnTestEncryptedPpt2003 = "function Test-Ppt2003HasOpenPassword { ${function:Test-Ppt2003HasOpenPassword} }"
    $fnAddContentSafe       = "function Add-ContentSafe { ${function:Add-ContentSafe} }"
    $fnHandleError          = "function Handle-FileProcessingError { ${function:Handle-FileProcessingError} }"
    # Write-Log and Write-LogProcess are called by Handle-FileProcessingError - must also be injected
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
        $ps.AddScript($fnReadOleStream)          | Out-Null
        $ps.AddScript($fnTestLegacyProtection)   | Out-Null 
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
 
            # NEW:
            $protection = Test-LegacyOfficeProtection -Path $file
            if ($protection.IsPasswordToOpen -or $protection.IsPasswordToModify) {
                $isPasswordProtected = $true
                Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file skipped — $($protection.Reason)"
            } else {
                $isPasswordProtected = $false
            } 
            $dtLastAccessedDoc = $item.LastAccessTime
            $dtCreated         = $item.CreationTime
            $dtLastModified    = $item.LastWriteTime
            $fileReadOnly      = $item.IsReadOnly
            $filesize          = $item.Length
 
            try {
                $startTime         = Get-Date
                $startTimeF        = $startTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
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
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file skipped - password-protected"
                    Write-Output "$file skipped - password-protected"
                }
                elseif ($isError) {
                    Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $file properties NOT updated - see error above"
                    Write-Output "$file failed - see log"
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
        $ps.AddArgument($filePath)          | Out-Null   # the actual file - was $fileToProcess (undefined)
 
        $jobs += [pscustomobject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
        }
    }
 
    return [pscustomobject]@{
        Jobs = $jobs
        Pool = $pool
    }
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
            $_.LastAccessTime -lt (Get-Date).AddDays(-540) -and
            $_.CreationTime   -lt (Get-Date).AddDays(-1095) -and
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
    #$targetFolder = ($($FolderName.TrimStart('\').Replace('\','_'))).Replace('C:','_')
	# Define the file collection location
	$targetDir = "$ScriptPath\Unstructured\$($FolderName.TrimStart('\').Replace('\','_'))"
    #$targetDir = "$($env:LOCALAPPDATA)\Temp\Unstructured\$targetFolder"
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
		$skippedFilesList = (Get-Content -Path $skippedFiles).Trim()
		$skippedFilesListUnique = $skippedFilesList | sort -Unique
		if((Get-Item $scannedFiles).Length -ne 0) {
			$scannedFilesList = (Get-Content -Path $scannedFiles).Trim()
			foreach ($targetFile in $targetFilesList) {
				if($scannedFilesList.contains($targetFile)) {
				    Write-Host "$targetFile has already been scanned"
				}
                elseif($skippedFilesListUnique.contains($targetFile)) {
				    Write-Host "$targetFile has been skipped previously"
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
		try{
            Update-FileAgeProperties -Files $filesToScanUnique -ProcessedFiles $scannedFiles -skippedFiles $skippedFiles
        }
        catch {
            Add-ContentSafe -Path $filePathLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Update-FileAgeProperties failed without warning - $($_.Exception.Message)"
            Write-Output "Update-FileAgeProperties failed without warning - $($_.Exception.Message)"
        }
            
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
        [string]$LogPath = "$ScriptPath\KillProcess.log",
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
    $monitorProc = Start-KillProcessMonitor -MaxRuntimeSeconds 60 -CheckIntervalSeconds 15 -LogPath "$ScriptPath\KillProcess.log" -ShowWindow -NoExit
    #$monitorProc = Start-KillProcessMonitor -MaxRuntimeSeconds 60 -CheckIntervalSeconds 15 -LogPath "$($env:LOCALAPPDATA)\Temp\KillProcess.log" -ShowWindow -NoExit

    if ($null -eq $monitorProc) {
        Write-Error "Process kill monitor failed to start"
        Exit 1
    }

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

                Write-Host "Still running: $($remaining -join ', ') - waiting..."
                Start-Sleep -Milliseconds $pollIntervalMs
                $elapsed += $pollIntervalMs
            }

            if ($elapsed -ge ($drainTimeoutSecs * 1000)) {
                Write-Warning "Drain timeout reached ($drainTimeoutSecs s) - Office processes may still be running. Stopping monitor anyway."
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
# SIG # Begin signature block
# MIIFjgYJKoZIhvcNAQcCoIIFfzCCBXsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUo94POIZDVbX82kfiko1dbJeg
# hhGgggMeMIIDGjCCAgKgAwIBAgIQE7EWVmkfAJtDYwn54Vq9kjANBgkqhkiG9w0B
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU5UDiLdd+jqZv/o8SA8PCO78UQuUw
# DQYJKoZIhvcNAQEBBQAEggEAeMv/y59+L6BHm9oOHylcQokdDh/cgZ/8Lt4UFARq
# 5zQH6ucHY+wSeteKqQzDDD/Q7rAWfzgO/14OBEQK2LGE1OcfRBPeh7c1w8PltPqY
# FpXSUBQl8832sCdg8WABlovXCNbJhzc4EFUKwem9pMLdRGaBpEOv1DISTfd7Q1I+
# b2xjn32KvhRzsGv5GQWdc/j1eTYVlIe9MXvNne8V+esAC9tRFdYs5Sm3Qq1Fc5xW
# 7xVY+YOh0qX1yjTBA0qe9quTIFl0047Al70FPHKL72lodqSsMHpf5qLh+6n/TN0f
# zb6Y9ljFeclaG44xZQ2X8Z+tm0hNMUcwKi9LcEFP9W8RuQ==
# SIG # End signature block

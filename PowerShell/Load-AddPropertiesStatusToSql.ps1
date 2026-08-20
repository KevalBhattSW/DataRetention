<#
.SYNOPSIS
    Loads AddPropertiesStatus pipe-delimited (|) files into a SQL Server
    table, skipping files already loaded and reloading any that changed
    since their last successful load.

.DESCRIPTION
    Scans a folder for files matching *AddPropertiesStatus*.txt, checks a
    tracking table (by filename + SHA256 hash) to decide whether each file
    is New, Unchanged, or Changed since it was last loaded, and bulk-loads
    accordingly:
      - New       -> load
      - Unchanged -> skip (and zip the file in place, replacing it, since
                     its data is already safely loaded and won't change)
      - Changed   -> delete previously loaded rows for that file, then reload

.PARAMETER ReportsPath
    Folder containing the AddPropertiesStatus files.

.PARAMETER FileFilter
    Filter used to find source files. Default "*AddPropertiesStatus*.txt".

.PARAMETER SqlServer
    SQL Server instance name (e.g. "SQLPROD01" or "SQLPROD01\INSTANCE").

.PARAMETER Database
    Target database name.

.PARAMETER SchemaName
    Schema to create/use for both tables. Default "dbo".

.PARAMETER DataTable
    Table that holds the AddPropertiesStatus rows. Default "AddPropertiesStatus".

.PARAMETER LogTable
    Table that tracks which files have been loaded. Default "AddPropertiesStatusLoadLog".

.PARAMETER SqlCredential
    Optional PSCredential for SQL authentication. If omitted, Windows
    integrated authentication is used.

.PARAMETER BatchSize
    SqlBulkCopy batch size. Default 5000.

.EXAMPLE
    .\Load-AddPropertiesStatusToSql.ps1 -ReportsPath "\\AZUKSWVPUNSD01\Temp" `
        -SqlServer "SQLPROD01" -Database "OpsReporting"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ReportsPath,

    [Parameter(Mandatory=$false)]
    [string]$FileFilter = "*AddPropertiesStatus*.txt",

    [Parameter(Mandatory=$true)]
    [string]$SqlServer,

    [Parameter(Mandatory=$true)]
    [string]$Database,

    [Parameter(Mandatory=$false)]
    [string]$SchemaName = "dbo",

    [Parameter(Mandatory=$false)]
    [string]$DataTable = "AddPropertiesStatus",

    [Parameter(Mandatory=$false)]
    [string]$LogTable = "AddPropertiesStatusLoadLog",

    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$SqlCredential,

    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 5000
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ReportsPath -PathType Container)) {
    Write-Error "ReportsPath '$ReportsPath' does not exist or is not accessible. Aborting."
    Exit 1
}

# ---------------------------------------------------------------------------
# Connection helper
# ---------------------------------------------------------------------------
function New-SqlConnection {
    param(
        [string]$Server,
        [string]$Db,
        [System.Management.Automation.PSCredential]$Credential
    )

    if ($Credential) {
        $plainPwd = $Credential.GetNetworkCredential().Password
        $connString = "Server=$Server;Database=$Db;User Id=$($Credential.UserName);Password=$plainPwd;TrustServerCertificate=True;"
    }
    else {
        $connString = "Server=$Server;Database=$Db;Integrated Security=True;TrustServerCertificate=True;"
    }

    $conn = New-Object System.Data.SqlClient.SqlConnection($connString)
    $conn.Open()
    return $conn
}

# ---------------------------------------------------------------------------
# Ensure target tables exist
# ---------------------------------------------------------------------------
function Initialize-Tables {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Schema,
        [string]$DataTbl,
        [string]$LogTbl
    )

    $ddl = @"
IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = '$Schema')
    EXEC('CREATE SCHEMA [$Schema]');

IF NOT EXISTS (SELECT 1 FROM sys.tables t JOIN sys.schemas s ON t.schema_id = s.schema_id
               WHERE s.name = '$Schema' AND t.name = '$DataTbl')
BEGIN
    CREATE TABLE [$Schema].[$DataTbl] (
        Id                    BIGINT IDENTITY(1,1) PRIMARY KEY,
        Filename              NVARCHAR(1000)   NULL,
        StartTime             DATETIME2        NULL,
        StartTimeParseFailed  BIT              NOT NULL DEFAULT 0,
        StartTimeAmbiguous    BIT              NOT NULL DEFAULT 0,
        EndTime               DATETIME2        NULL,
        EndTimeParseFailed    BIT              NOT NULL DEFAULT 0,
        EndTimeAmbiguous      BIT              NOT NULL DEFAULT 0,
        Format                NVARCHAR(50)     NULL,
        FilesizeBytes         BIGINT           NULL,
        PasswordProtected     BIT              NULL,
        SourceFile            NVARCHAR(500)    NOT NULL,
        SourceFileCreatedUtc  DATETIME2        NULL,
        LoadedDateUtc         DATETIME2        NOT NULL DEFAULT SYSUTCDATETIME()
    );
    CREATE INDEX IX_${DataTbl}_SourceFile ON [$Schema].[$DataTbl] (SourceFile);
    CREATE INDEX IX_${DataTbl}_Filename ON [$Schema].[$DataTbl] (Filename);
END;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('$Schema.$DataTbl') AND name = 'StartTimeParseFailed')
    ALTER TABLE [$Schema].[$DataTbl] ADD StartTimeParseFailed BIT NOT NULL DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('$Schema.$DataTbl') AND name = 'StartTimeAmbiguous')
    ALTER TABLE [$Schema].[$DataTbl] ADD StartTimeAmbiguous BIT NOT NULL DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('$Schema.$DataTbl') AND name = 'EndTimeParseFailed')
    ALTER TABLE [$Schema].[$DataTbl] ADD EndTimeParseFailed BIT NOT NULL DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('$Schema.$DataTbl') AND name = 'EndTimeAmbiguous')
    ALTER TABLE [$Schema].[$DataTbl] ADD EndTimeAmbiguous BIT NOT NULL DEFAULT 0;
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('$Schema.$DataTbl') AND name = 'SourceFileCreatedUtc')
    ALTER TABLE [$Schema].[$DataTbl] ADD SourceFileCreatedUtc DATETIME2 NULL;

IF NOT EXISTS (SELECT 1 FROM sys.tables t JOIN sys.schemas s ON t.schema_id = s.schema_id
               WHERE s.name = '$Schema' AND t.name = '$LogTbl')
BEGIN
    CREATE TABLE [$Schema].[$LogTbl] (
        Id                   BIGINT IDENTITY(1,1) PRIMARY KEY,
        FileName             NVARCHAR(500)    NOT NULL,
        FullPath             NVARCHAR(1000)   NULL,
        FileHash             CHAR(64)         NOT NULL,
        FileCreatedUtc       DATETIME2        NULL,
        RowsLoaded           INT              NOT NULL,
        LoadedDateUtc        DATETIME2        NOT NULL DEFAULT SYSUTCDATETIME(),
        Status               NVARCHAR(20)     NOT NULL,
        ErrorMessage         NVARCHAR(MAX)    NULL,
        CONSTRAINT UQ_${LogTbl}_FileName UNIQUE (FileName)
    );
END;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('$Schema.$LogTbl') AND name = 'FileCreatedUtc')
    ALTER TABLE [$Schema].[$LogTbl] ADD FileCreatedUtc DATETIME2 NULL;
"@

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $ddl
    $cmd.ExecuteNonQuery() | Out-Null
}

# ---------------------------------------------------------------------------
# Determine load status for a file: New, Unchanged, or Changed
# ---------------------------------------------------------------------------
function Get-FileLoadStatus {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Schema,
        [string]$LogTbl,
        [string]$FileName,
        [string]$CurrentHash
    )

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = "SELECT TOP 1 FileHash FROM [$Schema].[$LogTbl] WHERE FileName = @FileName AND Status = 'Success'"
    $cmd.Parameters.AddWithValue("@FileName", $FileName) | Out-Null
    $existingHash = $cmd.ExecuteScalar()

    if ($null -eq $existingHash -or $existingHash -is [DBNull]) {
        return "New"
    }
    elseif ($existingHash -eq $CurrentHash) {
        return "Unchanged"
    }
    else {
        return "Changed"
    }
}

# ---------------------------------------------------------------------------
# Delete previously loaded rows for a file (used on reload)
# ---------------------------------------------------------------------------
function Remove-PreviousFileRows {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Schema,
        [string]$DataTbl,
        [string]$FileName
    )

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = "DELETE FROM [$Schema].[$DataTbl] WHERE SourceFile = @FileName"
    $cmd.Parameters.AddWithValue("@FileName", $FileName) | Out-Null
    return $cmd.ExecuteNonQuery()
}

# ---------------------------------------------------------------------------
# Zip an unchanged source file in place and delete the original, to reduce
# space used on the file server. The file's data is already safely loaded
# in SQL and, being unchanged, won't be reloaded — the zip replaces it as
# a compressed archive copy. Because the replacement is a .zip, it will no
# longer match -FileFilter on future runs, so it naturally drops out of
# scanning once archived.
# ---------------------------------------------------------------------------
function Compress-AndRemoveSourceFile {
    param([string]$FilePath)

    $zipPath = "$FilePath.zip"
    try {
        if (Test-Path -LiteralPath $zipPath) {
            Remove-Item -LiteralPath $zipPath -Force
        }
        Compress-Archive -LiteralPath $FilePath -DestinationPath $zipPath -CompressionLevel Optimal
        Remove-Item -LiteralPath $FilePath -Force
        Write-Host "       zipped and removed original -> $(Split-Path -Leaf $zipPath)"
    }
    catch {
        Write-Warning "       failed to zip/remove '$FilePath': $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Record the outcome of a load attempt
# ---------------------------------------------------------------------------
function Write-LoadLogEntry {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Schema,
        [string]$LogTbl,
        [string]$FileName,
        [string]$FullPath,
        [string]$FileHash,
        [Nullable[DateTime]]$FileCreatedUtc,
        [int]$RowsLoaded,
        [string]$Status,
        [string]$ErrorMessage = $null
    )

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = @"
MERGE [$Schema].[$LogTbl] AS target
USING (SELECT @FileName AS FileName) AS src
ON target.FileName = src.FileName
WHEN MATCHED THEN
    UPDATE SET FullPath = @FullPath, FileHash = @FileHash, FileCreatedUtc = @FileCreatedUtc,
               RowsLoaded = @RowsLoaded, LoadedDateUtc = SYSUTCDATETIME(),
               Status = @Status, ErrorMessage = @ErrorMessage
WHEN NOT MATCHED THEN
    INSERT (FileName, FullPath, FileHash, FileCreatedUtc, RowsLoaded, LoadedDateUtc, Status, ErrorMessage)
    VALUES (@FileName, @FullPath, @FileHash, @FileCreatedUtc, @RowsLoaded, SYSUTCDATETIME(), @Status, @ErrorMessage);
"@
    $cmd.Parameters.AddWithValue("@FileName", $FileName) | Out-Null
    $cmd.Parameters.AddWithValue("@FullPath", $FullPath) | Out-Null
    $cmd.Parameters.AddWithValue("@FileHash", $FileHash) | Out-Null
    $createdValue = if ($FileCreatedUtc) { $FileCreatedUtc } else { [DBNull]::Value }
    $cmd.Parameters.AddWithValue("@FileCreatedUtc", $createdValue) | Out-Null
    $cmd.Parameters.AddWithValue("@RowsLoaded", $RowsLoaded) | Out-Null
    $cmd.Parameters.AddWithValue("@Status", $Status) | Out-Null
    $errMsgValue = if ($ErrorMessage) { $ErrorMessage } else { [DBNull]::Value }
    $cmd.Parameters.AddWithValue("@ErrorMessage", $errMsgValue) | Out-Null
    $cmd.ExecuteNonQuery() | Out-Null
}

# ---------------------------------------------------------------------------
# Parse a StartTime/EndTime value that may be in US (M/d/yyyy) or UK
# (d/M/yyyy) format, or ISO (yyyy-MM-dd). Returns a hashtable:
#   Value        - parsed [datetime] or $null
#   ParseFailed  - true if neither culture could parse it
#   Ambiguous    - true if both cultures parsed it validly to DIFFERENT dates
#                  (day <= 12 in both positions). When ReferenceTime is
#                  supplied (the file's own timestamp), Value is resolved to
#                  whichever candidate is closer to it. Left NULL only if
#                  both candidates are exactly equidistant, or no
#                  ReferenceTime was available.
# ---------------------------------------------------------------------------
function ConvertTo-AmbiguousDateTime {
    param(
        [string]$RawValue,
        [DateTime]$ReferenceTime = [DateTime]::MinValue
    )

    $result = @{ Value = $null; ParseFailed = $false; Ambiguous = $false; IsBlank = $true }

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $result
    }

    $result.IsBlank = $false
    $RawValue = $RawValue.Trim()

    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $styles    = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces

    # ISO-style (yyyy-MM-dd...) is unambiguous — use general TryParse rather
    # than an exact literal pattern match. TryParse recognizes the ISO
    # sortable pattern natively and tolerates minor formatting variation
    # (single- vs double-digit components, odd whitespace, etc.).
    if ($RawValue -match '^\d{4}-\d{1,2}-\d{1,2}') {
        $dtIso = New-Object DateTime
        if ([datetime]::TryParse($RawValue, $invariant, $styles, [ref]$dtIso)) {
            $result.Value = $dtIso
        }
        else {
            $result.ParseFailed = $true
        }
        return $result
    }

    # Slash-style — genuinely ambiguous between US (M/d/yyyy) and UK (d/M/yyyy).
    # NOTE: each format is tried individually with the single-format
    # TryParseExact overload rather than passing an array — the array-format
    # overload does not bind reliably via PowerShell's method resolution in
    # this environment and silently fails to match even clearly-valid input.
    $usFormats = @("M/d/yyyy H:mm:ss", "MM/dd/yyyy HH:mm:ss", "M/d/yyyy h:mm:ss tt", "MM/dd/yyyy hh:mm:ss tt")
    $ukFormats = @("d/M/yyyy H:mm:ss", "dd/MM/yyyy HH:mm:ss", "d/M/yyyy h:mm:ss tt", "dd/MM/yyyy hh:mm:ss tt")

    $dtUs = New-Object DateTime
    $usOk = $false
    foreach ($fmt in $usFormats) {
        if ([datetime]::TryParseExact($RawValue, $fmt, $invariant, $styles, [ref]$dtUs)) {
            $usOk = $true
            break
        }
    }

    $dtUk = New-Object DateTime
    $ukOk = $false
    foreach ($fmt in $ukFormats) {
        if ([datetime]::TryParseExact($RawValue, $fmt, $invariant, $styles, [ref]$dtUk)) {
            $ukOk = $true
            break
        }
    }

    if ($usOk -and $ukOk) {
        if ($dtUs -eq $dtUk) {
            # e.g. day > 12 in one slot makes the format unambiguous anyway
            $result.Value = $dtUs
        }
        else {
            # genuinely ambiguous (day <= 12 in both positions) — resolve
            # using proximity to the file's own timestamp, if available
            $result.Ambiguous = $true
            if ($ReferenceTime -ne [DateTime]::MinValue) {
                $diffUs = [Math]::Abs(($dtUs - $ReferenceTime).TotalSeconds)
                $diffUk = [Math]::Abs(($dtUk - $ReferenceTime).TotalSeconds)
                if ($diffUs -lt $diffUk) {
                    $result.Value = $dtUs
                }
                elseif ($diffUk -lt $diffUs) {
                    $result.Value = $dtUk
                }
                else {
                    # exact tie — genuinely indeterminate
                    $result.Value = $null
                }
            }
            else {
                $result.Value = $null
            }
        }
    }
    elseif ($usOk) {
        $result.Value = $dtUs
    }
    elseif ($ukOk) {
        $result.Value = $dtUk
    }
    else {
        $result.ParseFailed = $true
    }

    return $result
}

# ---------------------------------------------------------------------------
# Load a single AddPropertiesStatus file via SqlBulkCopy
# Expected header: Filename|StartTime|EndTime|Format|Filesize|PasswordProtected
# StartTime/EndTime may be US, UK, or ISO formatted — see ConvertTo-AmbiguousDateTime
# ---------------------------------------------------------------------------
function Import-AddPropertiesStatusFile {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Schema,
        [string]$DataTbl,
        [string]$FilePath,
        [Nullable[DateTime]]$SourceFileCreatedUtc,
        [int]$BatchSize
    )

    $fileName  = Split-Path -Path $FilePath -Leaf
    $csvRows   = Import-Csv -LiteralPath $FilePath -Delimiter '|'

    # Filenames are stamped YYYYMMDD_HHMMSS — use that as a reference point
    # to resolve ambiguous US/UK dates (pick whichever interpretation is
    # closer in time to when the file itself was generated).
    $fileNameReferenceTime = [DateTime]::MinValue
    if ($fileName -match '(\d{8})_(\d{6})') {
        $stampText = "$($Matches[1])$($Matches[2])"
        $stampParsed = New-Object DateTime
        $stampOk = [datetime]::TryParseExact(
            $stampText, "yyyyMMddHHmmss",
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref]$stampParsed)
        if ($stampOk) {
            $fileNameReferenceTime = $stampParsed
        }
    }

    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add("Filename",              [string])
    [void]$table.Columns.Add("StartTime",             [datetime])
    [void]$table.Columns.Add("StartTimeParseFailed",  [bool])
    [void]$table.Columns.Add("StartTimeAmbiguous",    [bool])
    [void]$table.Columns.Add("EndTime",               [datetime])
    [void]$table.Columns.Add("EndTimeParseFailed",    [bool])
    [void]$table.Columns.Add("EndTimeAmbiguous",      [bool])
    [void]$table.Columns.Add("Format",                [string])
    [void]$table.Columns.Add("FilesizeBytes",         [int64])
    [void]$table.Columns.Add("PasswordProtected",     [bool])
    [void]$table.Columns.Add("SourceFile",            [string])
    [void]$table.Columns.Add("SourceFileCreatedUtc",  [datetime])
    $table.Columns["StartTime"].AllowDBNull = $true
    $table.Columns["EndTime"].AllowDBNull = $true
    $table.Columns["FilesizeBytes"].AllowDBNull = $true
    $table.Columns["PasswordProtected"].AllowDBNull = $true
    $table.Columns["SourceFileCreatedUtc"].AllowDBNull = $true

    $ambiguousCount = 0
    $failedCount = 0
    $sampleFailed = New-Object System.Collections.Generic.List[string]
    $sampleAmbiguous = New-Object System.Collections.Generic.List[string]

    foreach ($row in $csvRows) {
        $dr = $table.NewRow()
        $dr["Filename"]   = [string]$row.Filename
        $dr["Format"]     = [string]$row.Format
        $dr["SourceFile"] = $fileName
        if ($SourceFileCreatedUtc) {
            $dr["SourceFileCreatedUtc"] = $SourceFileCreatedUtc
        } else {
            $dr["SourceFileCreatedUtc"] = [DBNull]::Value
        }

        $start = ConvertTo-AmbiguousDateTime -RawValue $row.StartTime -ReferenceTime $fileNameReferenceTime
        if ($start.Value) {
            $dr["StartTime"] = $start.Value
        } else {
            $dr["StartTime"] = [DBNull]::Value
        }
        $dr["StartTimeParseFailed"] = $start.ParseFailed
        $dr["StartTimeAmbiguous"]   = $start.Ambiguous
        if ($start.ParseFailed) {
            $failedCount++
            if ($sampleFailed.Count -lt 5) { $sampleFailed.Add("StartTime='$($row.StartTime)'") }
        }
        if ($start.Ambiguous) {
            $ambiguousCount++
            if ($sampleAmbiguous.Count -lt 5) { $sampleAmbiguous.Add("StartTime='$($row.StartTime)'") }
        }

        $end = ConvertTo-AmbiguousDateTime -RawValue $row.EndTime -ReferenceTime $fileNameReferenceTime
        if ($end.Value) {
            $dr["EndTime"] = $end.Value
        } else {
            $dr["EndTime"] = [DBNull]::Value
        }
        $dr["EndTimeParseFailed"] = $end.ParseFailed
        $dr["EndTimeAmbiguous"]   = $end.Ambiguous
        if ($end.ParseFailed) {
            $failedCount++
            if ($sampleFailed.Count -lt 5) { $sampleFailed.Add("EndTime='$($row.EndTime)'") }
        }
        if ($end.Ambiguous) {
            $ambiguousCount++
            if ($sampleAmbiguous.Count -lt 5) { $sampleAmbiguous.Add("EndTime='$($row.EndTime)'") }
        }

        $parsedSize = 0L
        if ([int64]::TryParse($row.Filesize, [ref]$parsedSize)) {
            $dr["FilesizeBytes"] = $parsedSize
        } else {
            $dr["FilesizeBytes"] = [DBNull]::Value
        }

        $parsedBool = $false
        if ([bool]::TryParse($row.PasswordProtected, [ref]$parsedBool)) {
            $dr["PasswordProtected"] = $parsedBool
        } else {
            $dr["PasswordProtected"] = [DBNull]::Value
        }

        [void]$table.Rows.Add($dr)
    }

    if ($ambiguousCount -gt 0) {
        Write-Warning "       $ambiguousCount date value(s) in $fileName were ambiguous (US/UK both valid, different results) — resolved using proximity to the file's timestamp, flagged in StartTimeAmbiguous/EndTimeAmbiguous"
        Write-Warning "       sample: $($sampleAmbiguous -join '; ')"
    }
    if ($failedCount -gt 0) {
        Write-Warning "       $failedCount date value(s) in $fileName could not be parsed in any known format — set to NULL, flagged in StartTimeParseFailed/EndTimeParseFailed"
        Write-Warning "       sample: $($sampleFailed -join '; ')"
    }

    if ($table.Rows.Count -eq 0) {
        return 0
    }

    $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($Connection)
    $bulkCopy.DestinationTableName = "[$Schema].[$DataTbl]"
    $bulkCopy.BatchSize = $BatchSize
    foreach ($col in $table.Columns.ColumnName) {
        [void]$bulkCopy.ColumnMappings.Add($col, $col)
    }
    $bulkCopy.WriteToServer($table)
    $bulkCopy.Close()

    return $table.Rows.Count
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host "Connecting to $SqlServer / $Database ..."
$connection = New-SqlConnection -Server $SqlServer -Db $Database -Credential $SqlCredential

try {
    Write-Host "Ensuring tables [$SchemaName].[$DataTable] and [$SchemaName].[$LogTable] exist..."
    Initialize-Tables -Connection $connection -Schema $SchemaName -DataTbl $DataTable -LogTbl $LogTable

    $sourceFiles = Get-ChildItem -LiteralPath $ReportsPath -Filter $FileFilter -File -Recurse
    if ($sourceFiles.Count -eq 0) {
        Write-Host "No files matching '$FileFilter' found in $ReportsPath"
        return
    }

    Write-Host "Found $($sourceFiles.Count) file(s) matching '$FileFilter' in $ReportsPath"

    $loaded   = 0
    $skipped  = 0
    $reloaded = 0
    $failed   = 0

    foreach ($file in $sourceFiles) {

        $hash       = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $createdUtc = $file.CreationTimeUtc

        $status = Get-FileLoadStatus -Connection $connection -Schema $SchemaName `
                    -LogTbl $LogTable -FileName $file.Name -CurrentHash $hash

        if ($status -eq "Unchanged") {
            Write-Host "SKIP   $($file.Name) — already loaded, unchanged"
            $skipped++
            Compress-AndRemoveSourceFile -FilePath $file.FullName
            continue
        }

        if ($status -eq "Changed") {
            Write-Host "RELOAD $($file.Name) — content changed since last load, removing old rows"
            $removed = Remove-PreviousFileRows -Connection $connection -Schema $SchemaName `
                            -DataTbl $DataTable -FileName $file.Name
            Write-Host "       removed $removed previously loaded row(s)"
        }

        try {
            $rowCount = Import-AddPropertiesStatusFile -Connection $connection -Schema $SchemaName `
                            -DataTbl $DataTable -FilePath $file.FullName -SourceFileCreatedUtc $createdUtc `
                            -BatchSize $BatchSize

            Write-LoadLogEntry -Connection $connection -Schema $SchemaName -LogTbl $LogTable `
                -FileName $file.Name -FullPath $file.FullName -FileHash $hash -FileCreatedUtc $createdUtc `
                -RowsLoaded $rowCount -Status "Success"

            if ($status -eq "Changed") {
                Write-Host "LOADED $($file.Name) — $rowCount rows (reload)"
                $reloaded++
            }
            else {
                Write-Host "LOADED $($file.Name) — $rowCount rows"
                $loaded++
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-LoadLogEntry -Connection $connection -Schema $SchemaName -LogTbl $LogTable `
                -FileName $file.Name -FullPath $file.FullName -FileHash $hash -FileCreatedUtc $createdUtc `
                -RowsLoaded 0 -Status "Failed" -ErrorMessage $errMsg

            Write-Warning "FAILED $($file.Name) — $errMsg"
            $failed++
        }
    }

    Write-Host ""
    Write-Host "Done. Loaded: $loaded  Reloaded: $reloaded  Skipped: $skipped  Failed: $failed"
}
finally {
    $connection.Close()
    $connection.Dispose()
}
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
      - Unchanged -> skip
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
#   Ambiguous    - true if both cultures parsed it validly but to DIFFERENT
#                  dates (day <= 12 in both positions) — Value defaults to
#                  the US interpretation in this case, flagged for review
# ---------------------------------------------------------------------------
function ConvertTo-AmbiguousDateTime {
    param([string]$RawValue)

    $result = @{ Value = $null; ParseFailed = $true; Ambiguous = $false }

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $result
    }

    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $styles    = [System.Globalization.DateTimeStyles]::None

    $isoFormats = @("yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss")
    $usFormats  = @("M/d/yyyy H:mm:ss", "MM/dd/yyyy HH:mm:ss")
    $ukFormats  = @("d/M/yyyy H:mm:ss", "dd/MM/yyyy HH:mm:ss")

    # ISO first — unambiguous, so if it matches we're done
    $dtIso = New-Object DateTime
    if ([datetime]::TryParseExact($RawValue, $isoFormats, $invariant, $styles, [ref]$dtIso)) {
        $result.Value = $dtIso
        $result.ParseFailed = $false
        return $result
    }

    $dtUs = New-Object DateTime
    $usOk = [datetime]::TryParseExact($RawValue, $usFormats, $invariant, $styles, [ref]$dtUs)

    $dtUk = New-Object DateTime
    $ukOk = [datetime]::TryParseExact($RawValue, $ukFormats, $invariant, $styles, [ref]$dtUk)

    if ($usOk -and $ukOk) {
        $result.ParseFailed = $false
        if ($dtUs -eq $dtUk) {
            # e.g. day > 12 in one slot makes the format unambiguous anyway
            $result.Value = $dtUs
        }
        else {
            # genuinely ambiguous (day <= 12 in both positions) — default to
            # US, flag for review
            $result.Value = $dtUs
            $result.Ambiguous = $true
        }
    }
    elseif ($usOk) {
        $result.Value = $dtUs
        $result.ParseFailed = $false
    }
    elseif ($ukOk) {
        $result.Value = $dtUk
        $result.ParseFailed = $false
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

        $start = ConvertTo-AmbiguousDateTime -RawValue $row.StartTime
        if ($start.ParseFailed) {
            $dr["StartTime"] = [DBNull]::Value
            $failedCount++
        } else {
            $dr["StartTime"] = $start.Value
        }
        $dr["StartTimeParseFailed"] = $start.ParseFailed
        $dr["StartTimeAmbiguous"]   = $start.Ambiguous
        if ($start.Ambiguous) { $ambiguousCount++ }

        $end = ConvertTo-AmbiguousDateTime -RawValue $row.EndTime
        if ($end.ParseFailed) {
            $dr["EndTime"] = [DBNull]::Value
            $failedCount++
        } else {
            $dr["EndTime"] = $end.Value
        }
        $dr["EndTimeParseFailed"] = $end.ParseFailed
        $dr["EndTimeAmbiguous"]   = $end.Ambiguous
        if ($end.Ambiguous) { $ambiguousCount++ }

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
        Write-Warning "       $ambiguousCount date value(s) in $fileName were ambiguous (US/UK both valid, different results) — defaulted to US, flagged in StartTimeAmbiguous/EndTimeAmbiguous"
    }
    if ($failedCount -gt 0) {
        Write-Warning "       $failedCount date value(s) in $fileName could not be parsed in any known format — set to NULL, flagged in StartTimeParseFailed/EndTimeParseFailed"
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

    $sourceFiles = Get-ChildItem -LiteralPath $ReportsPath -Filter $FileFilter -File
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

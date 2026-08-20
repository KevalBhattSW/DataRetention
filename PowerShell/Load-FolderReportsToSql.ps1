<#
.SYNOPSIS
    Loads UDR-FolderReport.ps1 CSV output into a SQL Server table, skipping
    any file that has already been loaded.

.DESCRIPTION
    Scans a folder for *.csv files (the output of UDR-FolderReport.ps1),
    checks a tracking table to see if each file has already been imported,
    and bulk-loads any new ones. Each successfully loaded file is recorded
    in the tracking table (filename + SHA256 hash + row count + timestamp)
    so reruns are safe and cheap. A file found unchanged since its last
    successful load is zipped in place (replacing the original) to reduce
    space used on the file server.

.PARAMETER ReportsPath
    Folder containing the CSV files to load (e.g. the FolderReports dir
    produced by UDR-FolderReport.ps1).

.PARAMETER SqlServer
    SQL Server instance name (e.g. "SQLPROD01" or "SQLPROD01\INSTANCE").

.PARAMETER Database
    Target database name.

.PARAMETER SchemaName
    Schema to create/use for both tables. Default "dbo".

.PARAMETER DataTable
    Table that holds the folder report rows. Default "FolderReport".

.PARAMETER LogTable
    Table that tracks which files have been loaded. Default "FolderReportLoadLog".

.PARAMETER SqlCredential
    Optional PSCredential for SQL authentication. If omitted, Windows
    integrated authentication is used.

.PARAMETER BatchSize
    SqlBulkCopy batch size. Default 5000.

.EXAMPLE
    .\Load-FolderReportsToSql.ps1 -ReportsPath "D:\Scripts\FolderReports" `
        -SqlServer "SQLPROD01" -Database "OpsReporting"

.EXAMPLE
    $cred = Get-Credential
    .\Load-FolderReportsToSql.ps1 -ReportsPath "\\fileserver\FolderReports" `
        -SqlServer "SQLPROD01" -Database "OpsReporting" -SqlCredential $cred
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ReportsPath = "\\AZUKSWVPUNSD01\Temp\FolderReports",

    [Parameter(Mandatory=$true)]
    [string]$SqlServer = "LWUKWVNTV25\INS1,11433",

    [Parameter(Mandatory=$true)]
    [string]$Database = "unstrdata",

    [Parameter(Mandatory=$false)]
    [string]$SchemaName = "dbo",

    [Parameter(Mandatory=$false)]
    [string]$DataTable = "FolderReport",

    [Parameter(Mandatory=$false)]
    [string]$LogTable = "FolderReportLoadLog",

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
        Id            BIGINT IDENTITY(1,1) PRIMARY KEY,
        FolderPath    NVARCHAR(1000)   NULL,
        RelativePath  NVARCHAR(1000)   NULL,
        Depth         INT              NULL,
        Files         INT              NULL,
        Subfolders    INT              NULL,
        SizeMB        DECIMAL(18,2)    NULL,
        CountType     NVARCHAR(20)     NULL,
        PDF INT NULL,
        COM INT NULL,
        [OpenXML] INT NULL,
        Other INT NULL,
        SourceFile    NVARCHAR(500)    NOT NULL,
        LoadedDateUtc DATETIME2        NOT NULL DEFAULT SYSUTCDATETIME()
    );
    CREATE INDEX IX_${DataTbl}_SourceFile ON [$Schema].[$DataTbl] (SourceFile);
END;

IF NOT EXISTS (SELECT 1 FROM sys.tables t JOIN sys.schemas s ON t.schema_id = s.schema_id
               WHERE s.name = '$Schema' AND t.name = '$LogTbl')
BEGIN
    CREATE TABLE [$Schema].[$LogTbl] (
        Id            BIGINT IDENTITY(1,1) PRIMARY KEY,
        FileName      NVARCHAR(500)    NOT NULL,
        FullPath      NVARCHAR(1000)   NULL,
        FileHash      CHAR(64)         NOT NULL,
        RowsLoaded    INT              NOT NULL,
        LoadedDateUtc DATETIME2        NOT NULL DEFAULT SYSUTCDATETIME(),
        Status        NVARCHAR(20)     NOT NULL,
        ErrorMessage  NVARCHAR(MAX)    NULL,
        CONSTRAINT UQ_${LogTbl}_FileName UNIQUE (FileName)
    );
END;
"@

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $ddl
    $cmd.ExecuteNonQuery() | Out-Null
}

# ---------------------------------------------------------------------------
# Determine load status for a file: New, Unchanged, or Changed
# (Changed = same filename previously loaded successfully, but content hash
#  now differs — e.g. the file was overwritten mid-run)
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
# Delete previously loaded rows for a file (used when the file's content
# has changed since it was last loaded, so we can reload it cleanly)
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
# Remove a file, retrying with backoff if it's briefly locked. Small files
# in particular can still show as "in use" for a moment right after
# Compress-Archive finishes with them — the underlying handle (held by
# Compress-Archive itself, AV real-time scanning, or file-server/OneDrive
# sync) isn't always released the instant the cmdlet returns. Retrying
# beats failing the whole zip/delete step for what's usually a sub-second
# race.
# ---------------------------------------------------------------------------
function Remove-ItemWithRetry {
    param(
        [string]$Path,
        [int]$MaxAttempts = 6,
        [int]$InitialDelayMs = 250
    )

    $delay = $InitialDelayMs
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
            return
        }
        catch {
            if ($attempt -eq $MaxAttempts) { throw }
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, 4000)
        }
    }
}

# ---------------------------------------------------------------------------
# Zip an unchanged source file in place and delete the original, to reduce
# space used on the file server. The file's data is already safely loaded
# in SQL and, being unchanged, won't be reloaded — the zip replaces it as
# a compressed archive copy. Because the replacement is a .zip, it will no
# longer match the *.csv filter on future runs, so it naturally drops out
# of scanning once archived.
# ---------------------------------------------------------------------------
function Compress-AndRemoveSourceFile {
    param([string]$FilePath)

    $zipPath = "$FilePath.zip"
    try {
        if (Test-Path -LiteralPath $zipPath) {
            Remove-ItemWithRetry -Path $zipPath
        }
        Compress-Archive -LiteralPath $FilePath -DestinationPath $zipPath -CompressionLevel Optimal
        Remove-ItemWithRetry -Path $FilePath
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
        [int]$RowsLoaded,
        [string]$Status,
        [string]$ErrorMessage = $null
    )

    # MERGE so a previously-failed attempt for the same filename can be
    # updated to Success on a later successful retry, instead of violating
    # the unique constraint on FileName.
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = @"
MERGE [$Schema].[$LogTbl] AS target
USING (SELECT @FileName AS FileName) AS src
ON target.FileName = src.FileName
WHEN MATCHED THEN
    UPDATE SET FullPath = @FullPath, FileHash = @FileHash, RowsLoaded = @RowsLoaded,
               LoadedDateUtc = SYSUTCDATETIME(), Status = @Status, ErrorMessage = @ErrorMessage
WHEN NOT MATCHED THEN
    INSERT (FileName, FullPath, FileHash, RowsLoaded, LoadedDateUtc, Status, ErrorMessage)
    VALUES (@FileName, @FullPath, @FileHash, @RowsLoaded, SYSUTCDATETIME(), @Status, @ErrorMessage);
"@
    $cmd.Parameters.AddWithValue("@FileName", $FileName) | Out-Null
    $cmd.Parameters.AddWithValue("@FullPath", $FullPath) | Out-Null
    $cmd.Parameters.AddWithValue("@FileHash", $FileHash) | Out-Null
    $cmd.Parameters.AddWithValue("@RowsLoaded", $RowsLoaded) | Out-Null
    $cmd.Parameters.AddWithValue("@Status", $Status) | Out-Null
    $errMsgValue = if ($ErrorMessage) { $ErrorMessage } else { [DBNull]::Value }
    $cmd.Parameters.AddWithValue("@ErrorMessage", $errMsgValue) | Out-Null
    $cmd.ExecuteNonQuery() | Out-Null
}

# ---------------------------------------------------------------------------
# Load a single CSV via SqlBulkCopy
# ---------------------------------------------------------------------------
function Import-FolderReportCsv {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Schema,
        [string]$DataTbl,
        [string]$FilePath,
        [int]$BatchSize
    )

    $fileName = Split-Path -Path $FilePath -Leaf
    $csvRows  = Import-Csv -LiteralPath $FilePath

    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add("FolderPath",   [string])
    [void]$table.Columns.Add("RelativePath", [string])
    [void]$table.Columns.Add("Depth",        [int])
    [void]$table.Columns.Add("Files",        [int])
    [void]$table.Columns.Add("Subfolders",   [int])
    [void]$table.Columns.Add("SizeMB",       [decimal])
    [void]$table.Columns.Add("CountType",    [string])
    [void]$table.Columns.Add("PDF",        [int])
    [void]$table.Columns.Add("COM",        [int])
    [void]$table.Columns.Add("OpenXML",        [int])
    [void]$table.Columns.Add("Other",        [int])
    [void]$table.Columns.Add("SourceFile",   [string])

    foreach ($row in $csvRows) {
        $dr = $table.NewRow()
        $dr["FolderPath"]   = [string]$row.FolderPath
        $dr["RelativePath"] = [string]$row.RelativePath
        $dr["Depth"]        = [int]$row.Depth
        $dr["Files"]        = [int]$row.Files
        $dr["Subfolders"]   = [int]$row.Subfolders
        $dr["SizeMB"]       = [decimal]$row.SizeMB
        $dr["CountType"]    = [string]$row.CountType
        $dr["PDF"]        = [int]$row.PDF
        $dr["COM"]        = [int]$row.COM
        $dr["OpenXML"]        = [int]$row.OpenXML
        $dr["Other"]        = [int]$row.Other
        $dr["SourceFile"]   = $fileName
        [void]$table.Rows.Add($dr)
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

    $csvFiles = Get-ChildItem -LiteralPath $ReportsPath -Filter "*FolderReport_Depth6.csv" -File
    if ($csvFiles.Count -eq 0) {
        Write-Host "No CSV files found in $ReportsPath"
        return
    }

    Write-Host "Found $($csvFiles.Count) CSV file(s) in $ReportsPath"

    $loaded  = 0
    $skipped = 0
    $reloaded = 0
    $failed  = 0

    foreach ($file in $csvFiles) {

        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash

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
            $rowCount = Import-FolderReportCsv -Connection $connection -Schema $SchemaName `
                            -DataTbl $DataTable -FilePath $file.FullName -BatchSize $BatchSize

            Write-LoadLogEntry -Connection $connection -Schema $SchemaName -LogTbl $LogTable `
                -FileName $file.Name -FullPath $file.FullName -FileHash $hash `
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
                -FileName $file.Name -FullPath $file.FullName -FileHash $hash `
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

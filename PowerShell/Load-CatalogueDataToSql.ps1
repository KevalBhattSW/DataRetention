<#
.SYNOPSIS
    Stages already-unzipped tab-delimited text files into SQL Server,
    enriching each row with clean_extension, file_type, parsed dates,
    date-scope flags, SourceFile, and LoadDateTime.

.DESCRIPTION
    Workflow:
      1. Optionally creates the target table (a load-history tracking
         table is always ensured, regardless of -CreateTable; the staging
         table is always rebuilt — see Step 1 notes below).
      2. For each .txt file in SourceFolder:
           a. Read the file's size and last-write time (metadata only —
              does not download/hydrate a OneDrive/SharePoint placeholder).
           b. Compare those to the values recorded the last time this
              SourceFile was loaded (dbo.catalogue_data_load_history).
              - If both match: skip the file entirely (no copy, no load).
                The file is then zipped in place (replacing the original)
                to free up space on the file server — see note below.
              - If new or either has changed: proceed to load it.
           c. Copy the file into WorkFolder (this is what hydrates a
              not-yet-downloaded SharePoint/OneDrive file).
           d. SqlBulkCopy streams it into the staging table (client-side
              read, no server-side file access needed), mapped by column
              name rather than ordinal position.
           e. Any existing rows in the target table for this SourceFile are
              deleted (so a changed file fully replaces its old data).
           f. A single UPDATE enriches the staging rows.
           g. Rows are moved into the target table in batches of
              -InsertBatchSize (default 5000), each batch its own short
              transaction, rather than one INSERT...SELECT for the whole
              file — this keeps lock duration and log growth per batch
              small. An optional -InsertBatchDelayMs pause between
              batches can further throttle server impact.
           h. Staging table is truncated ready for the next file.
           i. The load-history table is updated with the new size/date so
              the next run can detect whether the file has changed again.
      3. Writes a CSV log and summary (including SKIPPED files).

.NOTES
    Change detection uses file size + last-write-time rather than a content
    hash, specifically to avoid hydrating files in a synced SharePoint/
    OneDrive library just to check whether they've changed — metadata can
    be read from an online-only placeholder without downloading it, but
    hashing content cannot. The trade-off is that a change which doesn't
    move the size or the last-write timestamp (rare, but possible with
    some sync/copy tools) won't be detected; use -ForceReload to bypass
    the check for a specific run if that's a concern.

    Zipping an unchanged file (see Step 2b) DOES require reading its full
    content, so it hydrates a not-yet-downloaded file just like loading
    would. This only happens once per file though: after it's zipped, the
    original is deleted and the .zip no longer matches the *.txt filter,
    so the file drops out of scanning on every subsequent run.
#>

[CmdletBinding()]
param(
    [string]$SqlServer   = "LWUKWVNTV25\INS1,11433",
    [string]$Database    = "unstrdata",
    [string]$TargetTable = "dbo.catalogue_data",
    [string]$StagingTable = "dbo.catalogue_data_staging",
    [string]$HistoryTable = "dbo.catalogue_data_load_history",

    [string]$SourceFolder = "C:\Users\BHATTK\OneDrive - RSA Group\Unstructured Data Remediation - File Listing\Batch 2",
    [string]$ZipFolder    = "C:\Users\BHATTK\RSA Group\Unstructured Data Remediation - PowerBIReports\Tagging\Data\File Listing\FilelistingZips",
    [string]$WorkFolder   = "C:\Temp\Work\",
    [string]$LogFile      = "C:\Temp\import_log.csv",

    [switch]$CreateTable = $true,
    # Default is FALSE: dropping the target table wipes previously loaded
    # data, but change-detection would still SKIP any file whose content
    # hasn't changed — silently losing that file's rows. Only opt into
    # dropping (-DropTableIfExists) when you want a full rebuild; when you
    # do, the load-history table is cleared too so nothing gets wrongly
    # skipped afterwards (see Step 1 below).
    [switch]$DropTableIfExists = $false,

    # Reprocess a file even if its hash matches the last recorded load.
    [switch]$ForceReload,

    # Number of rows moved from staging into the target table per batch
    # (see Step 2g). Smaller batches mean smaller transactions/log growth
    # and shorter lock durations, at the cost of more round trips.
    [int]$InsertBatchSize = 20000,

    # Optional pause between insert batches, in milliseconds, to give other
    # activity on the server a chance to run. 0 = no pause (back-to-back).
    [int]$InsertBatchDelayMs = 0,

    # After a file has been loaded, set it back to "Online-only" in the
    # synced SharePoint/OneDrive library (attrib +u), freeing the local
    # disk space that hydration used. With size+date change detection,
    # SKIPPED files are never hydrated in the first place, so this mainly
    # matters for files that were copied/loaded this run — it's still
    # applied to every file for safety (e.g. if something else on the
    # machine had already hydrated it), but is a no-op in that case. Off
    # by default since it changes the sync state of files outside the
    # script's own working folder.
    [switch]$DehydrateAfterProcessing,

    # Buffer size for the streamed copy used to report hydration/copy
    # progress (see Copy-FileWithProgress). Larger = fewer progress
    # updates and slightly less overhead; smaller = smoother updates.
    [int]$CopyBufferSizeMB = 4,

    # How many rows between SqlBulkCopy progress updates (staging load)
    # and between RAISERROR progress messages (target insert).
    [int]$ProgressNotifyRows = 20000,

    [char]$Delimiter  = "`t",
    [int]$BatchSize   = 20000
)

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------------
# CsvDataReader — reads tab-delimited files client-side for SqlBulkCopy.
# Appends SourceFile and LoadDateTime as virtual columns so the server
# never needs to see the file path.
# ------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'CsvDataReader').Type) {
Add-Type @"
using System;
using System.Data;
using System.Data.Common;
using System.IO;

public class CsvDataReader : DbDataReader
{
    private readonly StreamReader _reader;
    private readonly char         _delimiter;
    private readonly string       _sourceFile;
    private readonly DateTime     _loadDt;
    private readonly string[]     _headers;
    private string[]              _current;
    private bool                  _closed;

    private int DataColCount { get { return _headers.Length; } }
    public  int RowCount     { get; private set; }

    public CsvDataReader(string filePath, char delimiter)
    {
        _reader     = new StreamReader(filePath, System.Text.Encoding.UTF8, true);
        _delimiter  = delimiter;
        _sourceFile = System.IO.Path.GetFileName(filePath);
        _loadDt     = DateTime.UtcNow;

        var headerLine = _reader.ReadLine();
        _headers = headerLine == null ? new string[0] : headerLine.Split(_delimiter);
    }

    public override int FieldCount { get { return _headers.Length + 2; } }

    public override bool Read()
    {
        if (_closed) return false;
        var line = _reader.ReadLine();
        if (line == null) return false;
        _current = line.Split(_delimiter);
        RowCount++;
        return true;
    }

    public override object GetValue(int i)
    {
        if (i < DataColCount)
        {
            if (_current == null || i >= _current.Length) return DBNull.Value;
            var v = _current[i];
            return string.IsNullOrEmpty(v) ? (object)DBNull.Value : v;
        }
        if (i == DataColCount)     return _sourceFile;
        if (i == DataColCount + 1) return _loadDt.ToString("yyyy-MM-dd HH:mm:ss");
        throw new IndexOutOfRangeException();
    }

    public override string GetName(int i)
    {
        if (i < DataColCount)      return _headers[i];
        if (i == DataColCount)     return "SourceFile";
        if (i == DataColCount + 1) return "LoadDateTime";
        throw new IndexOutOfRangeException();
    }

    public override int GetOrdinal(string name)
    {
        for (int i = 0; i < _headers.Length; i++)
            if (string.Equals(_headers[i], name, StringComparison.OrdinalIgnoreCase)) return i;
        if (name == "SourceFile")   return DataColCount;
        if (name == "LoadDateTime") return DataColCount + 1;
        throw new IndexOutOfRangeException(name);
    }

    public override bool IsDBNull(int i)   { return GetValue(i) == DBNull.Value; }

    public override int GetValues(object[] values)
    {
        int n = Math.Min(values.Length, FieldCount);
        for (int i = 0; i < n; i++) values[i] = GetValue(i);
        return n;
    }

    public override void Close()          { _closed = true; _reader.Dispose(); }
    public override bool IsClosed         { get { return _closed; } }
    public override int  Depth            { get { return 0; } }
    public override int  RecordsAffected  { get { return -1; } }
    public override bool HasRows          { get { return true; } }
    public override bool NextResult()     { return false; }

    public override bool     GetBoolean(int i)  { return Convert.ToBoolean(GetValue(i)); }
    public override byte     GetByte(int i)     { return Convert.ToByte(GetValue(i)); }
    public override char     GetChar(int i)     { return Convert.ToChar(GetValue(i)); }
    public override Guid     GetGuid(int i)     { return Guid.Parse(GetValue(i).ToString()); }
    public override short    GetInt16(int i)    { return Convert.ToInt16(GetValue(i)); }
    public override int      GetInt32(int i)    { return Convert.ToInt32(GetValue(i)); }
    public override long     GetInt64(int i)    { return Convert.ToInt64(GetValue(i)); }
    public override float    GetFloat(int i)    { return Convert.ToSingle(GetValue(i)); }
    public override double   GetDouble(int i)   { return Convert.ToDouble(GetValue(i)); }
    public override decimal  GetDecimal(int i)  { return Convert.ToDecimal(GetValue(i)); }
    public override DateTime GetDateTime(int i) { return Convert.ToDateTime(GetValue(i)); }
    public override string   GetString(int i)   { return IsDBNull(i) ? null : GetValue(i).ToString(); }
    public override string   GetDataTypeName(int i) { return "nvarchar"; }
    public override Type     GetFieldType(int i)    { return typeof(string); }

    public override long GetBytes(int i, long fo, byte[] buf, int bo, int len) { return 0; }
    public override long GetChars(int i, long fo, char[] buf, int bo, int len) { return 0; }

    public override System.Collections.IEnumerator GetEnumerator()
    {
        return new DbEnumerator(this, true);
    }

    public override object this[int i]    { get { return GetValue(i); } }
    public override object this[string n] { get { return GetValue(GetOrdinal(n)); } }
}
"@ -ReferencedAssemblies "System.Data", "System.Xml"
}


$testReader = New-Object CsvDataReader(
    (Get-ChildItem $SourceFolder -Filter *.txt | Select -First 1).FullName, "`t")
#0..($testReader.FieldCount - 1) | ForEach-Object {
#    Write-Host "$_ : '$($testReader.GetName($_))'"
#}


# ------------------------------------------------------------------
# 0. Setup
# ------------------------------------------------------------------
foreach ($folder in @($WorkFolder, (Split-Path $LogFile -Parent))) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

$connectionString = "Server=$SqlServer;Database=$Database;Integrated Security=True;"

function Invoke-Sql {
    param([string]$Query, [int]$TimeoutSeconds = 0)
    $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSeconds
        $cmd.ExecuteNonQuery() | Out-Null
    }
    finally { $conn.Close() }
}

# Runs a query and returns the first column of the first row (or $null).
function Invoke-SqlScalar {
    param([string]$Query, [int]$TimeoutSeconds = 0)
    $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSeconds
        $result = $cmd.ExecuteScalar()
        if ($result -eq [DBNull]::Value -or $null -eq $result) { return $null }
        return $result
    }
    finally { $conn.Close() }
}

# Runs a query and returns the first row as a PSCustomObject (or $null if
# no rows). Used where more than one column is needed from a single row
# (e.g. comparing FileSizeBytes + FileLastWriteUtc together).
function Invoke-SqlRow {
    param([string]$Query, [int]$TimeoutSeconds = 0)
    $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSeconds
        $sqlReader = $cmd.ExecuteReader()
        try {
            if (-not $sqlReader.Read()) { return $null }
            $row = [ordered]@{}
            for ($i = 0; $i -lt $sqlReader.FieldCount; $i++) {
                $val = $sqlReader.GetValue($i)
                if ($val -eq [DBNull]::Value) { $val = $null }
                $row[$sqlReader.GetName($i)] = $val
            }
            return [PSCustomObject]$row
        }
        finally { $sqlReader.Close() }
    }
    finally { $conn.Close() }
}

$logEntries = New-Object System.Collections.Generic.List[object]
function Write-LogEntry {
    param([string]$File, [string]$Status, [string]$Detail, [int]$RowsLoaded = -1)
    $entry = [PSCustomObject]@{
        TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
        File         = $File
        Status       = $Status
        RowsLoaded   = $RowsLoaded
        Detail       = $Detail
    }
    $logEntries.Add($entry)
    $color = switch ($Status) {
        "SUCCESS" { "Green" }
        "SKIPPED" { "Yellow" }
        default   { "Red" }
    }
    Write-Host ("[{0}] {1} - {2}" -f $Status, $File, $Detail) -ForegroundColor $color
}

# Sets a synced SharePoint/OneDrive file back to "Online-only" (attrib +u),
# releasing the local disk space that hydration (e.g. Copy-FileWithProgress
# below) claimed for it. No-op / warns rather than throws, since a failed
# dehydration shouldn't abort the run — it's just housekeeping.
function Invoke-Dehydrate {
    param([string]$Path)
    try {
        $result = attrib.exe +u $Path 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "attrib +u returned exit code $LASTEXITCODE for '$Path': $result"
        }
    }
    catch {
        Write-Warning "Failed to dehydrate '$Path': $($_.Exception.Message)"
    }
}

# Removes a file, retrying with backoff if it's briefly locked. Small
# files in particular can still show as "in use" for a moment right after
# Compress-Archive finishes with them — the underlying handle (held by
# Compress-Archive itself, AV real-time scanning, or the OneDrive/
# SharePoint sync client) isn't always released the instant the cmdlet
# returns. Retrying beats failing the whole zip/delete step for what's
# usually a sub-second race.
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

# Zips an unchanged source file in place and deletes the original, to
# reduce space used on the file server. The file's data is already safely
# loaded in SQL and, being unchanged (same size + last-write-time as the
# last successful load), won't be reloaded — the zip replaces it as a
# compressed archive copy. Because the replacement is a .zip, it no longer
# matches the *.txt filter used in Step 2, so it naturally drops out of
# scanning on every subsequent run.
#
# NOTE: unlike the size+date check itself, this DOES read the file's full
# content, so it will hydrate a not-yet-downloaded SharePoint/OneDrive
# placeholder — but only once, since the file is removed from SourceFolder
# immediately afterwards.
function Compress-AndRemoveSourceFile {
    param([string]$FilePath)

    $zipPath = "$FilePath.zip"
    try {
        if (Test-Path -LiteralPath $zipPath) {
            Remove-ItemWithRetry -Path $zipPath
        }
        Compress-Archive -LiteralPath $FilePath -DestinationPath $zipPath -CompressionLevel Optimal
        Remove-ItemWithRetry -Path $FilePath
        Write-Host "         zipped and removed original -> $(Split-Path -Leaf $zipPath)" -ForegroundColor DarkYellow
    }
    catch {
        Write-Warning "         failed to zip/remove '$FilePath': $($_.Exception.Message)"
    }
}

# Streams SourcePath -> DestinationPath in chunks, reporting byte-level
# progress via Write-Progress as it goes. Used in place of Copy-Item so we
# get a progress bar; as a side effect this is also the only point at
# which a not-yet-downloaded SharePoint/OneDrive file actually gets
# hydrated, so the bar reflects download+copy together, not copy alone.
function Copy-FileWithProgress {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [int]$ProgressId,
        [int]$ParentId,
        [int]$BufferSizeBytes = 4MB
    )

    $sourceStream = [System.IO.File]::OpenRead($SourcePath)
    try {
        $destStream = [System.IO.File]::Create($DestinationPath)
        try {
            $totalBytes = $sourceStream.Length
            $buffer     = New-Object byte[] $BufferSizeBytes
            $totalRead  = 0L
            $lastUpdate = [DateTime]::MinValue

            while (($read = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $destStream.Write($buffer, 0, $read)
                $totalRead += $read

                # Throttle UI updates so Write-Progress itself isn't the bottleneck.
                if ((([DateTime]::UtcNow) - $lastUpdate).TotalMilliseconds -ge 200 -or $totalRead -eq $totalBytes) {
                    $pct = if ($totalBytes -gt 0) { [int](($totalRead / $totalBytes) * 100) } else { 100 }
                    Write-Progress -Id $ProgressId -ParentId $ParentId -Activity "Hydrating / copying" `
                        -Status ("{0:N1} MB of {1:N1} MB" -f ($totalRead / 1MB), ($totalBytes / 1MB)) `
                        -PercentComplete $pct
                    $lastUpdate = [DateTime]::UtcNow
                }
            }
        }
        finally { $destStream.Dispose() }
    }
    finally { $sourceStream.Dispose() }

    Write-Progress -Id $ProgressId -ParentId $ParentId -Activity "Hydrating / copying" -Completed
}

# ------------------------------------------------------------------
# 1. Create tables
# Note: staging table uses NVARCHAR for all columns so SqlBulkCopy
#       can push raw strings in without type conversion errors.
#       combined_date_scope / combined_date_scope_2026 / dates_updated /
#       scope_updated are BIT in the target but derived from the raw
#       staging strings in the INSERT...SELECT below, not stored as
#       strings.
#
#       The load-history table is intentionally created OUTSIDE the
#       -DropTableIfExists switch. By default (-DropTableIfExists:$false)
#       it persists indefinitely alongside the target table, which is
#       what makes change-detection possible across runs — if it got
#       wiped every run, every file would look "new" every time. It is
#       only ever cleared (not dropped) when -DropTableIfExists triggers
#       a full rebuild, to stay in sync with the now-empty target table.
# ------------------------------------------------------------------
if ($CreateTable) {
    Write-Host "`n=== Step 1: Creating tables ===" -ForegroundColor Cyan

    if ($DropTableIfExists) {
        Invoke-Sql -Query "IF OBJECT_ID('$TargetTable',  'U') IS NOT NULL DROP TABLE $TargetTable;"

        # The target table just lost all its data, so any "already loaded"
        # entries in history are now stale — without this, an unchanged
        # file would be SKIPPED and its data would never come back.
        # Clear (not drop) history so the table still exists for the
        # CREATE TABLE IF NOT EXISTS check further down, and every file
        # in this run is treated as new.
        Invoke-Sql -Query "IF OBJECT_ID('$HistoryTable', 'U') IS NOT NULL TRUNCATE TABLE $HistoryTable;"
        Write-Host "  Full rebuild requested: target dropped and load-history cleared." -ForegroundColor Yellow
    }

    # Staging always gets dropped/recreated, independent of -DropTableIfExists.
    # It holds no data across runs (each file TRUNCATEs it before loading),
    # so there's nothing to lose, and it guarantees the StagingId column
    # below always exists — needed to batch the staging->target INSERT.
    Invoke-Sql -Query "IF OBJECT_ID('$StagingTable', 'U') IS NOT NULL DROP TABLE $StagingTable;"

    $createSql = @"
IF OBJECT_ID('$TargetTable', 'U') IS NULL
BEGIN
    CREATE TABLE $TargetTable (
        Id                       INT           IDENTITY(1,1) PRIMARY KEY CLUSTERED,
        [Name]                   NVARCHAR(max) NULL,
        [Containing Path]        NVARCHAR(max) NULL,
        [Size]                   NVARCHAR(50)  NULL,
        [Last Modified]          NVARCHAR(255) NULL,
        [Last Accessed]          NVARCHAR(255) NULL,
        [Creation Date]          NVARCHAR(255) NULL,
        [Extension]              NVARCHAR(255) NULL,
        [Last Save Date]         NVARCHAR(255) NULL,
        [Date Checked]           NVARCHAR(255) NULL,
        clean_extension          NVARCHAR(255) NULL,
        file_type                NVARCHAR(255) NULL,
        combined_date_scope      BIT           NULL,
        combined_date_scope_2026 BIT           NULL,
        SourceFile               NVARCHAR(260) NULL,
        LoadDateTime             DATETIME2     NOT NULL DEFAULT SYSDATETIME(),
        LastModifiedDT           DATETIME2     NULL,
        LastAccessedDT           DATETIME2     NULL,
        CreationDateDT           DATETIME2     NULL,
        LastSaveDateDT           DATETIME2     NULL,
        DateCheckedDT            DATETIME2     NULL,
        dates_updated            BIT           NULL,
        scope_updated            BIT           NULL
    );
END

CREATE TABLE $StagingTable (
    StagingId        INT           IDENTITY(1,1) PRIMARY KEY CLUSTERED,
    [Name]           NVARCHAR(max) NULL,
    [Containing Path] NVARCHAR(max) NULL,
    [Size]           NVARCHAR(max)  NULL,
    [Last Modified]  NVARCHAR(max) NULL,
    [Last Accessed]  NVARCHAR(max) NULL,
    [Creation Date]  NVARCHAR(max) NULL,
    [Extension]      NVARCHAR(max) NULL,
    [Last Save Date] NVARCHAR(max) NULL,
    [Date Checked]   NVARCHAR(max) NULL,
    SourceFile       NVARCHAR(max) NULL,
    LoadDateTime     NVARCHAR(max)  NULL
);
"@
    Invoke-Sql -Query $createSql
    Write-Host "  Target/staging tables ready." -ForegroundColor Green
}

# Always ensure the history table exists, regardless of -CreateTable,
# so change-detection works even if the caller skips table creation.
$historySql = @"
IF OBJECT_ID('$HistoryTable', 'U') IS NULL
BEGIN
    CREATE TABLE $HistoryTable (
        SourceFile        NVARCHAR(260) NOT NULL PRIMARY KEY CLUSTERED,
        FileHash          CHAR(64)      NULL,   -- no longer populated; change detection now uses size+date. Retained for backward compatibility.
        FileSizeBytes     BIGINT        NULL,
        FileLastWriteUtc  DATETIME2     NULL,
        LastLoadDateTime  DATETIME2     NOT NULL DEFAULT SYSDATETIME(),
        RowsLoaded        INT           NULL
    );
END
"@
Invoke-Sql -Query $historySql

# Migration for history tables created before this change: FileHash used
# to be NOT NULL, but it's no longer populated. Safe/idempotent to run
# every time — a no-op if the column is already nullable.
Invoke-Sql -Query "ALTER TABLE $HistoryTable ALTER COLUMN FileHash CHAR(64) NULL;"

Write-Host "  Load-history table ready ($HistoryTable)." -ForegroundColor Green

# ------------------------------------------------------------------
# 2. Per-file: hash it, compare to history, copy/process only if
#    new or changed, delete it, move to next.
#    Only one copy of each file exists on disk at a time.
# ------------------------------------------------------------------
Write-Host "`n=== Step 2: Loading files (one at a time) ===" -ForegroundColor Cyan

$plainFiles = Get-ChildItem -Path $SourceFolder -Filter "*.txt" -File -ErrorAction SilentlyContinue |
              Sort-Object Name

$totalFileCount = $plainFiles.Count
$fileIndex      = 0

foreach ($f in $plainFiles) {
  $fileIndex++
  Write-Progress -Id 1 -Activity "Loading files" `
      -Status "File $fileIndex of $totalFileCount`: $($f.Name)" `
      -PercentComplete ([int]((($fileIndex - 1) / [Math]::Max($totalFileCount,1)) * 100))
  try {

    # ---- Change detection -----------------------------------------
    # Size + last-write-time only — both come from Get-ChildItem's
    # placeholder metadata for a synced SharePoint/OneDrive file, so
    # checking them never triggers a download. (A content hash would be
    # more airtight, but requires reading the file's bytes, which forces
    # hydration even for files that turn out to be unchanged.)
    $currentSize     = $f.Length
    $currentWriteUtc = $f.LastWriteTimeUtc.ToString("yyyy-MM-dd HH:mm:ss")

    $escapedName = $f.Name.Replace("'", "''")
    $previous = Invoke-SqlRow -Query "SELECT FileSizeBytes, CONVERT(NVARCHAR(19), FileLastWriteUtc, 120) AS FileLastWriteUtc FROM $HistoryTable WHERE SourceFile = '$escapedName';"

    $unchanged = $previous -and
                 ($null -ne $previous.FileSizeBytes) -and
                 ([int64]$previous.FileSizeBytes -eq [int64]$currentSize) -and
                 ($previous.FileLastWriteUtc -eq $currentWriteUtc)

    if (-not $ForceReload -and $unchanged) {
        Write-LogEntry -File $f.Name -Status "SKIPPED" -Detail "Unchanged since last load (size+date match)"
        Compress-AndRemoveSourceFile -FilePath $f.FullName
        continue
    }
    # ------------------------------------------------------------------

    # Copy single file into work folder — Copy-FileWithProgress (not
    # Copy-Item) so we get a live byte-level progress bar; this is also
    # the point where a not-yet-downloaded file actually hydrates.
    $dest = Join-Path $WorkFolder $f.Name
    Copy-FileWithProgress -SourcePath $f.FullName -DestinationPath $dest `
        -ProgressId 2 -ParentId 1 -BufferSizeBytes ($CopyBufferSizeMB * 1MB)

    $file      = Get-Item -Path $dest
    $csvReader = $null
    $bulkCopy  = $null
    try {
        # a) Truncate staging ready for this file
        Invoke-Sql -Query "TRUNCATE TABLE $StagingTable;"

        # b) Stream file into staging via SqlBulkCopy
        $csvReader = New-Object CsvDataReader($file.FullName, $Delimiter)

        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy(
            $connectionString,
            [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock
        )
        $bulkCopy.DestinationTableName = $StagingTable
        $bulkCopy.BatchSize            = $BatchSize
        $bulkCopy.BulkCopyTimeout      = 0
        $bulkCopy.NotifyAfter          = $ProgressNotifyRows

        # Row-count progress for the staging load. Total row count isn't
        # known in advance (would mean a separate full read of the file
        # just to count lines), so this is a running count rather than a
        # percentage. add_SqlRowsCopied attaches the scriptblock directly
        # as a delegate, so it fires synchronously on this thread during
        # WriteToServer below (unlike Register-ObjectEvent, which would
        # queue it and only run once WriteToServer returns).
        $bulkCopy.add_SqlRowsCopied({
            param($bcSender, $bcEventArgs)
            Write-Progress -Id 3 -ParentId 1 -Activity "Loading into staging" `
                -Status ("{0:N0} rows copied" -f $bcEventArgs.RowsCopied)
        })

        # Name-based mappings: CsvDataReader field name -> staging column name.
        # (Not ordinal/positional — the staging table has a leading
        # StagingId identity column, which shifts destination ordinals;
        # name mapping sidesteps that and isn't affected by future column
        # reordering either.)
        @("Name","Containing Path","Size","Last Modified","Last Accessed","Creation Date","Extension","Last Save Date","Date Checked","SourceFile","LoadDateTime") | ForEach-Object {
            $bulkCopy.ColumnMappings.Add($_, $_) | Out-Null
        }

        $diagReader = New-Object CsvDataReader($file.FullName,$delimiter)
        if ($diagReader.Read()) {
            0..($diagReader.FieldCount - 1) | ForEach-Object {
                $val = $diagReader.GetValue($_)
                #Write-Host "$_ : '$($diagReader.GetName($_))' = '$val' [$($val.GetType().Name)]"
            }
        }
        $diagReader.Close()

        $bulkCopy.WriteToServer($csvReader)
        Write-Progress -Id 3 -ParentId 1 -Activity "Loading into staging" -Completed
        $bulkCopy.Close()
        $csvReader.Close()

        # c) Delete any rows previously loaded for this SourceFile (in case
        #    it changed), enrich staging, then batch-insert into target.
        $enrichSql = @"
DELETE FROM $TargetTable WHERE SourceFile = '$escapedName';

UPDATE $StagingTable
SET
    SourceFile   = '$($file.Name)',
    LoadDateTime = CONVERT(NVARCHAR(50), SYSDATETIME(), 126);

-- Batched insert: move staging rows into the target table in chunks of
-- @BatchSize, keyed by StagingId, instead of one INSERT...SELECT covering
-- every row. Each iteration is its own short transaction, so locks are
-- held briefly and the transaction log grows incrementally rather than
-- all at once.
DECLARE @BatchSize INT = $InsertBatchSize;
DECLARE @MinId INT, @MaxId INT, @StartId INT, @TotalRows INT, @RowsDone INT = 0;

SELECT @MinId = MIN(StagingId), @MaxId = MAX(StagingId) FROM $StagingTable;
SET @StartId   = @MinId;
SET @TotalRows = ISNULL(@MaxId - @StartId + 1, 0);

WHILE @MinId IS NOT NULL AND @MinId <= @MaxId
BEGIN
    INSERT INTO $TargetTable (
        [Name], [Containing Path], [Size],
        [Last Modified], [Last Accessed], [Creation Date],
        [Extension], [Last Save Date], [Date Checked],
        clean_extension, file_type,
        SourceFile, LoadDateTime,
        LastModifiedDT
        ,LastAccessedDT
        ,CreationDateDT
        ,LastSaveDateDT
        ,DateCheckedDT
        ,dates_updated
        ,combined_date_scope
        ,combined_date_scope_2026
        ,scope_updated
    )
    SELECT
        [Name],
        [Containing Path],
        [Size],
        [Last Modified],
        [Last Accessed],
        [Creation Date],
        [Extension],
        [Last Save Date],
        [Date Checked],
        -- clean_extension
        CASE
            WHEN RIGHT(LOWER([Extension]),3) = 'pdf'  THEN '.pdf'
            WHEN RIGHT(LOWER([Extension]),4) = 'docx' THEN '.docx'
            WHEN RIGHT(LOWER([Extension]),4) = 'docm' THEN '.docm'
            WHEN RIGHT(LOWER([Extension]),3) = 'doc'  THEN '.doc'
            WHEN RIGHT(LOWER([Extension]),4) = 'xlsx' THEN '.xlsx'
            WHEN RIGHT(LOWER([Extension]),4) = 'xlsm' THEN '.xlsm'
            WHEN RIGHT(LOWER([Extension]),4) = 'xlsb' THEN '.xlsb'
            WHEN RIGHT(LOWER([Extension]),3) = 'xls'  THEN '.xls'
            WHEN RIGHT(LOWER([Extension]),4) = 'pptx' THEN '.pptx'
            WHEN RIGHT(LOWER([Extension]),4) = 'pptm' THEN '.pptm'
            WHEN RIGHT(LOWER([Extension]),3) = 'ppt'  THEN '.ppt'
            ELSE 'Other'
        END,
        -- file_type
        CASE
            WHEN RIGHT(LOWER([Extension]),4) IN ('docx','docm','xlsx','xlsm','xlsb','pptx','pptm') THEN 'OpenXML'
            WHEN RIGHT(LOWER([Extension]),3) IN ('doc','xls','ppt')                                THEN 'COM'
            WHEN RIGHT(LOWER([Extension]),3) = 'pdf'                                               THEN 'PDF'
            ELSE 'Other'
        END,
        SourceFile,
        SYSDATETIME(),
        -- parsed dates
        TRY_PARSE([Last Modified] AS DATETIME USING 'en-us'),
        TRY_PARSE([Last Accessed] AS DATETIME USING 'en-us'),
        TRY_PARSE([Creation Date] AS DATETIME USING 'en-us'),
        TRY_PARSE([Last Save Date] AS DATETIME USING 'en-us'),
        TRY_PARSE([Date Checked] AS DATETIME USING 'en-us'),
        1,  -- dates_updated
        -- combined_date_scope (date-scoped based on parsed dates)
        CAST(CASE
            WHEN TRY_CAST([Size] AS BIGINT) = 0 THEN 0
            WHEN LEFT([Name],1) = '~' THEN 0
            WHEN PATINDEX('%~snapshot%', [Containing Path]) > 0 THEN 0
            WHEN DATEADD(YEAR,-3,GETDATE()) < TRY_PARSE([Creation Date] AS DATETIME USING 'en-us')
              OR DATEADD(MONTH,-18,GETDATE()) < TRY_PARSE([Last Accessed] AS DATETIME USING 'en-us')
            THEN 0
            ELSE 1
        END AS BIT),
        -- combined_date_scope_2026
        CAST(CASE
            WHEN TRY_CAST([Size] AS BIGINT) = 0 THEN 0
            WHEN LEFT([Name],1) = '~' THEN 0
            WHEN PATINDEX('%~snapshot%', [Containing Path]) > 0 THEN 0
            WHEN DATEADD(YEAR,-3,DATEFROMPARTS(2026,12,31)) < TRY_PARSE([Creation Date] AS DATETIME USING 'en-us')
              OR DATEADD(MONTH,-18,DATEFROMPARTS(2026,12,31)) < TRY_PARSE([Last Accessed] AS DATETIME USING 'en-us')
            THEN 0
            ELSE 1
        END AS BIT),
        1  -- scope_updated
    FROM $StagingTable
    WHERE StagingId BETWEEN @MinId AND @MinId + @BatchSize - 1;

    SET @RowsDone = @RowsDone + @@ROWCOUNT;

    -- Progress message picked up by the connection's InfoMessage event in
    -- PowerShell (see below). WITH NOWAIT forces immediate delivery to the
    -- client rather than buffering until the batch finishes.
    RAISERROR('PROGRESS:%d:%d', 0, 1, @RowsDone, @TotalRows) WITH NOWAIT;

    SET @MinId = @MinId + @BatchSize;
$(if ($InsertBatchDelayMs -gt 0) { "    WAITFOR DELAY '$([TimeSpan]::FromMilliseconds($InsertBatchDelayMs).ToString('hh\:mm\:ss\.fff'))';" })
END

SELECT COUNT(*) FROM $StagingTable;
"@

        $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $conn.FireInfoMessageEventOnUserErrors = $false

        # Live batch-level progress for the target insert. Like
        # add_SqlRowsCopied above, add_InfoMessage attaches the scriptblock
        # as a real delegate, so it fires synchronously as each RAISERROR
        # WITH NOWAIT arrives during ExecuteReader below — not deferred
        # the way Register-ObjectEvent would be.
        $conn.add_InfoMessage({
            param($icSender, $icEventArgs)
            foreach ($sqlErr in $icEventArgs.Errors) {
                if ($sqlErr.Message -like "PROGRESS:*") {
                    $parts = ($sqlErr.Message -replace '^PROGRESS:','') -split ':'
                    $done  = [int64]$parts[0]
                    $total = [int64]$parts[1]
                    $pct   = if ($total -gt 0) { [int](($done / $total) * 100) } else { 100 }
                    Write-Progress -Id 4 -ParentId 1 -Activity "Inserting into target table" `
                        -Status ("{0:N0} of {1:N0} rows" -f $done, $total) -PercentComplete $pct
                }
            }
        })

        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $enrichSql
        $cmd.CommandTimeout = 0
        $reader = $cmd.ExecuteReader()
        $rows = 0
        if ($reader.Read()) { $rows = $reader.GetInt32(0) }
        $reader.Close()
        $conn.Close()
        Write-Progress -Id 4 -ParentId 1 -Activity "Inserting into target table" -Completed

        # d) Record/refresh this file's size + last-write-time in the
        #    history table so the next run can tell whether it has changed.
        $historyUpsertSql = @"
MERGE $HistoryTable AS tgt
USING (SELECT
           '$escapedName'                          AS SourceFile,
           $currentSize                             AS FileSizeBytes,
           CONVERT(DATETIME2, '$currentWriteUtc')   AS FileLastWriteUtc,
           $rows                                    AS RowsLoaded
       ) AS src
ON (tgt.SourceFile = src.SourceFile)
WHEN MATCHED THEN
    UPDATE SET FileSizeBytes = src.FileSizeBytes,
               FileLastWriteUtc = src.FileLastWriteUtc,
               LastLoadDateTime = SYSDATETIME(),
               RowsLoaded = src.RowsLoaded
WHEN NOT MATCHED THEN
    INSERT (SourceFile, FileSizeBytes, FileLastWriteUtc, LastLoadDateTime, RowsLoaded)
    VALUES (src.SourceFile, src.FileSizeBytes, src.FileLastWriteUtc, SYSDATETIME(), src.RowsLoaded);
"@
        Invoke-Sql -Query $historyUpsertSql

        Write-LogEntry -File $file.Name -Status "SUCCESS" -Detail "Loaded OK (new or changed file)" -RowsLoaded $rows
    }
    catch {
        Write-LogEntry -File $file.Name -Status "FAILED" -Detail $_.Exception.Message
    }
    finally {
        if ($bulkCopy  -and -not $bulkCopy.GetType().GetMethod('IsClosed')) { try { $bulkCopy.Close()  } catch {} }
        if ($csvReader -and -not $csvReader.IsClosed)                       { try { $csvReader.Close() } catch {} }

        # Delete the work-folder copy regardless of success/failure so disk
        # space is freed before the next file is copied in.
        if (Test-Path $dest) { Remove-Item -Path $dest -Force }
    }
  }
  finally {
    # Runs for every file regardless of outcome (skipped/succeeded/failed).
    # With size+date change detection the SKIPPED path never hydrates the
    # file, so this is mostly relevant for files that were copied/loaded
    # this run — harmless no-op otherwise.
    if ($DehydrateAfterProcessing) {
        Invoke-Dehydrate -Path $f.FullName
    }
  }
}

# Clear all progress bars (overall + the three per-file sub-stages) now
# that every file has been processed.
Write-Progress -Id 1 -Activity "Loading files" -Completed
Write-Progress -Id 2 -ParentId 1 -Activity "Hydrating / copying" -Completed
Write-Progress -Id 3 -ParentId 1 -Activity "Loading into staging" -Completed
Write-Progress -Id 4 -ParentId 1 -Activity "Inserting into target table" -Completed

# ------------------------------------------------------------------
# 4. Log and summary
# ------------------------------------------------------------------
$logEntries | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8

$successCount = ($logEntries | Where-Object Status -eq "SUCCESS").Count
$skippedCount = ($logEntries | Where-Object Status -eq "SKIPPED").Count
$failCount    = ($logEntries | Where-Object Status -eq "FAILED").Count
$totalRows    = ($logEntries | Where-Object Status -eq "SUCCESS" | Measure-Object -Property RowsLoaded -Sum).Sum

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "  Succeeded : $successCount" -ForegroundColor Green
Write-Host "  Skipped   : $skippedCount (unchanged since last load, zipped in place)" -ForegroundColor Yellow
Write-Host "  Failed    : $failCount"    -ForegroundColor $(if ($failCount -gt 0) {"Red"} else {"Green"})
Write-Host "  Total rows: $totalRows"
Write-Host "  Log       : $LogFile"

if ($failCount -gt 0) {
    Write-Host "`n  Review $LogFile for details on failed files." -ForegroundColor Yellow
}

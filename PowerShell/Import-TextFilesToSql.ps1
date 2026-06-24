<#
.SYNOPSIS
    Stages already-unzipped tab-delimited text files into SQL Server,
    enriching each row with clean_extension, file_type, date-scope flags,
    SourceFile, and LoadDateTime.

.DESCRIPTION
    Workflow:
      1. Optionally creates the target and staging tables.
      2. Copies .txt files from SourceFolder into WorkFolder.
      3. For each file:
           a. SqlBulkCopy streams it into the staging table (client-side read,
              no server-side file access needed).
           b. A single UPDATE enriches the staging rows.
           c. INSERT...SELECT moves them into the target table.
           d. Staging table is truncated ready for the next file.
      4. Writes a CSV log and summary.
#>

[CmdletBinding()]
param(
    [string]$SqlServer   = "LWUKWVNTV25\INS1,11433",
    [string]$Database    = "unstrdata",
    [string]$TargetTable = "dbo.catalogue_data",
    [string]$StagingTable = "dbo.catalogue_data_staging",

    [string]$SourceFolder = "C:\Users\BHATTK\RSA Group\Unstructured Data Remediation - PowerBIReports\Tagging\Data\File Listing",
    [string]$ZipFolder    = "C:\Users\BHATTK\RSA Group\Unstructured Data Remediation - PowerBIReports\Tagging\Data\File Listing\FilelistingZips",
    [string]$WorkFolder   = "C:\Temp\Work\",
    [string]$LogFile      = "C:\Temp\import_log.csv",

    [switch]$CreateTable,
    [switch]$DropTableIfExists,

    [char]$Delimiter  = "`t",
    [int]$BatchSize   = 5000
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
        if (i == DataColCount + 1) return _loadDt;
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
    $color = if ($Status -eq "SUCCESS") { "Green" } else { "Red" }
    Write-Host ("[{0}] {1} - {2}" -f $Status, $File, $Detail) -ForegroundColor $color
}

# ------------------------------------------------------------------
# 1. Create tables
# Note: staging table uses NVARCHAR for all columns so SqlBulkCopy
#       can push raw strings in without type conversion errors.
#       combined_date_scope columns are BIT in the target but derived
#       via CAST in the INSERT...SELECT, not stored as strings.
# ------------------------------------------------------------------
if ($CreateTable) {
    Write-Host "`n=== Step 1: Creating tables ===" -ForegroundColor Cyan

    if ($DropTableIfExists) {
        Invoke-Sql -Query "IF OBJECT_ID('$TargetTable',  'U') IS NOT NULL DROP TABLE $TargetTable;"
        Invoke-Sql -Query "IF OBJECT_ID('$StagingTable', 'U') IS NOT NULL DROP TABLE $StagingTable;"
    }

    $createSql = @"
IF OBJECT_ID('$TargetTable', 'U') IS NULL
BEGIN
    CREATE TABLE $TargetTable (
        Id                       INT           IDENTITY(1,1) PRIMARY KEY CLUSTERED,
        [Name]                   NVARCHAR(500) NULL,
        [Containing Path]        NVARCHAR(255) NULL,
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
        LoadDateTime             DATETIME2     NOT NULL DEFAULT SYSDATETIME()
    );
END

IF OBJECT_ID('$StagingTable', 'U') IS NULL
BEGIN
    CREATE TABLE $StagingTable (
        [Name]           NVARCHAR(500) NULL,
        [Containing Path] NVARCHAR(255) NULL,
        [Size]           NVARCHAR(50)  NULL,
        [Last Modified]  NVARCHAR(255) NULL,
        [Last Accessed]  NVARCHAR(255) NULL,
        [Creation Date]  NVARCHAR(255) NULL,
        [Extension]      NVARCHAR(255) NULL,
        [Last Save Date] NVARCHAR(255) NULL,
        [Date Checked]   NVARCHAR(255) NULL,
        SourceFile       NVARCHAR(260) NULL,
        LoadDateTime     NVARCHAR(50)  NULL
    );
END
"@
    Invoke-Sql -Query $createSql
    Write-Host "  Tables ready." -ForegroundColor Green
}

# ------------------------------------------------------------------
# 2. Per-file: copy one file, process it, delete it, move to next
#    Only one copy of each file exists on disk at a time.
# ------------------------------------------------------------------
Write-Host "`n=== Step 2: Loading files (one at a time) ===" -ForegroundColor Cyan

$plainFiles = Get-ChildItem -Path $SourceFolder -Filter "*.txt" -File -ErrorAction SilentlyContinue |
              Sort-Object Name

foreach ($f in $plainFiles) {

    # Copy single file into work folder
    $dest = Join-Path $WorkFolder $f.Name
    Copy-Item -Path $f.FullName -Destination $dest -Force

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


        # Explicit mappings: file column name -> staging table column name.
        @("Name","Containing Path","Size","Last Modified","Last Accessed","Creation Date","Extension","Last Save Date","Date Checked","SourceFile","LoadDateTime") | ForEach-Object {
            $bulkCopy.ColumnMappings.Add($_, $_) | Out-Null
        }

        $bulkCopy.WriteToServer($csvReader)
        $bulkCopy.Close()
        $csvReader.Close()

        # c) Enrich staging + insert into target in one SQL batch
        $enrichSql = @"
UPDATE $StagingTable
SET
    SourceFile   = '$($file.Name)',
    LoadDateTime = CONVERT(NVARCHAR(50), SYSDATETIME(), 126);

INSERT INTO $TargetTable (
    [Name], [Containing Path], [Size],
    [Last Modified], [Last Accessed], [Creation Date],
    [Extension], [Last Save Date], [Date Checked],
    clean_extension, file_type,
    combined_date_scope, combined_date_scope_2026,
    SourceFile, LoadDateTime
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
    -- combined_date_scope
    CAST(CASE
        WHEN DATEADD(YEAR,-3,CAST(GETDATE() AS DATE)) > TRY_CAST([Creation Date] AS DATE)
         AND DATEADD(MONTH,-18,CAST(GETDATE() AS DATE)) > TRY_CAST([Last Accessed] AS DATE)
        THEN 1 ELSE 0
    END AS BIT),
    -- combined_date_scope_2026
    CAST(CASE
        WHEN DATEADD(YEAR,-3,DATEFROMPARTS(2026,12,31)) > TRY_CAST([Creation Date] AS DATE)
         AND DATEADD(MONTH,-18,DATEFROMPARTS(2026,12,31)) > TRY_CAST([Last Accessed] AS DATE)
        THEN 1 ELSE 0
    END AS BIT),
    SourceFile,
    SYSDATETIME()
FROM $StagingTable;

SELECT COUNT(*) FROM $StagingTable;
"@

        $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $enrichSql
        $cmd.CommandTimeout = 0
        $reader = $cmd.ExecuteReader()
        $rows = 0
        if ($reader.Read()) { $rows = $reader.GetInt32(0) }
        $reader.Close()
        $conn.Close()

        Write-LogEntry -File $file.Name -Status "SUCCESS" -Detail "Loaded OK" -RowsLoaded $rows
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

# ------------------------------------------------------------------
# 4. Log and summary
# ------------------------------------------------------------------
$logEntries | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8

$successCount = ($logEntries | Where-Object Status -eq "SUCCESS").Count
$failCount    = ($logEntries | Where-Object Status -eq "FAILED").Count
$totalRows    = ($logEntries | Where-Object Status -eq "SUCCESS" | Measure-Object -Property RowsLoaded -Sum).Sum

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "  Succeeded : $successCount" -ForegroundColor Green
Write-Host "  Failed    : $failCount"    -ForegroundColor $(if ($failCount -gt 0) {"Red"} else {"Green"})
Write-Host "  Total rows: $totalRows"
Write-Host "  Log       : $LogFile"

if ($failCount -gt 0) {
    Write-Host "`n  Review $LogFile for details on failed files." -ForegroundColor Yellow
}

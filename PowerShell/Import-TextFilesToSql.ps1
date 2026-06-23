<#
.SYNOPSIS
    Unzips archived text files, creates a target table, and bulk-loads all
    pipe-delimited text files (zipped + unzipped) into a single SQL Server table.

.DESCRIPTION
    Workflow:
      1. Unzips any .zip archives found in -ZipFolder into -WorkFolder.
      2. Copies/uses already-unzipped .txt files from -SourceFolder into -WorkFolder.
      3. Optionally creates the target table (edit the CREATE TABLE block below
         to match your real columns before running with -CreateTable).
      4. Loops over every .txt file in -WorkFolder and streams it into SQL Server
         using SqlBulkCopy — file is read client-side, so the SQL Server machine
         never needs to see or share the file path.
      5. Logs success/failure per file to a log file and to the console.

.NOTES
    Auth      : Windows Authentication (Integrated Security).
    Delimiter : pipe '|'. Change $Delimiter below if yours differs.
    SqlBulkCopy reads and parses files on this machine, then streams rows over
    the connection — no UNC shares or server-side file access required.
#>

[CmdletBinding()]
param(
    [string]$SqlServer    = "localhost",
    [string]$Database     = "YourDatabase",
    [string]$TargetTable  = "dbo.ImportedData",

    [string]$SourceFolder = "C:\Import\UnzippedFiles",  # the ~40 already-unzipped .txt files
    [string]$ZipFolder    = "C:\Import\ZippedFiles",    # the ~15 .zip archives
    [string]$WorkFolder   = "C:\Import\Work",           # everything gets staged/unzipped here
    [string]$LogFile      = "C:\Import\import_log.csv",

    [switch]$CreateTable,       # pass this switch to (re)create the table first
    [switch]$DropTableIfExists, # pass alongside -CreateTable to DROP first

    [char]$Delimiter      = '|',
    [int]$BatchSize       = 5000  # rows sent to SQL Server per network round-trip
)

$ErrorActionPreference = "Stop"

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
# 1. Unzip archives into the work folder
# ------------------------------------------------------------------
Write-Host "`n=== Step 1: Unzipping archives ===" -ForegroundColor Cyan

$zipFiles = Get-ChildItem -Path $ZipFolder -Filter "*.zip" -File -ErrorAction SilentlyContinue
foreach ($zip in $zipFiles) {
    try {
        Write-Host "  Extracting $($zip.Name)..."
        Expand-Archive -Path $zip.FullName -DestinationPath $WorkFolder -Force
        Write-LogEntry -File $zip.Name -Status "SUCCESS" -Detail "Unzipped"
    }
    catch {
        Write-LogEntry -File $zip.Name -Status "FAILED" -Detail "Unzip error: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------
# 2. Stage the already-unzipped .txt files into WorkFolder
# ------------------------------------------------------------------
Write-Host "`n=== Step 2: Staging already-unzipped files ===" -ForegroundColor Cyan

$plainFiles = Get-ChildItem -Path $SourceFolder -Filter "*.txt" -File -ErrorAction SilentlyContinue
foreach ($f in $plainFiles) {
    $dest = Join-Path $WorkFolder $f.Name
    if (-not (Test-Path $dest)) {
        Copy-Item -Path $f.FullName -Destination $dest
    }
}
Write-Host "  $($plainFiles.Count) file(s) staged."

# ------------------------------------------------------------------
# 3. Optionally create the target table
# ------------------------------------------------------------------
if ($CreateTable) {
    Write-Host "`n=== Step 3: Creating target table ===" -ForegroundColor Cyan

    if ($DropTableIfExists) {
        Invoke-Sql -Query "IF OBJECT_ID('$TargetTable', 'U') IS NOT NULL DROP TABLE $TargetTable;"
    }

    # >>> EDIT THIS block to match your actual columns and types.       <<<
    # >>> Columns must be listed in the same order as they appear in    <<<
    # >>> the files. SourceFile and LoadDateTime are appended by the    <<<
    # >>> script — do NOT add them to your source files.                <<<
    #
    # Tip: inspect a real header row first:
    #   Get-Content (Get-ChildItem $WorkFolder -Filter *.txt | Select -First 1).FullName -Head 1
    $createTableSql = @"
IF OBJECT_ID('$TargetTable', 'U') IS NULL
BEGIN
    CREATE TABLE $TargetTable
    (
        Id              INT           IDENTITY(1,1) PRIMARY KEY,
        Column1         NVARCHAR(255) NULL,
        Column2         NVARCHAR(255) NULL,
        Column3         NVARCHAR(255) NULL,
        Column4         NVARCHAR(255) NULL,
        -- add/remove/rename columns here to match your files --
        SourceFile      NVARCHAR(260) NULL,
        LoadDateTime    DATETIME2     NOT NULL DEFAULT SYSDATETIME()
    );
END
"@
    Invoke-Sql -Query $createTableSql
    Write-Host "  Table $TargetTable ready." -ForegroundColor Green
}

# ------------------------------------------------------------------
# 4. SqlBulkCopy loop
#
#    How it works:
#      - Opens each file as a StreamReader and wraps it in a
#        lightweight IDataReader shim (CsvDataReader class below).
#      - SqlBulkCopy streams the rows from that reader directly over
#        the SQL connection in batches of $BatchSize.
#      - SourceFile and LoadDateTime are added as extra computed
#        columns by the reader shim, so they arrive pre-filled
#        without any server-side staging table.
#      - The file never needs to be visible to the SQL Server host.
# ------------------------------------------------------------------
Write-Host "`n=== Step 4: Loading files via SqlBulkCopy ===" -ForegroundColor Cyan

# --- CsvDataReader: a minimal IDataReader over a pipe-delimited file ----
# Exposes the file's data columns plus two extra columns appended at the end:
#   [n-1] SourceFile    — the bare filename
#   [n]   LoadDateTime  — current UTC datetime (fixed per file)
Add-Type @"
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

public class CsvDataReader : IDataReader
{
    private readonly StreamReader  _reader;
    private readonly char          _delimiter;
    private readonly string        _sourceFile;
    private readonly DateTime      _loadDt;
    private readonly string[]      _headers;       // data columns from file header
    private string[]               _current;
    private bool                   _closed;

    // Total columns = data columns + SourceFile + LoadDateTime
    private int DataColCount  => _headers.Length;
    private int TotalColCount => _headers.Length + 2;

    public CsvDataReader(string filePath, char delimiter)
    {
        _reader     = new StreamReader(filePath, System.Text.Encoding.UTF8, true);
        _delimiter  = delimiter;
        _sourceFile = Path.GetFileName(filePath);
        _loadDt     = DateTime.UtcNow;

        var headerLine = _reader.ReadLine();
        _headers = headerLine == null
            ? Array.Empty<string>()
            : headerLine.Split(_delimiter);
    }

    public int RowCount { get; private set; }

    // IDataReader / IDataRecord -----------------------------------------
    public bool Read()
    {
        if (_closed) return false;
        var line = _reader.ReadLine();
        if (line == null) return false;
        _current = line.Split(_delimiter);
        RowCount++;
        return true;
    }

    public int FieldCount => TotalColCount;

    public object GetValue(int i)
    {
        if (i < DataColCount)
        {
            if (_current == null || i >= _current.Length) return DBNull.Value;
            var v = _current[i];
            return string.IsNullOrEmpty(v) ? (object)DBNull.Value : v;
        }
        if (i == DataColCount)     return _sourceFile;  // SourceFile
        if (i == DataColCount + 1) return _loadDt;      // LoadDateTime
        throw new IndexOutOfRangeException();
    }

    public string GetName(int i)
    {
        if (i < DataColCount)      return _headers[i];
        if (i == DataColCount)     return "SourceFile";
        if (i == DataColCount + 1) return "LoadDateTime";
        throw new IndexOutOfRangeException();
    }

    public int GetOrdinal(string name)
    {
        for (int i = 0; i < _headers.Length; i++)
            if (string.Equals(_headers[i], name, StringComparison.OrdinalIgnoreCase)) return i;
        if (name == "SourceFile")   return DataColCount;
        if (name == "LoadDateTime") return DataColCount + 1;
        throw new IndexOutOfRangeException(name);
    }

    // --- stubs required by IDataReader that we don't need ---
    public void Close()  { _closed = true; _reader.Dispose(); }
    public void Dispose(){ Close(); }
    public int  Depth    => 0;
    public bool IsClosed => _closed;
    public int  RecordsAffected => -1;
    public bool NextResult()    => false;
    public DataTable GetSchemaTable() => null;

    public bool    GetBoolean(int i)  => Convert.ToBoolean(GetValue(i));
    public byte    GetByte(int i)     => Convert.ToByte(GetValue(i));
    public long    GetBytes(int i, long fo, byte[] buf, int bo, int len) => 0;
    public char    GetChar(int i)     => Convert.ToChar(GetValue(i));
    public long    GetChars(int i, long fo, char[] buf, int bo, int len) => 0;
    public Guid    GetGuid(int i)     => Guid.Parse(GetValue(i).ToString());
    public short   GetInt16(int i)    => Convert.ToInt16(GetValue(i));
    public int     GetInt32(int i)    => Convert.ToInt32(GetValue(i));
    public long    GetInt64(int i)    => Convert.ToInt64(GetValue(i));
    public float   GetFloat(int i)    => Convert.ToSingle(GetValue(i));
    public double  GetDouble(int i)   => Convert.ToDouble(GetValue(i));
    public decimal GetDecimal(int i)  => Convert.ToDecimal(GetValue(i));
    public DateTime GetDateTime(int i)=> Convert.ToDateTime(GetValue(i));
    public string  GetString(int i)   => GetValue(i)?.ToString();
    public string  GetDataTypeName(int i) => "nvarchar";
    public Type    GetFieldType(int i)    => typeof(string);
    public int     GetValues(object[] values)
    {
        int n = Math.Min(values.Length, TotalColCount);
        for (int i = 0; i < n; i++) values[i] = GetValue(i);
        return n;
    }
    public bool IsDBNull(int i) => GetValue(i) == DBNull.Value;
    public object this[int i]    => GetValue(i);
    public object this[string n] => GetValue(GetOrdinal(n));
}
"@ -ReferencedAssemblies "System.Data"

# --- Load loop -----------------------------------------------------------
$txtFiles = Get-ChildItem -Path $WorkFolder -Filter "*.txt" -File | Sort-Object Name

foreach ($file in $txtFiles) {
    $csvReader = $null
    $bulkCopy  = $null
    try {
        $csvReader = New-Object CsvDataReader($file.FullName, $Delimiter)

        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy(
            $connectionString,
            [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock
        )
        $bulkCopy.DestinationTableName = $TargetTable
        $bulkCopy.BatchSize            = $BatchSize
        $bulkCopy.BulkCopyTimeout      = 0   # no timeout

        # Map each column by name so order in the file doesn't need to
        # exactly match the physical column order in the table.
        for ($i = 0; $i -lt $csvReader.FieldCount; $i++) {
            $colName = $csvReader.GetName($i)
            $bulkCopy.ColumnMappings.Add($colName, $colName) | Out-Null
        }

        $bulkCopy.WriteToServer($csvReader)

        Write-LogEntry -File $file.Name -Status "SUCCESS" `
            -Detail "Loaded OK" -RowsLoaded $csvReader.RowCount
    }
    catch {
        Write-LogEntry -File $file.Name -Status "FAILED" `
            -Detail $_.Exception.Message
    }
    finally {
        if ($bulkCopy)  { $bulkCopy.Close() }
        if ($csvReader) { $csvReader.Close() }
    }
}

# ------------------------------------------------------------------
# 5. Write log and print summary
# ------------------------------------------------------------------
$logEntries | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8

$successCount = ($logEntries | Where-Object Status -eq "SUCCESS").Count
$failCount    = ($logEntries | Where-Object Status -eq "FAILED").Count
$totalRows    = ($logEntries |
    Where-Object Status -eq "SUCCESS" |
    Measure-Object -Property RowsLoaded -Sum).Sum

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "  Succeeded : $successCount" -ForegroundColor Green
Write-Host "  Failed    : $failCount" -ForegroundColor $(if ($failCount -gt 0) {"Red"} else {"Green"})
Write-Host "  Total rows: $totalRows"
Write-Host "  Log       : $LogFile"

if ($failCount -gt 0) {
    Write-Host "`n  Review $LogFile for details on failed files." -ForegroundColor Yellow
}

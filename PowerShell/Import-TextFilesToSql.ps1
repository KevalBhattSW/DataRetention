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
      4. Loops over every .txt file in -WorkFolder and BULK INSERTs it into the
         target table, skipping the header row on each file.
      5. Logs success/failure per file to a log file and to the console.

.NOTES
    Auth: Windows Authentication (Trusted_Connection).
    Delimiter: pipe '|'. Change $FieldTerminator below if yours differs.
    Row terminator assumed '\n'. Change $RowTerminator if your files use \r\n only
    or some other line ending (mixed line endings across files are common when
    files come from different systems — see the note near $RowTerminator below).
#>

[CmdletBinding()]
param(
    [string]$SqlServer        = "localhost",
    [string]$Database         = "YourDatabase",
    [string]$TargetTable      = "dbo.ImportedData",

    [string]$SourceFolder     = "C:\Import\UnzippedFiles",   # the ~40 already-unzipped .txt files
    [string]$ZipFolder        = "C:\Import\ZippedFiles",     # the ~15 .zip archives
    [string]$WorkFolder       = "C:\Import\Work",            # everything gets staged/unzipped here
    [string]$LogFile          = "C:\Import\import_log.csv",

    [switch]$CreateTable,                                    # pass this switch to (re)create the table first
    [switch]$DropTableIfExists,                               # pass this to DROP the table before creating it

    [string]$FieldTerminator = "|",
    [string]$RowTerminator   = "\n",                         # use "\r\n" if your files are CRLF-only and BULK INSERT misbehaves
    [int]$FirstRow            = 2                             # row 2 = skip one header row
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
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = $TimeoutSeconds
        $cmd.ExecuteNonQuery() | Out-Null
    }
    finally {
        $conn.Close()
    }
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
        Write-Host "Extracting $($zip.Name)..."
        Expand-Archive -Path $zip.FullName -DestinationPath $WorkFolder -Force
        Write-LogEntry -File $zip.Name -Status "SUCCESS" -Detail "Unzipped"
    }
    catch {
        Write-LogEntry -File $zip.Name -Status "FAILED" -Detail "Unzip error: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------
# 2. Stage the already-unzipped .txt files into the work folder too,
#    so step 4 only has to loop over one location.
# ------------------------------------------------------------------
Write-Host "`n=== Step 2: Staging already-unzipped files ===" -ForegroundColor Cyan

$plainFiles = Get-ChildItem -Path $SourceFolder -Filter "*.txt" -File -ErrorAction SilentlyContinue
foreach ($f in $plainFiles) {
    $dest = Join-Path $WorkFolder $f.Name
    if (-not (Test-Path $dest)) {
        Copy-Item -Path $f.FullName -Destination $dest
    }
}

# ------------------------------------------------------------------
# 3. Create the target table (edit columns to match your real data!)
# ------------------------------------------------------------------
if ($CreateTable) {
    Write-Host "`n=== Step 3: Creating target table ===" -ForegroundColor Cyan

    if ($DropTableIfExists) {
        Invoke-Sql -Query "IF OBJECT_ID('$TargetTable', 'U') IS NOT NULL DROP TABLE $TargetTable;"
    }

    # >>> EDIT THIS to match your actual file columns/types <<<
    # Tip: run Get-Content on one file's first line to see your real column names:
    #   Get-Content (Get-ChildItem $WorkFolder -Filter *.txt | Select -First 1).FullName -Head 1
    $createTableSql = @"
IF OBJECT_ID('$TargetTable', 'U') IS NULL
BEGIN
    CREATE TABLE $TargetTable (
        Id              INT IDENTITY(1,1) PRIMARY KEY,
        Column1         NVARCHAR(255)   NULL,
        Column2         NVARCHAR(255)   NULL,
        Column3         NVARCHAR(255)   NULL,
        Column4         NVARCHAR(255)   NULL,
        -- add/remove columns here to match your files, in file column order --
        SourceFile      NVARCHAR(260)   NULL,
        LoadDateTime    DATETIME2       NOT NULL DEFAULT SYSDATETIME()
    );
END
"@
    Invoke-Sql -Query $createTableSql
    Write-Host "Table $TargetTable ready." -ForegroundColor Green
}

# ------------------------------------------------------------------
# 4. Bulk insert every .txt file in the work folder
# ------------------------------------------------------------------
Write-Host "`n=== Step 4: Bulk loading files into $TargetTable ===" -ForegroundColor Cyan

$txtFiles = Get-ChildItem -Path $WorkFolder -Filter "*.txt" -File | Sort-Object Name

foreach ($file in $txtFiles) {

    # Use a staging table per file so we can stamp SourceFile/LoadDateTime
    # without needing those columns physically present in the source text.
    $stagingTable = "##Staging_$([Guid]::NewGuid().ToString('N'))"

    $bulkInsertSql = @"
SELECT * INTO $stagingTable FROM $TargetTable WHERE 1=0;
ALTER TABLE $stagingTable DROP COLUMN Id;
ALTER TABLE $stagingTable DROP COLUMN SourceFile;
ALTER TABLE $stagingTable DROP COLUMN LoadDateTime;

BULK INSERT $stagingTable
FROM '$($file.FullName)'
WITH (
    FIELDTERMINATOR = '$FieldTerminator',
    ROWTERMINATOR   = '$RowTerminator',
    FIRSTROW        = $FirstRow,
    CODEPAGE        = 'ACP',
    TABLOCK
);

INSERT INTO $TargetTable
SELECT s.*, '$($file.Name)', SYSDATETIME()
FROM $stagingTable s;

DECLARE @rc INT = (SELECT COUNT(*) FROM $stagingTable);
DROP TABLE $stagingTable;
SELECT @rc AS RowsLoaded;
"@

    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $bulkInsertSql
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
}

# ------------------------------------------------------------------
# 5. Write log and summary
# ------------------------------------------------------------------
$logEntries | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8

$successCount = ($logEntries | Where-Object Status -eq "SUCCESS").Count
$failCount    = ($logEntries | Where-Object Status -eq "FAILED").Count
$totalRows    = ($logEntries | Where-Object Status -eq "SUCCESS" | Measure-Object -Property RowsLoaded -Sum).Sum

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Succeeded: $successCount" -ForegroundColor Green
Write-Host "Failed:    $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "Total rows loaded: $totalRows"
Write-Host "Log written to: $LogFile"

if ($failCount -gt 0) {
    Write-Host "`nReview $LogFile for details on failed files." -ForegroundColor Yellow
}

# Ensure the target directory exists
$targetDir = "C:\Users\KB9\AppData\Local\Temp\Unstructured\labelling"
if (!(Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir
}

# Function to create Excel files
function Create-ExcelFile {
    param ([string]$strPath,
            [Int]$countFiles)
    $strFilename = "TestFile"
    $strExtension = ".xlsx"
    $strNewPath = "$strPath\Excel"
    if (!(Test-Path $strNewPath)) {
        New-Item -ItemType Directory -Path $strNewPath
    }
    for ($intCounter = 1; $intCounter -le $countFiles; $intCounter++) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        $workbook.SaveAs("$strNewPath\$strFilename$intCounter$strExtension")
        $workbook.Close($true)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excel = $null

        $item = Get-Item "$strNewPath\$strFilename$intCounter$strExtension"
        $item.CreationTime = [datetime]"2022-01-01 12:00"
        $item.LastWriteTime = [datetime]"2022-01-01 12:00"

        Start-Sleep -Milliseconds 100
        $item.LastAccessTime = [datetime]"2023-01-01 12:00"
    }
}

# Function to create Word files
function Create-WordFile {
    param ([string]$strPath,
            [Int]$countFiles)
    $strFilename = "TestFile"
    $strExtension = ".docx"
    $strNewPath = "$strPath\Word"
    if (!(Test-Path $strNewPath)) {
        New-Item -ItemType Directory -Path $strNewPath
    }
    for ($intCounter = 1; $intCounter -le $countFiles; $intCounter++) {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Add()
        $document.SaveAs("$strNewPath\$strFilename$intCounter$strExtension")
        $document.Close()
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        $word = $null

        $item = Get-Item "$strNewPath\$strFilename$intCounter$strExtension"
        $item.CreationTime = [datetime]"2022-01-01 12:00"
        $item.LastWriteTime = [datetime]"2022-01-01 12:00"

        Start-Sleep -Milliseconds 100
        $item.LastAccessTime = [datetime]"2023-01-01 12:00"
    }
}

# Function to create PowerPoint files
function Create-PowerPointFile {
    param ([string]$strPath,
            [Int]$countFiles)
    $strFilename = "TestFile"
    $strExtension = ".pptx"
    $strNewPath = "$strPath\PowerPoint"
    if (!(Test-Path $strNewPath)) {
        New-Item -ItemType Directory -Path $strNewPath
    }

    for ($intCounter = 1; $intCounter -le $countFiles; $intCounter++) {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        
        $presentation = $powerPoint.Presentations.Add(0)
        $presentation.SaveAs("$strNewPath\$strFilename$intCounter$strExtension")
        $presentation.Close()
        $powerPoint.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
        $powerPoint = $null

        $item = Get-Item "$strNewPath\$strFilename$intCounter$strExtension"
        $item.CreationTime = [datetime]"2022-01-01 12:00"
        $item.LastWriteTime = [datetime]"2022-01-01 12:00"

        Start-Sleep -Milliseconds 100
        $item.LastAccessTime = [datetime]"2023-01-01 12:00"
    }
}


# Function to create PowerPoint files
function Create-PowerPointFilePwd {
    param ([string]$strPath,
            [Int]$countFiles)
    $strFilename = "TestFilePwd"
    $strExtension = ".pptx"
    $strNewPath = "$strPath\PowerPoint"
    if (!(Test-Path $strNewPath)) {
        New-Item -ItemType Directory -Path $strNewPath
    }

    for ($intCounter = 1; $intCounter -le $countFiles; $intCounter++) {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        
        $presentation = $powerPoint.Presentations.Add(0)
        $presentation.Password = "Test"
        $presentation.SaveAs("$strNewPath\$strFilename$intCounter$strExtension")
        $presentation.Close()
        $powerPoint.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
        $powerPoint = $null

        $item = Get-Item "$strNewPath\$strFilename$intCounter$strExtension"
        $item.CreationTime = [datetime]"2022-01-01 12:00"
        $item.LastWriteTime = [datetime]"2022-01-01 12:00"

        Start-Sleep -Milliseconds 100
        $item.LastAccessTime = [datetime]"2023-01-01 12:00"
    }
}


# Function to create PowerPoint files
function Create-PowerPointFile2003 {
    param ([string]$strPath,
            [Int]$countFiles)
    $strFilename = "TestFile2003"
    $strExtension = ".ppt"
    $strNewPath = "$strPath\PowerPoint"
    if (!(Test-Path $strNewPath)) {
        New-Item -ItemType Directory -Path $strNewPath
    }

    for ($intCounter = 1; $intCounter -le $countFiles; $intCounter++) {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        
        $presentation = $powerPoint.Presentations.Add(0)
        $presentation.SaveAs("$strNewPath\$strFilename$intCounter$strExtension",1)
        $presentation.Close()
        $powerPoint.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
        $powerPoint = $null

        $item = Get-Item "$strNewPath\$strFilename$intCounter$strExtension"
        $item.CreationTime = [datetime]"2022-01-01 12:00"
        $item.LastWriteTime = [datetime]"2022-01-01 12:00"

        Start-Sleep -Milliseconds 100
        $item.LastAccessTime = [datetime]"2023-01-01 12:00"
    }
}

# Function to create PowerPoint files
function Create-PowerPointFile2003Pwd {
    param ([string]$strPath,
            [Int]$countFiles)
    $strFilename = "TestFile2003Pwd"
    $strExtension = ".ppt"
    $strNewPath = "$strPath\PowerPoint"
    if (!(Test-Path $strNewPath)) {
        New-Item -ItemType Directory -Path $strNewPath
    }

    for ($intCounter = 1; $intCounter -le $countFiles; $intCounter++) {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        
        $presentation = $powerPoint.Presentations.Add(0)
        $presentation.Password = "Test"
        $presentation.SaveAs("$strNewPath\$strFilename$intCounter$strExtension",1)
        $presentation.Close()
        $powerPoint.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
        $powerPoint = $null

        $item = Get-Item "$strNewPath\$strFilename$intCounter$strExtension"
        $item.CreationTime = [datetime]"2022-01-01 12:00"
        $item.LastWriteTime = [datetime]"2022-01-01 12:00"

        Start-Sleep -Milliseconds 100
        $item.LastAccessTime = [datetime]"2023-01-01 12:00"
    }
}

function Create-PDFFile {
    param ([string]$strPath,
            [Int]$countFiles)

    $generatedPdf = "C:\Temp\empty.pdf"
    for ($i=1; $i -le $countFiles; $i++) {
        Copy-Item -Path "C:\Temp\empty.pdf" -Destination "$strPath\TestPdf$i.pdf" -Force
        $item = Get-Item "$strPath\TestPdf$i.pdf"
        $item.CreationTime = [datetime]"2022-01-01 12:00"
        $item.LastWriteTime = [datetime]"2022-01-01 12:00"

        Start-Sleep -Milliseconds 100
        $item.LastAccessTime = [datetime]"2023-01-01 12:00"
    }
}
clear

# Execute all three functions in succession
Create-ExcelFile -strPath $targetDir -countFiles 2
Create-WordFile -strPath $targetDir -countFiles 2
Create-PowerPointFile -strPath $targetDir -countFiles 2
Create-PowerPointFilePwd -strPath $targetDir -countFiles 2
Create-PowerPointFile2003 -strPath $targetDir -countFiles 2
Create-PowerPointFile2003Pwd -strPath $targetDir -countFiles 2
#Create-PDFFile -strPath $targetDir -countFiles 2

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

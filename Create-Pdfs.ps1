clear

for ($i=1; $i -le 10; $i++) {
    $pdfPath = "C:\Temp\Labelling\TestPdf$i.pdf"
<#
    # Create an empty file to print
    $emptyFile = "C:\Temp\empty.txt"
    "" | Out-File -FilePath $emptyFile

    # Set default printer to Microsoft Print to PDF
    if (-not ("PrintHelper" -as [type])) {  
        Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class PrintHelper {
            [DllImport("winspool.drv")]
            public static extern bool SetDefaultPrinter(string Name);}
"@
}
    [PrintHelper]::SetDefaultPrinter("Microsoft Print to PDF")

    # Print the empty file to create a blank PDF
    $shell = New-Object -ComObject Shell.Application
    $shell.NameSpace(0).ParseName($emptyFile).InvokeVerb("Print")

    Start-Sleep -Seconds 5  # Wait for print job to complete
    #>
    # Move and rename the PDF file
    $generatedPdf = "C:\Temp\empty.pdf"
    Copy-Item -Path "C:\Temp\empty.pdf" -Destination "C:\Temp\Labelling\TestPdf$i.pdf" -Force
    }

Write-Output "10 blank PDFs successfully created in C:\Temp\Labelling!"
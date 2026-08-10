$thumbprint = "83D742C9BB001CAE0E70F85D310B64145E2C6407"

$cert = Get-ChildItem Cert:\LocalMachine\My -CodeSigningCert | Where-Object { $_.thumbprint -eq $thumbprint } | Select-Object -First 1

if ($null -eq $cert) {
    Write-Error "Code signing certificate not found - cannot find script"
}

$result = Set-AuthenticodeSignature `
    -FilePath "C:\Temp\Start-Tagging-Detached.ps1" `
    -Certificate $cert

if($result.Status -ne "Valid") {
    Write-Error "Script signing failed : $($result.StatusMessage)"
    exit 1
}

Write-Host "Script signed successfully. Thumbprint: $($cert.Thumbprint)"
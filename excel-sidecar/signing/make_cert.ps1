# Creates the long-lived code-signing certificate for the Automated Testing
# Template VBA project and exports the PUBLIC half next to this script.
# Run ONCE as the template owner. The private key stays in YOUR user store
# (certmgr.msc -> Personal). Never commit a .pfx.
$ErrorActionPreference = "Stop"
$existing = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { $_.Subject -eq "CN=SDR DataViewer Templates" }
if ($existing) {
    Write-Host "Certificate already exists (thumbprint $($existing[0].Thumbprint)); re-exporting public key."
    $cert = $existing[0]
} else {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert `
        -Subject "CN=SDR DataViewer Templates" `
        -CertStoreLocation Cert:\CurrentUser\My `
        -KeyUsage DigitalSignature `
        -NotAfter (Get-Date).AddYears(10)
    Write-Host "Created code-signing certificate, thumbprint $($cert.Thumbprint)"
}
$out = Join-Path $PSScriptRoot "DataViewerTemplates.cer"
Export-Certificate -Cert $cert -FilePath $out | Out-Null
Write-Host "Public certificate exported -> $out"
Write-Host "Commit DataViewerTemplates.cer; testers trust it via tester-setup.ps1."

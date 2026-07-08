# One-time tester setup for the Automated Testing Template: trusts the
# template's code-signing certificate so signed macros run with NO prompts,
# even on copies that arrive with Mark of the Web (email/Teams/downloads).
# No admin rights needed (current-user stores only).
$ErrorActionPreference = "Stop"
$cer = Join-Path $PSScriptRoot "DataViewerTemplates.cer"
if (-not (Test-Path $cer)) {
    Write-Error "DataViewerTemplates.cer not found next to this script."
}
certutil -user -addstore Root $cer | Out-Null
certutil -user -addstore TrustedPublisher $cer | Out-Null
Write-Host ""
Write-Host "Done. The DataViewer testing template's macros are now trusted on this"
Write-Host "account. Close and reopen the template; there should be no macro prompts."

[CmdletBinding()]
param(
    [switch] $Rebuild
)

$ErrorActionPreference = 'Stop'
$testsRoot = $PSScriptRoot

function Find-Qt {
    $candidates = Get-ChildItem 'C:\Qt' -Directory -Filter '6.10.*' -ErrorAction SilentlyContinue |
                  Sort-Object Name -Descending
    foreach ($c in $candidates) {
        $qmake = Join-Path $c.FullName 'mingw_64\bin\qmake.exe'
        if (Test-Path $qmake) { return @{ Qmake = $qmake; Bin = (Split-Path $qmake) } }
    }
    throw 'Could not find a Qt 6.10.x install at C:\Qt\6.10.*\mingw_64\bin\qmake.exe.'
}

function Find-MinGW {
    $candidates = Get-ChildItem 'C:\Qt\Tools' -Directory -Filter 'mingw1*_64' -ErrorAction SilentlyContinue |
                  Sort-Object Name -Descending
    foreach ($c in $candidates) {
        $make = Join-Path $c.FullName 'bin\mingw32-make.exe'
        if (Test-Path $make) { return @{ Make = $make; Bin = (Split-Path $make) } }
    }
    throw 'Could not find MinGW at C:\Qt\Tools\mingw1*_64\bin\mingw32-make.exe.'
}

$qt    = Find-Qt
$mingw = Find-MinGW
Write-Host "Qt    : $($qt.Bin)"
Write-Host "MinGW : $($mingw.Bin)"

$env:PATH = "$($mingw.Bin);$($qt.Bin);$testsRoot\..\vendor\libpq-16;$env:PATH"

Push-Location $testsRoot
try {
    if ($Rebuild -and (Test-Path 'Makefile')) {
        & $mingw.Make distclean | Out-Null
        Get-ChildItem 'Makefile*' -ErrorAction SilentlyContinue | Remove-Item -Force
    }

    if (-not (Test-Path 'Makefile')) {
        Write-Host 'Running qmake...'
        & $qt.Qmake -recursive tests.pro
        if ($LASTEXITCODE -ne 0) { throw "qmake failed (exit $LASTEXITCODE)" }
    }

    Write-Host 'Building tests...'
    $jobs = [Environment]::ProcessorCount
    & $mingw.Make "-j$jobs" release
    if ($LASTEXITCODE -ne 0) { throw "build failed (exit $LASTEXITCODE)" }
}
finally { Pop-Location }

$tstDirs = Get-ChildItem $testsRoot -Directory -Filter 'tst_*' | Sort-Object Name
$pass = 0; $fail = 0; $skip = 0
$failures = @()

Write-Host ''
Write-Host '============================================================'
Write-Host '  Running tests'
Write-Host '============================================================'

foreach ($d in $tstDirs) {
    $exeFile = Join-Path $d.FullName "release\$($d.Name).exe"
    if (-not (Test-Path $exeFile)) {
        Write-Host ("  SKIP  {0,-30} (not built)" -f $d.Name)
        $skip++
        continue
    }

    $logFile = Join-Path $env:TEMP "$($d.Name).txt"
    if (Test-Path $logFile) { Remove-Item $logFile -Force }

    $proc = Start-Process -FilePath $exeFile -ArgumentList @('-o', "`"$logFile,txt`"") `
                          -Wait -PassThru -NoNewWindow

    $totals = $null
    if (Test-Path $logFile) {
        $totals = (Select-String -Path $logFile -Pattern '^Totals:' -SimpleMatch | Select-Object -First 1).Line
    }

    if ($proc.ExitCode -eq 0) {
        $line = if ($totals) { $totals -replace '^Totals: ', '' } else { 'no totals' }
        Write-Host ("  PASS  {0,-30} {1}" -f $d.Name, $line) -ForegroundColor Green
        $pass++
    } else {
        Write-Host ("  FAIL  {0,-30} exit={1} {2}" -f $d.Name, $proc.ExitCode, $totals) -ForegroundColor Red
        $fail++
        $failures += [pscustomobject]@{ Name = $d.Name; Log = $logFile }
    }
}

Write-Host ''
Write-Host '============================================================'
Write-Host ("  Results: {0} passed, {1} failed, {2} skipped" -f $pass, $fail, $skip)
Write-Host '============================================================'

if ($failures.Count -gt 0) {
    Write-Host ''
    Write-Host 'FAILURE DETAILS:'
    foreach ($f in $failures) {
        Write-Host "--- $($f.Name) ---"
        if (Test-Path $f.Log) { Select-String -Path $f.Log -Pattern '^FAIL!' | ForEach-Object { Write-Host $_.Line } }
    }
}

exit ($(if ($fail -gt 0) { 1 } else { 0 }))

<#
.SYNOPSIS
    Deployment self-test for DataViewer Enterprise.

.DESCRIPTION
    Runs three checks on an installed copy of DataViewer Enterprise:

      1. Install tree validation - every DLL, plugin, and bundled resource
         the installer is supposed to drop is present.
      2. In-process diagnostics - launches DataViewer.exe with --self-test,
         which exercises the registry, SQLite driver, bundled Python /
         openpyxl, AppData write permission, and the Synology version-scan.
         Output is read from a JSON report.
      3. Synology folder probe - independently verifies the synced update
         directory exists and lists the version subdirectories Claude
         would see.

    Prints a single PASS/FAIL summary plus a per-check breakdown. Exit
    code 0 iff every check passed.

.PARAMETER InstallDir
    Path to the install directory. If omitted, common per-user and
    machine-wide install locations are searched.

.PARAMETER ReportPath
    Where the binary writes its JSON report. Defaults to
    $env:TEMP\dataviewer_selftest.json.

.EXAMPLE
    .\Test-Deployment.ps1
    .\Test-Deployment.ps1 -InstallDir "C:\Program Files\DataViewer Enterprise"
#>

[CmdletBinding()]
param(
    [string] $InstallDir,
    [string] $ReportPath = (Join-Path $env:TEMP 'dataviewer_selftest.json')
)

$ErrorActionPreference = 'Stop'

# ────────────────────────────────────────────────────────────────────────────
# Resolve install directory
# ────────────────────────────────────────────────────────────────────────────

function Find-InstallDir {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\DataViewer Enterprise'),
        (Join-Path ${env:ProgramFiles} 'DataViewer Enterprise'),
        (Join-Path ${env:ProgramFiles(x86)} 'DataViewer Enterprise')
    )
    foreach ($c in $candidates) {
        if (Test-Path (Join-Path $c 'DataViewer.exe')) { return $c }
    }
    return $null
}

if (-not $InstallDir) {
    $InstallDir = Find-InstallDir
    if (-not $InstallDir) {
        Write-Host 'FAIL: DataViewer.exe not found in standard install locations.' -ForegroundColor Red
        Write-Host '      Pass -InstallDir explicitly.'
        exit 2
    }
}

if (-not (Test-Path (Join-Path $InstallDir 'DataViewer.exe'))) {
    Write-Host "FAIL: DataViewer.exe not found in '$InstallDir'." -ForegroundColor Red
    exit 2
}

Write-Host "Install dir : $InstallDir"
Write-Host "Report path : $ReportPath"
Write-Host ""

# ────────────────────────────────────────────────────────────────────────────
# Phase 1 — install tree validation
# ────────────────────────────────────────────────────────────────────────────

$required = @(
    'DataViewer.exe',
    'Qt6Core.dll',
    'Qt6Gui.dll',
    'Qt6Widgets.dll',
    'Qt6Sql.dll',
    'Qt6Network.dll',
    'libgcc_s_seh-1.dll',
    'libstdc++-6.dll',
    'libwinpthread-1.dll',
    'platforms\qwindows.dll',
    'sqldrivers\qsqlite.dll',
    'python\python.exe',
    'resources\images\ccell_icon.ico',
    'resources\templates\Standardized Test Template - December 2025.xlsx'
)

$missing = @()
foreach ($rel in $required) {
    $p = Join-Path $InstallDir $rel
    if (-not (Test-Path $p)) { $missing += $rel }
}

$treeOK = ($missing.Count -eq 0)
if ($treeOK) {
    Write-Host 'PHASE 1 - install tree                                    : PASS' -ForegroundColor Green
} else {
    Write-Host 'PHASE 1 - install tree                                    : FAIL' -ForegroundColor Red
    foreach ($m in $missing) { Write-Host "   missing: $m" }
}

# ────────────────────────────────────────────────────────────────────────────
# Phase 2 — in-process self-test
# ────────────────────────────────────────────────────────────────────────────

if (Test-Path $ReportPath) { Remove-Item $ReportPath -Force }
$exe = Join-Path $InstallDir 'DataViewer.exe'

$proc = Start-Process -FilePath $exe `
    -ArgumentList @('--self-test', '--self-test-out', "`"$ReportPath`"") `
    -Wait -PassThru -NoNewWindow

$report = $null
if (Test-Path $ReportPath) {
    try { $report = Get-Content -Raw $ReportPath | ConvertFrom-Json } catch { $report = $null }
}

$selfTestOK = ($proc.ExitCode -eq 0) -and ($report -ne $null) -and $report.all_pass
if ($selfTestOK) {
    Write-Host 'PHASE 2 - in-process diagnostics                          : PASS' -ForegroundColor Green
} else {
    Write-Host 'PHASE 2 - in-process diagnostics                          : FAIL' -ForegroundColor Red
    Write-Host "   exit code: $($proc.ExitCode)"
    if ($report -eq $null) {
        Write-Host '   no JSON report produced'
    } else {
        foreach ($t in $report.tests) {
            $tag = if ($t.pass) { 'OK' } else { 'FAIL' }
            $color = if ($t.pass) { 'Gray' } else { 'Yellow' }
            Write-Host ("   [{0,4}] {1,-22} {2}" -f $tag, $t.name, $t.detail) -ForegroundColor $color
        }
    }
}

# ────────────────────────────────────────────────────────────────────────────
# Phase 3 — independent Synology probe
# ────────────────────────────────────────────────────────────────────────────

$synRoot = Join-Path $env:USERPROFILE 'SynologyDrive\SDR\Device Group\Software Release\Current'
$synOK = $false
$synVersions = @()
if (Test-Path $synRoot) {
    $synVersions = Get-ChildItem -Path $synRoot -Directory -ErrorAction SilentlyContinue |
                   ForEach-Object {
                       $name = $_.Name
                       $hasInstaller = Test-Path (Join-Path $_.FullName 'DataViewer-setup.exe')
                       [pscustomobject]@{ Name = $name; HasInstaller = $hasInstaller }
                   }
    $synOK = ($synVersions | Where-Object { $_.HasInstaller }).Count -gt 0
}

if ($synOK) {
    Write-Host 'PHASE 3 - Synology update folder                          : PASS' -ForegroundColor Green
    foreach ($v in $synVersions) {
        $tag = if ($v.HasInstaller) { 'INSTALLER' } else { 'no setup.exe' }
        Write-Host ("   {0,-12} {1}" -f $v.Name, $tag)
    }
} else {
    Write-Host 'PHASE 3 - Synology update folder                          : FAIL' -ForegroundColor Red
    Write-Host "   path: $synRoot"
    if (-not (Test-Path $synRoot)) {
        Write-Host '   reason: folder does not exist (Synology Drive not synced?)'
    } else {
        Write-Host '   reason: no version subdir contains DataViewer-setup.exe'
        foreach ($v in $synVersions) {
            Write-Host ("   {0}" -f $v.Name)
        }
    }
}

# ────────────────────────────────────────────────────────────────────────────
# Summary
# ────────────────────────────────────────────────────────────────────────────

Write-Host ""
$allOK = $treeOK -and $selfTestOK -and $synOK
if ($allOK) {
    Write-Host 'OVERALL: PASS' -ForegroundColor Green
    exit 0
} else {
    Write-Host 'OVERALL: FAIL' -ForegroundColor Red
    Write-Host "Detailed report: $ReportPath"
    exit 1
}

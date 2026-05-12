<#
.SYNOPSIS
  Spin up an ephemeral postgres:16 container for the test suite.

.DESCRIPTION
  Starts a container on port 5433, applies deploy/postgres/init.sql
  (the BEGIN..COMMIT block only — the pg_cron tail is skipped because
  this lightweight container does not load the pg_cron extension), sets
  $env:DVE_TEST_PG_CONN, prepends the vendored libpq DLLs to PATH, and
  prints the teardown command.

  Idempotent: if a "dve-test-pg" container already exists, removes it
  first.

.EXAMPLE
  PS> .\tests\start-test-postgres.ps1
#>

$ErrorActionPreference = "Stop"

$existing = docker ps -aq --filter "name=dve-test-pg"
if ($existing) {
    Write-Host "Removing existing dve-test-pg container..."
    docker rm -f dve-test-pg | Out-Null
}

$repoRoot = Resolve-Path "$PSScriptRoot\.."
$initSql  = Join-Path $repoRoot "deploy\postgres\init.sql"
$libpqDir = Join-Path $repoRoot "vendor\libpq-16"

if (-not (Test-Path $initSql)) {
    throw "init.sql not found at $initSql"
}
if (-not (Test-Path $libpqDir)) {
    throw "vendored libpq not found at $libpqDir"
}

Write-Host "Starting postgres:16 on port 5433..."
docker run -d --name dve-test-pg `
    -p 5433:5432 `
    -e POSTGRES_DB=dve_test `
    -e POSTGRES_USER=test `
    -e POSTGRES_PASSWORD=test `
    postgres:16 | Out-Null

Write-Host "Waiting for ready..."
$maxWait = 30
for ($i = 0; $i -lt $maxWait; $i++) {
    docker exec dve-test-pg pg_isready -U test 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { break }
    Start-Sleep -Seconds 1
}
if ($LASTEXITCODE -ne 0) {
    throw "Postgres did not become ready within $maxWait seconds"
}

Write-Host "Applying schema (BEGIN..COMMIT block from init.sql, skipping pg_cron tail)..."
$sqlAll = Get-Content $initSql -Raw
# Extract from "BEGIN;" up to and including the first "COMMIT;"
$beginIdx = $sqlAll.IndexOf("BEGIN;")
$commitIdx = $sqlAll.IndexOf("COMMIT;", $beginIdx)
if ($beginIdx -lt 0 -or $commitIdx -lt 0) {
    throw "Could not locate BEGIN;..COMMIT; in init.sql"
}
$schemaBlock = $sqlAll.Substring($beginIdx, $commitIdx + 7 - $beginIdx)
$schemaBlock | docker exec -i dve-test-pg psql -U test -d dve_test | Out-Null

# Set environment for the current PowerShell session
$env:DVE_TEST_PG_CONN = "host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"
$env:PATH = "$libpqDir;$env:PATH"

Write-Host ""
Write-Host "Ready. Environment set in this PowerShell session:"
Write-Host "  DVE_TEST_PG_CONN = $($env:DVE_TEST_PG_CONN)"
Write-Host "  libpq prepended to PATH from: $libpqDir"
Write-Host ""
Write-Host "To tear down: docker rm -f dve-test-pg"

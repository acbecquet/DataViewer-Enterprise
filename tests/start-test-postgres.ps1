<#
.SYNOPSIS
  Spin up an ephemeral postgres:16 container for the test suite.

.DESCRIPTION
  Starts a container on port 5433, applies deploy/postgres/init.sql and every
  file in deploy/postgres/migrations/, sets $env:DVE_TEST_PG_CONN, prepends the
  vendored libpq DLLs to PATH, and prints the teardown command.

  Provisioning fails LOUDLY. Every psql invocation runs with ON_ERROR_STOP=1 and
  has its exit status checked, and a failure throws with the offending filename
  and psql's full output. The whole test suite reads its schema through this
  script, so a clean-looking run over a silently wrong schema is the worst
  possible outcome.

  Stock postgres:16 does not load pg_cron, so pg_cron statements are stripped
  from the SQL before it is piped to psql - both init.sql's trailing schedule
  block and the two migrations that schedule cleanup jobs inline. The strip is
  positive-match only (nothing but a pg_cron statement is ever dropped - see the
  atomic trivia group on $CronStatementRe, which is what makes that true), every
  stripped statement's opening line is echoed, and anything sitting in init.sql's
  pg_cron tail that is NOT a pg_cron statement throws. New DDL appended to
  init.sql can therefore never be silently skipped.

  SQL is read as UTF-8 and piped as UTF-8. Under PowerShell's ANSI defaults a
  migration containing a micro sign or a degree sign reached Postgres with those
  characters replaced by '?', with no warning. This file stays pure ASCII so
  PowerShell 5.1's own BOM-less read of it cannot hit the same problem.

  Idempotent: if a "dve-test-pg" container already exists, removes it first.

.EXAMPLE
  PS> .\tests\start-test-postgres.ps1
#>

$ErrorActionPreference = "Stop"

# $OutputEncoding is what PowerShell uses to encode a string piped INTO a native
# command. It defaults to the ANSI code page, which cannot represent 'u' with a
# micro sign or a degree sign and substitutes '?' without a word. Paired with
# -Encoding UTF8 on the Get-Content calls below, this makes a migration's
# non-ASCII literals reach Postgres byte-for-byte. Assigned once here at script
# scope so every psql invocation in the script inherits it, rather than being
# re-set per call. Note this is an ordinary variable, not a per-process setting:
# dot-sourcing this script leaves it changed in the caller's session.
$OutputEncoding = New-Object System.Text.UTF8Encoding $false

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Statements stock postgres:16 cannot run: the throwaway container has no
# pg_cron extension. Matched against a whole statement, skipping any leading
# whitespace and "--" comment lines it carries.
#
# The leading-trivia group is ATOMIC - (?>...) - and that is load-bearing, not
# style. A backtrackable group can give back part of a comment and let the
# keyword alternation match INSIDE it, so a statement whose header block merely
# MENTIONS pg_cron gets stripped:
#
#     -- see SELECT cron.schedule('x') for the NAS job
#     CREATE TABLE metric_values (...);        <- silently dropped
#
# which is exactly the "positive-match only" guarantee this file advertises. The
# house style here is dense leading comment blocks and init.sql's own cron tail
# already carries the line "-- cron.schedule(jobname, ...) is idempotent:", so
# the shape is one comment away at all times. Atomic means the trivia is
# consumed once, in full, and the alternation must then match real SQL.
$CronStatementRe = [regex] '(?i)^(?>(?:\s|--[^\r\n]*)*)(?:CREATE\s+EXTENSION\b[^;]*\bpg_cron\b|SELECT\s+cron\.)'

# Migrations that must never be auto-applied to the test container, by filename,
# each with the reason. An entry here is a decision; a migration that simply
# fails is a bug and must throw.
$MigrationSkips = @{
    '2026-06-25-dv15-rekey-forked-sensory.sql' =
        'one-off DATAVIEWER-15 data repair - its own header says run manually after a backup'
    # TEMP until Phase 3d Task 3 splits the databases: applying the cutover
    # here would rename data_rows/samples out from under every other suite
    # before the app code is cut over.
    '2026-08-26-v3-cutover.sql' =
        'Phase 3d cutover - applied to dve_test only after Task 3 provisions dve_test_precut'
}

# Split SQL into statements on semicolons at depth 0, respecting single-quoted
# literals, dollar-quoted blocks ($$ / $tag$), "--" comments and "/* */" block
# comments. Each piece keeps its leading trivia and its trailing ";", so
# -join '' reproduces the input exactly - that is what lets the pg_cron filter
# drop whole statements without disturbing a single byte of the rest of the file.
function Split-SqlStatements {
    param([Parameter(Mandatory = $true)][string] $Sql)

    $dollarTag = [regex] '\G\$[A-Za-z_0-9]*\$'
    $pieces = New-Object System.Collections.Generic.List[string]
    $start = 0
    $i = 0

    while ($i -lt $Sql.Length) {
        $c = $Sql[$i]

        if ($c -eq "'") {
            $i++
            while ($i -lt $Sql.Length) {
                if ($Sql[$i] -ne "'") { $i++; continue }
                # '' inside a literal is an escaped quote, not the end of it.
                if ($i + 1 -lt $Sql.Length -and $Sql[$i + 1] -eq "'") { $i += 2; continue }
                break
            }
            $i++
        }
        elseif ($c -eq '-' -and $i + 1 -lt $Sql.Length -and $Sql[$i + 1] -eq '-') {
            while ($i -lt $Sql.Length -and $Sql[$i] -ne "`n") { $i++ }
        }
        elseif ($c -eq '/' -and $i + 1 -lt $Sql.Length -and $Sql[$i + 1] -eq '*') {
            # Everything inside a block comment is inert: a ';' must not split
            # the statement and an apostrophe must not open a phantom literal
            # that swallows the rest of the file. Postgres block comments NEST,
            # so track depth instead of stopping at the first "*/".
            $depth = 1
            $i += 2
            while ($i -lt $Sql.Length -and $depth -gt 0) {
                if ($Sql[$i] -eq '/' -and $i + 1 -lt $Sql.Length -and $Sql[$i + 1] -eq '*') {
                    $depth++; $i += 2
                }
                elseif ($Sql[$i] -eq '*' -and $i + 1 -lt $Sql.Length -and $Sql[$i + 1] -eq '/') {
                    $depth--; $i += 2
                }
                else { $i++ }
            }
            if ($depth -gt 0) { throw "Unterminated /* block comment in SQL input" }
        }
        elseif ($c -eq '$') {
            $m = $dollarTag.Match($Sql, $i)
            if ($m.Success) {
                $close = $Sql.IndexOf($m.Value, $i + $m.Value.Length, [StringComparison]::Ordinal)
                if ($close -lt 0) {
                    throw "Unterminated dollar-quoted block '$($m.Value)' in SQL input"
                }
                $i = $close + $m.Value.Length
            }
            else { $i++ }   # a positional parameter such as $1, not a quote tag
        }
        elseif ($c -eq ';') {
            $pieces.Add($Sql.Substring($start, $i - $start + 1))
            $start = $i + 1
            $i++
        }
        else { $i++ }
    }
    if ($start -lt $Sql.Length) { $pieces.Add($Sql.Substring($start)) }

    , $pieces.ToArray()
}

# True when a statement piece is nothing but whitespace and comments, i.e. the
# trailing scrap after a file's last semicolon. Both comment forms Split-Sql-
# Statements knows about are recognised here, so a purely decorative "/* */"
# banner in init.sql's pg_cron tail is not mistaken for dropped DDL.
function Test-SqlTrivia {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string] $Statement)
    # Non-greedy on purpose: it strips each block comment individually rather
    # than everything between the first "/*" and the last "*/", so real DDL
    # sandwiched between two comments still reads as non-trivia and throws.
    $bare = $Statement -replace '(?s)/\*.*?\*/', '' -replace '(?m)--[^\r\n]*', ''
    return ($bare.Trim() -eq '')
}

# The first line of a statement that is neither blank nor a "--" comment - the
# line a reader uses to recognise it. Echoed for every stripped pg_cron
# statement so a mis-strip is visible in the output instead of hiding inside a
# bare count.
function Get-SqlStatementHead {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string] $Statement)
    foreach ($line in ($Statement -split "`n")) {
        $t = $line.Trim()
        if ($t -eq '' -or $t.StartsWith('--')) { continue }
        if ($t.Length -gt 100) { return $t.Substring(0, 100) + '...' }
        return $t
    }
    return '(comment-only statement)'
}

# Pipe SQL into psql inside the container and throw, quoting psql's own output,
# if it exits non-zero. Quiet on success: the "already exists, skipping" NOTICE
# spam is exactly what used to bury the two real ERROR lines this script now
# refuses to walk past.
function Invoke-Psql {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string] $Sql,
        [Parameter(Mandatory = $true)][string] $Label,
        # v3 Phase 3d (index D-3d-8): the same machinery provisions two
        # databases - dve_test (post-cutover, for the app-facing suites) and
        # dve_test_precut (pre-cutover, for the migration-rehearsal harness).
        [string] $Database = 'dve_test'
    )

    # Nothing left to apply. Two inputs land here: a migration whose statements
    # are all pg_cron (or that is nothing but comments), and an init.sql whose
    # very first statement is a pg_cron statement. Without AllowEmptyString the
    # Mandatory parameter rejects those with "cannot bind argument ... empty
    # string", which names neither the file nor the reason.
    if ($Sql.Trim() -eq '') {
        # ${Label}, not $Label: a bare colon after a variable name is parsed as
        # a drive qualifier and the script will not even load.
        Write-Host "  (${Label}: nothing left to apply after filtering)"
        return
    }

    # PowerShell 5.1 turns a native command's stderr into a terminating error
    # when $ErrorActionPreference is 'Stop' and stderr is redirected - and psql
    # writes every NOTICE to stderr. Relax it just around the call so the exit
    # status, not the noise, decides.
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = $Sql | docker exec -i dve-test-pg psql -U test -d $Database -v ON_ERROR_STOP=1 2>&1
        $exit = $LASTEXITCODE
    }
    finally { $ErrorActionPreference = $prevEap }

    if ($exit -ne 0) {
        # Redirected native stderr arrives as ErrorRecord objects; ToString()
        # gives back the raw psql line instead of PowerShell's boxed rendering.
        $lines = @($output | ForEach-Object { $_.ToString() })
        Write-Host ""
        Write-Host "--- psql output for $Label ---" -ForegroundColor Yellow
        $lines | ForEach-Object { Write-Host $_ }
        Write-Host "--- end psql output ---" -ForegroundColor Yellow
        Write-Host ""

        $firstError = $lines | Where-Object { $_ -match '^\s*ERROR:' } | Select-Object -First 1
        if (-not $firstError) { $firstError = 'see the psql output above' }
        throw "Applying $Label failed (psql exit $exit): $firstError"
    }
}

# ---------------------------------------------------------------------------
# Container
# ---------------------------------------------------------------------------

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

# ---------------------------------------------------------------------------
# Schema: init.sql, minus its pg_cron tail
# ---------------------------------------------------------------------------

Write-Host "Applying schema from init.sql (pg_cron tail excluded)..."
$initStatements = Split-SqlStatements (Get-Content $initSql -Raw -Encoding UTF8)

$firstCron = -1
for ($i = 0; $i -lt $initStatements.Count; $i++) {
    if ($CronStatementRe.IsMatch($initStatements[$i])) { $firstCron = $i; break }
}

if ($firstCron -lt 0) {
    # No pg_cron anywhere - the whole file is applicable. Nothing to skip and
    # nothing to police.
    $schemaBlock = $initStatements -join ''
}
else {
    # Everything from the first pg_cron statement to EOF is the documented tail.
    # It may contain pg_cron statements and trailing trivia and nothing else -
    # otherwise real DDL would be dropped on the floor without a word.
    $cronHeads = New-Object System.Collections.Generic.List[string]
    for ($i = $firstCron; $i -lt $initStatements.Count; $i++) {
        $stmt = $initStatements[$i]
        if ($CronStatementRe.IsMatch($stmt)) {
            $cronHeads.Add((Get-SqlStatementHead $stmt))
            continue
        }
        if (Test-SqlTrivia $stmt) { continue }
        $preview = ($stmt.Trim() -replace '\s+', ' ')
        if ($preview.Length -gt 160) { $preview = $preview.Substring(0, 160) + '...' }
        throw @"
init.sql has a statement after its pg_cron tail that is not a pg_cron statement,
so the test container would silently miss it:

    $preview

Move it above the pg_cron block in init.sql, or teach this script how to apply it.
"@
    }
    # Select-Object, not [0..($firstCron - 1)]: the range operator turns 0..-1
    # into "0, -1" and would silently splice in the LAST statement.
    $schemaBlock = ($initStatements | Select-Object -First $firstCron) -join ''
    Write-Host "  ($($cronHeads.Count) pg_cron statement(s) excluded - stock postgres:16 has no pg_cron)"
    $cronHeads | ForEach-Object { Write-Host "      $_" }
}

Invoke-Psql -Sql $schemaBlock -Label "init.sql"

# ---------------------------------------------------------------------------
# Migrations
# ---------------------------------------------------------------------------

# init.sql is the v2.0 BASELINE only; every schema change since then ships as a
# migration file (the canonical record of what ensureSchema() heals at runtime
# AND of manual NAS migrations). Without applying them the test container is
# missing post-baseline objects the app actually uses - most importantly the
# OCC overloads dve_commit_cell(6 args) / dve_commit_cell_json(7 args) that
# LiveSyncWorker dispatches (2026-05-16-livesync-commit-fns + 2026-05-17-v2.0.2-
# fixes). A fresh container without them silently fails every per-cell commit
# test (focusCell passes, commitCell/commitJson don't) - see the v2.4.2 sub-plan
# 1 investigation.
#
# Two migrations schedule pg_cron cleanup jobs (2026-05-16-cell-focus.sql and
# 2026-05-17-v2.0.2-fixes.sql). Those statements are stripped for the same
# reason as init.sql's tail. This is not cosmetic: cell-focus.sql wraps its
# cron.schedule call INSIDE its own BEGIN/COMMIT, so under the old
# ON_ERROR_STOP=0 loop the failing statement poisoned the transaction and the
# trailing COMMIT silently degraded to a ROLLBACK - the entire migration was
# discarded on every provisioning run. It was masked only because init.sql
# happens to define the same objects.
$migDir = Join-Path $repoRoot "deploy\postgres\migrations"
if (Test-Path $migDir) {
    Write-Host "Applying migrations (deploy/postgres/migrations/*.sql in order)..."
    Get-ChildItem -Path $migDir -Filter *.sql | Sort-Object Name | ForEach-Object {
        $name = $_.Name
        if ($MigrationSkips.ContainsKey($name)) {
            Write-Host "  - $name  [SKIPPED: $($MigrationSkips[$name])]"
            return
        }

        $statements = Split-SqlStatements (Get-Content $_.FullName -Raw -Encoding UTF8)
        $keep    = @($statements | Where-Object { -not $CronStatementRe.IsMatch($_) })
        $dropped = @($statements | Where-Object { $CronStatementRe.IsMatch($_) })

        if ($dropped.Count -gt 0) {
            Write-Host "  - $name  ($($dropped.Count) pg_cron statement(s) stripped)"
            $dropped | ForEach-Object { Write-Host "      $(Get-SqlStatementHead $_)" }
        }
        else {
            Write-Host "  - $name"
        }
        Invoke-Psql -Sql ($keep -join '') -Label "migration $name"
    }
}

# ---------------------------------------------------------------------------
# v3 Phase 3d (index D-3d-8): dve_test_precut - the PRE-cutover database
# ---------------------------------------------------------------------------
# dve_test above is post-cutover (once 2026-08-26-v3-cutover.sql applies there):
# data_rows and samples are the name-holder VIEWS, which is what every
# app-facing suite must test against. But tst_v3longformat's migration
# rehearsal seeds wide rows and runs dve_migrate_to_long_format itself, which
# needs the wide tables to still be REAL tables. It gets its own database,
# provisioned by the SAME machinery with ONE difference: the cutover file is
# excluded. The harness cuts its private copy over in its own cutover slots
# and un-cuts it at init for re-runs.

Write-Host ""
Write-Host "Provisioning dve_test_precut (pre-cutover rehearsal database)..."
$prevEapDb = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
try {
    "DROP DATABASE IF EXISTS dve_test_precut;" |
        docker exec -i dve-test-pg psql -U test -d postgres -v ON_ERROR_STOP=1 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "could not drop dve_test_precut (psql exit $LASTEXITCODE)" }
    "CREATE DATABASE dve_test_precut OWNER test;" |
        docker exec -i dve-test-pg psql -U test -d postgres -v ON_ERROR_STOP=1 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "could not create dve_test_precut (psql exit $LASTEXITCODE)" }
}
finally { $ErrorActionPreference = $prevEapDb }

Invoke-Psql -Sql $schemaBlock -Label "init.sql (precut)" -Database 'dve_test_precut'

if (Test-Path $migDir) {
    Get-ChildItem -Path $migDir -Filter *.sql | Sort-Object Name | ForEach-Object {
        $name = $_.Name
        if ($MigrationSkips.ContainsKey($name)) {
            Write-Host "  - $name  [SKIPPED: $($MigrationSkips[$name])]"
            return
        }
        if ($name -like '*v3-cutover*') {
            Write-Host "  - $name  [precut: EXCLUDED - this database stays pre-cutover by design]"
            return
        }
        $statements = Split-SqlStatements (Get-Content $_.FullName -Raw -Encoding UTF8)
        $keep = @($statements | Where-Object { -not $CronStatementRe.IsMatch($_) })
        Write-Host "  - $name  (precut)"
        Invoke-Psql -Sql ($keep -join '') -Label "migration $name (precut)" -Database 'dve_test_precut'
    }
}

# ---------------------------------------------------------------------------
# Session environment
# ---------------------------------------------------------------------------

$env:DVE_TEST_PG_CONN = "host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"
$env:PATH = "$libpqDir;$env:PATH"

Write-Host ""
Write-Host "Ready. Environment set in this PowerShell session:"
Write-Host "  DVE_TEST_PG_CONN = $($env:DVE_TEST_PG_CONN)"
Write-Host "  libpq prepended to PATH from: $libpqDir"
Write-Host ""
Write-Host "To tear down: docker rm -f dve-test-pg"

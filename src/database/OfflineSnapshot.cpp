#include "OfflineSnapshot.h"
#include <QHash>

#include "MetricDefCache.h"
#include "PostgresConnection.h"
#include "RawGridJson.h"
#include "../utils/MipFallback.h"

#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QStandardPaths>
#include <QSqlDatabase>
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlRecord>
#include <QUuid>
#include <atomic>
#include <QVariant>
#include <QDateTime>
#include <QTimeZone>
#include <QDebug>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QRectF>
#include <QStringList>

#include <string>

// R7: crash-safe snapshot promotion uses the Win32 atomic replace
// (MoveFileExW with MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) so a
// crash mid-promotion can never leave the user with NO snapshot. This is a
// Windows-only project (see CLAUDE.md); the guard documents the assumption
// without pretending to be portable.
#ifdef Q_OS_WIN
#  ifndef WIN32_LEAN_AND_MEAN
#    define WIN32_LEAN_AND_MEAN
#  endif
#  ifndef NOMINMAX
#    define NOMINMAX
#  endif
#  include <windows.h>
#endif

namespace DVE {

// Snapshot SQLite schema version, written into _snapshot_meta on regenerate.
// Bumped whenever the snapshot table SHAPE changes so a newer app refuses to
// trust an older snapshot's layout.
//   v2: baseline shipped layout.
//   v3 (v2.4.4 R7b): app_version column added to files / sensory_sessions /
//       detailed_sensory_sessions.
//   v4 (v3 Phase 3c): metric_defs / measurements / sample_headers mirrored, so
//       open metrics (DataRow::extra / SampleResult::extra) survive offline.
// The read-side validation that rejects a mismatching version lands in
// SP3-T2 (R7); this task only writes the new value so that change is recorded.
//
// A BUMP INVALIDATES EVERY SNAPSHOT ALREADY ON A USER'S MACHINE. That is safe
// by design and needs no migration: openReadOnly() rejects the mismatching file
// and the caller treats a reject exactly like "no snapshot yet", regenToPath()
// refuses to reuse it incrementally, and the next clean online close rebuilds
// it in full. The fallback path IS a regen. The only cost is one full rebuild
// per user instead of an incremental one, and one offline session without a
// local cache if the NAS goes down before that close.
static constexpr int kSnapshotSchemaVersion = 4;

// R7: how stale a snapshot may get before openReadOnly() logs a LOUD warning.
// Staleness alone never rejects the snapshot -- a stale-but-readable offline
// cache is strictly better than none when the NAS is unreachable -- it only
// surfaces a warning the offline banner can echo. A schema-version MISMATCH,
// by contrast, IS a hard reject (the layout can't be trusted).
static constexpr int kSnapshotStaleWarnDays = 30;

// ----------------------------------------------------------------------------
// SQLite schema (mirrors Postgres -- minus presence + schema_meta).
// ----------------------------------------------------------------------------
//   * BIGSERIAL -> INTEGER PRIMARY KEY (rowid alias in SQLite).
//   * TIMESTAMPTZ -> TEXT (ISO-8601 round-trip).
//   * JSONB -> TEXT (we don't query JSON paths offline; the read accessors
//     simply round-trip the textual JSON back into the in-memory structs).
//   * BYTEA -> BLOB.
//
// Column NAMES match Postgres so SELECT/INSERT can copy schemas trivially.
// Triggers and FK CASCADE are intentionally omitted -- the snapshot is
// read-only at runtime; child rows are deleted as a unit when the file is
// regenerated, not via FK cascades.
// SP4.5 Stage 2b: regen progress phase budget (prepare + 13 tables + meta +
// finalize). v3 Phase 3c added metric_defs / measurements / sample_headers.
static constexpr int kRegenPhases = 16;

// v3 Phase 3c: the three long-format tables the snapshot mirrors as of schema
// v4. Named once so the copy blocks, the fingerprint's absent-table fallback,
// and the incremental reset all agree on the list.
static const char* const kLongFormatTables[] = {
    "metric_defs", "measurements", "sample_headers"
};

static bool isLongFormatTable(const char* table) {
    for (const char* t : kLongFormatTables)
        if (qstrcmp(t, table) == 0) return true;
    return false;
}

static const char* const kCreateStatements[] = {
    R"(CREATE TABLE files (
        id               INTEGER PRIMARY KEY,
        file_path        TEXT NOT NULL,
        file_name        TEXT NOT NULL,
        loaded_at        TEXT NOT NULL,
        template_version TEXT,
        sheet_count      INTEGER DEFAULT 0,
        sample_count     INTEGER DEFAULT 0,
        added_at         TEXT,
        updated_at       TEXT NOT NULL,
        updated_by       TEXT NOT NULL,
        version          INTEGER NOT NULL DEFAULT 1,
        app_version      TEXT
    ))",
    // F6: a path may now have several versioned rows (file_path, added_at).
    // Mirror Postgres' composite uniqueness so regenerate() can copy every
    // version without colliding; a single-column UNIQUE(file_path) would reject
    // all but the first version.
    R"(CREATE UNIQUE INDEX idx_files_path ON files(file_path, added_at))",

    R"(CREATE TABLE tests (
        id                 INTEGER PRIMARY KEY,
        file_id            INTEGER NOT NULL,
        sheet_name         TEXT NOT NULL,
        template_version   TEXT,
        overall_avg_tpm    REAL DEFAULT 0.0,
        overall_stddev_tpm REAL DEFAULT 0.0,
        is_raw_table       INTEGER DEFAULT 0,
        from_inferred_schema INTEGER NOT NULL DEFAULT 0,
        sort_order         INTEGER DEFAULT 0,
        updated_at         TEXT NOT NULL,
        updated_by         TEXT NOT NULL,
        version            INTEGER NOT NULL DEFAULT 1,
        raw_grid           TEXT
    ))",
    R"(CREATE INDEX idx_tests_file ON tests(file_id))",

    R"(CREATE TABLE samples (
        id                  INTEGER PRIMARY KEY,
        test_id             INTEGER NOT NULL,
        sort_order          INTEGER DEFAULT 0,
        sample_name         TEXT,
        sample_id           TEXT,
        date                TEXT,
        tester              TEXT,
        media               TEXT,
        viscosity           REAL DEFAULT 0.0,
        resistance          REAL DEFAULT 0.0,
        voltage             REAL DEFAULT 0.0,
        power               REAL DEFAULT 0.0,
        heating_technology  TEXT,
        puffing_regime      TEXT,
        initial_oil_mass    REAL DEFAULT 0.0,
        average_tpm         REAL DEFAULT 0.0,
        stddev_tpm          REAL DEFAULT 0.0,
        avg_power_density   REAL DEFAULT 0.0,
        efficiency_percent  REAL DEFAULT 0.0,
        total_oil_consumed  REAL DEFAULT 0.0,
        total_puffs         INTEGER DEFAULT 0,
        normalized_tpm      REAL DEFAULT 0.0,
        burn_status         TEXT,
        clog_status         TEXT,
        leak_status         TEXT,
        updated_at          TEXT NOT NULL,
        updated_by          TEXT NOT NULL,
        version             INTEGER NOT NULL DEFAULT 1
    ))",
    R"(CREATE INDEX idx_samples_test ON samples(test_id))",

    R"(CREATE TABLE data_rows (
        id                INTEGER PRIMARY KEY,
        sample_id         INTEGER NOT NULL,
        sort_order        INTEGER DEFAULT 0,
        puffs             REAL DEFAULT 0.0,
        before_weight     REAL DEFAULT 0.0,
        after_weight      REAL DEFAULT 0.0,
        draw_pressure     REAL DEFAULT 0.0,
        resistance        REAL DEFAULT 0.0,
        smell             TEXT,
        clog              TEXT,
        notes             TEXT,
        tpm               REAL DEFAULT 0.0,
        tpm_power_density REAL DEFAULT 0.0,
        variation_tpm     REAL DEFAULT 0.0,
        oil_consumed      REAL DEFAULT 0.0,
        puffing_regime    TEXT,
        updated_at        TEXT NOT NULL,
        updated_by        TEXT NOT NULL,
        version           INTEGER NOT NULL DEFAULT 1
    ))",
    R"(CREATE INDEX idx_data_rows_sample ON data_rows(sample_id))",

    R"(CREATE TABLE images (
        id          INTEGER PRIMARY KEY,
        sample_id   INTEGER NOT NULL,
        sort_order  INTEGER DEFAULT 0,
        file_name   TEXT,
        image_data  BLOB,
        layout_x    REAL,
        layout_y    REAL,
        layout_w    REAL,
        layout_h    REAL,
        crop_x      REAL,
        crop_y      REAL,
        crop_w      REAL,
        crop_h      REAL,
        updated_at  TEXT NOT NULL,
        updated_by  TEXT NOT NULL,
        version     INTEGER NOT NULL DEFAULT 1
    ))",
    R"(CREATE INDEX idx_images_sample ON images(sample_id))",

    R"(CREATE TABLE sensory_sessions (
        id            INTEGER PRIMARY KEY,
        session_name  TEXT,
        tester_name   TEXT,
        assessor_name TEXT,
        media         TEXT,
        puff_length   TEXT,
        date          TEXT,
        timestamp     TEXT,
        json_data     TEXT,
        layout_json   TEXT,
        updated_at    TEXT NOT NULL,
        updated_by    TEXT NOT NULL,
        version       INTEGER NOT NULL DEFAULT 1,
        app_version   TEXT
    ))",

    R"(CREATE TABLE sensory_images (
        id          INTEGER PRIMARY KEY,
        session_id  INTEGER NOT NULL,
        sort_order  INTEGER DEFAULT 0,
        file_name   TEXT,
        image_data  BLOB,
        layout_x    REAL,
        layout_y    REAL,
        layout_w    REAL,
        layout_h    REAL,
        crop_x      REAL,
        crop_y      REAL,
        crop_w      REAL,
        crop_h      REAL,
        updated_at  TEXT NOT NULL,
        updated_by  TEXT NOT NULL,
        version     INTEGER NOT NULL DEFAULT 1
    ))",
    R"(CREATE INDEX idx_sensory_images_session ON sensory_images(session_id))",

    R"(CREATE TABLE detailed_sensory_sessions (
        id            INTEGER PRIMARY KEY,
        session_name  TEXT,
        tester_name   TEXT,
        assessor_name TEXT,
        media         TEXT,
        date          TEXT,
        timestamp     TEXT,
        json_data     TEXT,
        updated_at    TEXT NOT NULL,
        updated_by    TEXT NOT NULL,
        version       INTEGER NOT NULL DEFAULT 1,
        app_version   TEXT
    ))",

    R"(CREATE TABLE detailed_sensory_images (
        id          INTEGER PRIMARY KEY,
        session_id  INTEGER NOT NULL,
        sort_order  INTEGER DEFAULT 0,
        file_name   TEXT,
        image_data  BLOB,
        layout_x    REAL,
        layout_y    REAL,
        layout_w    REAL,
        layout_h    REAL,
        crop_x      REAL,
        crop_y      REAL,
        crop_w      REAL,
        crop_h      REAL,
        updated_at  TEXT NOT NULL,
        updated_by  TEXT NOT NULL,
        version     INTEGER NOT NULL DEFAULT 1
    ))",
    R"(CREATE INDEX idx_detailed_sensory_images_session
        ON detailed_sensory_images(session_id))",

    R"(CREATE TABLE settings (
        key        TEXT PRIMARY KEY,
        value      TEXT,
        updated_at TEXT NOT NULL,
        updated_by TEXT NOT NULL,
        version    INTEGER NOT NULL DEFAULT 1
    ))",

    // ── v3 Phase 3c (snapshot schema v4): the long-format tables ────────────
    // Mirrors of deploy/postgres/migrations/2026-07-31-v3-long-format.sql.
    // Column names match Postgres exactly, same as every table above, so the
    // regen copy blocks stay a straight SELECT -> INSERT with no renaming.
    //
    // The UNIQUE constraints are carried over rather than dropped: they are the
    // identity of a measurement / header, and (as in Postgres) their btree
    // leads with sample_id, so the per-sample read in loadFile() is served off
    // the leftmost prefix and no separate FK index is needed. Nothing else
    // about the Postgres definitions comes across -- FK CASCADE and the
    // bump_version trigger are omitted here exactly as they are for every other
    // table (the snapshot is read-only at runtime).
    R"(CREATE TABLE metric_defs (
        id           INTEGER PRIMARY KEY,
        kind         TEXT NOT NULL,
        key          TEXT NOT NULL,
        display_name TEXT NOT NULL,
        value_type   TEXT NOT NULL,
        unit         TEXT,
        role         TEXT,
        tags         TEXT,
        updated_at   TEXT NOT NULL,
        updated_by   TEXT NOT NULL,
        version      INTEGER NOT NULL DEFAULT 1,
        UNIQUE (kind, key)
    ))",

    R"(CREATE TABLE measurements (
        id         INTEGER PRIMARY KEY,
        sample_id  INTEGER NOT NULL,
        metric_id  INTEGER NOT NULL,
        sort_order INTEGER NOT NULL,
        value_num  REAL,
        value_text TEXT,
        updated_at TEXT NOT NULL,
        updated_by TEXT NOT NULL,
        version    INTEGER NOT NULL DEFAULT 1,
        UNIQUE (sample_id, metric_id, sort_order)
    ))",

    R"(CREATE TABLE sample_headers (
        id         INTEGER PRIMARY KEY,
        sample_id  INTEGER NOT NULL,
        field_id   INTEGER NOT NULL,
        value_num  REAL,
        value_text TEXT,
        updated_at TEXT NOT NULL,
        updated_by TEXT NOT NULL,
        version    INTEGER NOT NULL DEFAULT 1,
        UNIQUE (sample_id, field_id)
    ))",

    R"(CREATE TABLE _snapshot_meta (
        key   TEXT PRIMARY KEY,
        value TEXT NOT NULL
    ))",
};

// Forward declarations of file-static helpers used by accessors below.
namespace {
void loadImagesForSnapshot(QSqlDatabase& db,
                           const QString& tableName,
                           qint64 parentId,
                           const QString& parentCol,
                           const QString& imageCacheDir,
                           QStringList* outPaths,
                           QVector<QRectF>* outLayouts,
                           QVector<QRectF>* outCrops,
                           QString* outError);

bool deserializeSensoryJsonLocal(const QByteArray& bytes, SensorySession& sess);
bool deserializeDetailedSensoryJsonLocal(const QByteArray& bytes,
                                         DetailedSensorySession& sess);

} // anonymous

// H14: guard for the hand-maintained SELECT / INSERT / bind-loop triples in
// regenToPath(). The three column lists must stay in lock-step: the SELECT
// result must have exactly as many columns as the INSERT has placeholders, and
// the bind loop must iterate over exactly that many. When they drift (e.g. a
// new column appended to one but not the others) the snapshot silently
// mis-copies data -- and the offline read path is the one place a user sees
// data with no server to cross-check against.
//
// This used to be a Q_ASSERT_X, which compiled out of exactly the release build
// that ships. It is now a RUNTIME check, active in release: on a mismatch it
// logs at critical severity naming the table and all three counts, fills
// *outError, and returns false so the caller can abort the regeneration.
// The query is taken by const reference: QSqlQuery::record() is const, so the
// guard demonstrably does not disturb the caller's cursor.
bool OfflineSnapshot::checkColumnArity(const QSqlQuery& sel, int insertPlaceholders,
                                       int loopBound, const char* table,
                                       QString* outError) {
    const int selCols = sel.record().count();
    if (selCols == insertPlaceholders && insertPlaceholders == loopBound)
        return true;
    const QString msg =
        QStringLiteral("OfflineSnapshot::regenToPath: column arity mismatch for "
                       "table '%1' -- SELECT columns=%2, INSERT placeholders=%3, "
                       "bind-loop bound=%4. Aborting the regeneration; the "
                       "previous snapshot is left in place.")
            .arg(QLatin1String(table))
            .arg(selCols).arg(insertPlaceholders).arg(loopBound);
    qCritical().noquote() << msg;
    if (outError) *outError = msg;
    return false;
}

// ----------------------------------------------------------------------------
// Construction / destruction
// ----------------------------------------------------------------------------
OfflineSnapshot::OfflineSnapshot(QObject* parent) : QObject(parent) {
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    m_connName      = QStringLiteral("dve_snapshot_ro_") + tag;
    m_queueConnName = QStringLiteral("dve_snapshot_queue_") + tag;
}

OfflineSnapshot::~OfflineSnapshot() {
    close();
}

void OfflineSnapshot::setOverrideDirForTesting(const QString& dir) {
    m_overrideDir = dir;
    m_path.clear();
    // The queue file path is derived from the snapshot path; if the queue
    // is already open against the old location, close it so the next
    // ensureQueueOpen() picks up the new override dir.
    if (m_queueOpen) {
        m_queueDb.close();
        m_queueDb = QSqlDatabase();
        if (QSqlDatabase::contains(m_queueConnName)) {
            QSqlDatabase::removeDatabase(m_queueConnName);
        }
        m_queueOpen = false;
    }
}

QString OfflineSnapshot::path() const {
    if (!m_path.isEmpty()) return m_path;
    QString baseDir = m_overrideDir;
    if (baseDir.isEmpty()) {
        baseDir = QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation);
    }
    QDir().mkpath(baseDir);
    const_cast<OfflineSnapshot*>(this)->m_path = baseDir + "/snapshot.sqlite";
    return m_path;
}

// ----------------------------------------------------------------------------
// regenerate -- atomic write to .tmp, rename over production.
//
// PG read consistency: all SELECTs against `live` are issued inside a
// REPEATABLE READ READ ONLY transaction so the snapshot is a consistent
// point-in-time view of Postgres -- another client committing between
// SELECTs cannot leave the snapshot with orphan child rows.
// ----------------------------------------------------------------------------
// SP4.5: a cheap "has the live DB changed?" fingerprint — per-table row COUNT
// (catches inserts/deletes) plus MAX(updated_at) as an epoch (catches in-place
// UPDATEs, since the bump_version trigger stamps updated_at on every UPDATE,
// including dve_commit_cell_json). It is computed IDENTICALLY at regenerate time
// (persisted into _snapshot_meta) and again at close time against the live DB; an
// exact match means the snapshot is still current and the expensive full
// regenerate() (which copies every table + image blob) can be skipped. Returns an
// empty string on any error — the caller treats empty as "unknown" and therefore
// regenerates (the safe default). Every fingerprinted table carries updated_at.
//
// The SQL is BUILT from OfflineSnapshot::kFingerprintTables rather than written
// out, so the statement, the segment count, and the segment order cannot drift
// apart (hazard H10 -- that number used to be spelled out in five coupled
// places). Adding a table is one line in that array.
//
// `outLongTables` reports whether the v3 long-format tables were found. A
// Postgres that predates them - or one where ensureSchema's best-effort DDL did
// not land - must still produce a fingerprint AND still regenerate, so their
// segments are emitted as the literal '0/0' instead of a count that would error
// out. This matters more here than anywhere else: regenToPath computes the
// fingerprint INSIDE its REPEATABLE READ transaction, and one failed statement
// poisons a Postgres transaction, which would turn "this server has no long
// tables" into "this user can never refresh their offline snapshot again".
static QString snapshotContentFingerprint(QSqlDatabase& db,
                                          bool* outLongTables = nullptr)
{
    bool longTables = false;
    {
        // to_regclass() yields NULL for an absent relation instead of raising,
        // so this probe is safe to run inside an open transaction.
        QSqlQuery probe(db);
        if (probe.exec(QStringLiteral(
                "SELECT to_regclass('public.metric_defs')    IS NOT NULL "
                "   AND to_regclass('public.measurements')   IS NOT NULL "
                "   AND to_regclass('public.sample_headers') IS NOT NULL"))
            && probe.next())
            longTables = probe.value(0).toBool();
    }
    if (outLongTables) *outLongTables = longTables;

    QString sql = QStringLiteral("SELECT ");
    for (int i = 0; i < OfflineSnapshot::kFingerprintTableCount; ++i) {
        const QLatin1String t(OfflineSnapshot::kFingerprintTables[i]);
        if (i) sql += QStringLiteral("||';'||");
        if (!longTables && isLongFormatTable(OfflineSnapshot::kFingerprintTables[i])) {
            sql += QStringLiteral("'0/0'");
            continue;
        }
        sql += QStringLiteral(
                   "(SELECT count(*) FROM %1)||'/'||"
                   "coalesce((SELECT extract(epoch FROM max(updated_at))::bigint "
                   "FROM %1),0)").arg(t);
    }

    QSqlQuery q(db);
    if (q.exec(sql) && q.next()) {
        const QString fp = q.value(0).toString();
        return fp.isEmpty() ? QStringLiteral("0") : fp;   // never "" on success
    }
    return QString();
}

QString OfflineSnapshot::liveContentFingerprint(PostgresConnection* live)
{
    if (!live || !live->isOpen()) return QString();
    QSqlDatabase db = live->queryDb();
    return snapshotContentFingerprint(db);
}

QString OfflineSnapshot::storedContentFingerprint() const
{
    if (!m_open) return QString();
    QSqlQuery q(m_db);
    q.prepare("SELECT value FROM _snapshot_meta WHERE key = ?");
    q.addBindValue(QStringLiteral("content_fingerprint"));
    if (q.exec() && q.next()) return q.value(0).toString();
    return QString();
}

bool OfflineSnapshot::isCurrentVsLive(PostgresConnection* live) const
{
    const QString stored = storedContentFingerprint();
    if (stored.isEmpty()) return false;                       // no/old snapshot → regen
    const QString liveFp = liveContentFingerprint(live);
    if (liveFp.isEmpty()) return false;                       // couldn't read live → regen
    return stored == liveFp;
}

// SP4.5 Stage 2b: per-table fingerprint segment helpers. The table ORDER and
// the segment COUNT both come from OfflineSnapshot::kFingerprintTables (see the
// hazard-H10 note on its declaration) -- neither is restated here.
QStringList OfflineSnapshot::fingerprintSegments(const QString& fp) {
    const QStringList parts = fp.split(';');
    return parts.size() == kFingerprintTableCount ? parts : QStringList{};
}

bool OfflineSnapshot::segmentChanged(const QString& priorFp, const QString& liveFp,
                                     const char* table) {
    const QStringList a = fingerprintSegments(priorFp);
    const QStringList b = fingerprintSegments(liveFp);
    if (a.isEmpty() || b.isEmpty()) return true;              // unparseable → refresh
    int idx = -1;
    for (int i = 0; i < kFingerprintTableCount; ++i)
        if (qstrcmp(kFingerprintTables[i], table) == 0) { idx = i; break; }
    if (idx < 0) return true;                                 // unknown table → refresh
    return a[idx] != b[idx];
}

bool OfflineSnapshot::regenToPath(PostgresConnection*  live,
                                  const QString&       destPath,
                                  std::atomic<bool>*   cancel,
                                  QString*             outFingerprint,
                                  QDateTime*           outServerTimeUtc,
                                  QString*             outError,
                                  const RegenProgress& progress,
                                  RegenStats*          outStats)
{
    // SP4.5 Stage 2a: thread-agnostic body of the snapshot regen, extracted
    // VERBATIM from OfflineSnapshot::regenerate (v2.4.5). Runs on whatever
    // thread owns `live`. The error member m_lastError is redirected to
    // *outError via a reference (body unchanged); path() became destPath; the
    // instance-only `if (m_open) close()` is dropped; the server-clock stamp
    // and content fingerprint are returned via out-params; `cancel` (polled by
    // cancelled() between blob tables) lets a close abort a long regen. The
    // REPEATABLE READ READ ONLY transaction is preserved exactly.
    QString _discard;
    QString& m_lastError = outError ? *outError : _discard;
    m_lastError.clear();
    if (!live || !live->isOpen()) {
        m_lastError = QStringLiteral("regenerate: PostgresConnection is not open");
        return false;
    }

    if (outStats) outStats->wasIncremental = false;
    int _phase = 0;
    auto step = [&](const QString& label) {
        // Intermediate ticks cap at kRegenPhases-1; the success path reports a
        // final kRegenPhases/kRegenPhases so the bar always lands at 100%.
        if (progress) progress(qMin(++_phase, kRegenPhases - 1), kRegenPhases, label);
    };

    const QString prodPath = destPath;
    // SP4.5 audit fix: per-call-unique tmp name. A FIXED "<snapshot>.tmp" let a
    // manual "Refresh Snapshot" and the background regen worker create/rename the
    // SAME temp file concurrently -> corrupted snapshot / failed atomic replace.
    // A unique suffix (like tmpConn) lets concurrent regens each use their own tmp
    // and promote atomically (last writer wins; both are valid snapshots).
    const QString uniq     = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    const QString tmpPath  = prodPath + "." + uniq + ".tmp";
    const QString tmpConn  = QStringLiteral("dve_snapshot_tmp_") + uniq;

    // ---- SP4.5 Stage 2b: decide incremental vs full ----------------------
    // Incremental requires a prior prod snapshot that (a) exists, (b) opens,
    // (c) matches the current schema version, (d) has a stored fingerprint to
    // diff against. Any miss -> full rebuild (the safe default). On the
    // incremental path we copy the prior snapshot (reusing its image blobs) and
    // refresh only what changed; otherwise we build from an empty schema.
    QString priorFp;
    bool incremental = false;
    if (QFile::exists(prodPath)) {
        const QString roConn = QStringLiteral("dve_snap_ro_") +
            QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
        {
            QSqlDatabase ro = QSqlDatabase::addDatabase("QSQLITE", roConn);
            ro.setDatabaseName(prodPath);
            ro.setConnectOptions("QSQLITE_OPEN_READONLY");
            if (ro.open()) {
                int ver = -1;
                QSqlQuery vq(ro);
                if (vq.exec("SELECT value FROM _snapshot_meta WHERE key='source_schema_version'")
                    && vq.next())
                    ver = vq.value(0).toInt();
                QSqlQuery fq(ro);
                if (ver == kSnapshotSchemaVersion
                    && fq.exec("SELECT value FROM _snapshot_meta WHERE key='content_fingerprint'")
                    && fq.next())
                    priorFp = fq.value(0).toString();
                ro.close();
            }
        }
        QSqlDatabase::removeDatabase(roConn);
        incremental = !priorFp.isEmpty();
    }
    if (outStats) outStats->wasIncremental = incremental;

    // Wipe any stale .tmp / sidecars from a previous failed/killed regen. The tmp
    // name is now per-call unique, so sweep every "<snapshot>.*.tmp*" orphan (a
    // crash mid-regen would otherwise leak one). Ours doesn't exist yet (created
    // below). A *concurrent* regen's live tmp could match the glob, but removing it
    // is harmless: on Windows SQLite holds the open file without FILE_SHARE_DELETE,
    // so QFile::remove simply fails, and each regen promotes atomically regardless.
    {
        const QFileInfo  prodInfo(prodPath);
        const QDir       snapDir = prodInfo.absoluteDir();
        const QString    base    = prodInfo.fileName();
        const QStringList orphans = snapDir.entryList(
            QStringList{ base + QStringLiteral(".*.tmp"),
                         base + QStringLiteral(".*.tmp-wal"),
                         base + QStringLiteral(".*.tmp-shm") },
            QDir::Files);
        for (const QString& f : orphans)
            QFile::remove(snapDir.absoluteFilePath(f));
    }

    // SP4.5 Stage 2b: incremental base. Copy the prior snapshot wholesale so its
    // image blobs ride along (never re-read from PG); the refresh below patches
    // only the changed tables. A copy failure safely degrades to a full rebuild.
    if (incremental) {
        if (!QFile::copy(prodPath, tmpPath)) {
            incremental = false;
            if (outStats) outStats->wasIncremental = false;
        }
    }

    // The cleanup lambda only handles on-disk artifacts. The named SQLite
    // connection is dropped explicitly on each exit path AFTER the local
    // tmpDb handle is cleared -- otherwise QSqlDatabase emits a
    // "connection still in use" warning (same anti-pattern that
    // PostgresConnection::close() goes out of its way to avoid).
    bool success = false;
    auto cleanup = [&]() {
        if (!success) {
            QFile::remove(tmpPath);
            QFile::remove(tmpPath + "-wal");
            QFile::remove(tmpPath + "-shm");
        }
    };

    QSqlDatabase& pg = live->queryDb();
    bool pgTxnActive = false;
    // Helper to issue a ROLLBACK on PG before bailing out. Safe to call
    // multiple times -- only fires the first time.
    auto rollbackPg = [&]() {
        if (pgTxnActive) {
            QSqlQuery r(pg);
            r.exec("ROLLBACK");
            pgTxnActive = false;
        }
    };

    {
        QSqlDatabase tmpDb = QSqlDatabase::addDatabase("QSQLITE", tmpConn);
        tmpDb.setDatabaseName(tmpPath);
        if (!tmpDb.open()) {
            m_lastError = QStringLiteral("regenerate: failed to open tmp SQLite: ")
                          + tmpDb.lastError().text();
            tmpDb = QSqlDatabase();
            QSqlDatabase::removeDatabase(tmpConn);
            cleanup();
            return false;
        }

        // SQLite write tuning: WAL while we bulk-load (fast), then FULL sync
        // so the tmp's contents are durably on disk BEFORE the atomic promotion
        // (R7). synchronous=NORMAL could leave recently-written pages only in
        // the OS cache; if the machine crashed between the rename and the OS
        // flush, the promoted snapshot could be torn. FULL costs a little on
        // regenerate but the snapshot is the crash-recovery store -- durability
        // wins. Before the move we additionally wal_checkpoint(TRUNCATE) +
        // close so the tmp is a single self-contained file (no -wal/-shm
        // siblings to carry over -- see the promotion block below).
        {
            QSqlQuery pragma(tmpDb);
            pragma.exec("PRAGMA journal_mode=WAL");
            pragma.exec("PRAGMA synchronous=FULL");
        }

        step(QStringLiteral("Preparing snapshot"));

        // Incremental reuses the copied schema + blobs; only a full rebuild
        // creates the schema in the empty tmp.
        if (!incremental)
        for (const char* stmt : kCreateStatements) {
            QSqlQuery q(tmpDb);
            if (!q.exec(stmt)) {
                m_lastError = QStringLiteral("regenerate(CREATE): ")
                              + q.lastError().text();
                tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn);
                cleanup();
                return false;
            }
        }

        // Open the PG-side consistency transaction BEFORE we begin the
        // SQLite-side transaction so any failure to start it doesn't leave
        // a SQLite txn open.
        {
            QSqlQuery pgTxn(pg);
            if (!pgTxn.exec("BEGIN ISOLATION LEVEL REPEATABLE READ READ ONLY")) {
                m_lastError = QStringLiteral("regenerate(BEGIN pg): ")
                              + pgTxn.lastError().text();
                tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn);
                cleanup();
                return false;
            }
            pgTxnActive = true;

            // SP4.5 audit fix: the PG connection carries statement_timeout=10000ms
            // (applyPgSessionSettings, reconnect-durable). A FULL regen over the NAS
            // copies tens of MB of image blobs and can exceed that, aborting the
            // snapshot. This is one bounded maintenance read inside a REPEATABLE READ
            // txn -- clear the timeout for its duration (SET LOCAL reverts on
            // COMMIT/ROLLBACK; mirrors the legacy normalizer). Non-fatal on failure.
            {
                QSqlQuery noTimeout(pg);
                if (!noTimeout.exec(QStringLiteral("SET LOCAL statement_timeout = 0")))
                    qWarning() << "OfflineSnapshot::regenToPath: could not clear "
                                  "statement_timeout (continuing with default):"
                               << noTimeout.lastError().text();
            }
        }

        if (!tmpDb.transaction()) {
            m_lastError = QStringLiteral("regenerate(begin): ") + tmpDb.lastError().text();
            rollbackPg();
            tmpDb.close();
            tmpDb = QSqlDatabase();
            QSqlDatabase::removeDatabase(tmpConn);
            cleanup();
            return false;
        }

        // SP4.5 Stage 2a: cooperative cancel. Tears down the in-flight SQLite
        // + PG transactions and the tmp file exactly like the error paths
        // below, then returns true; callers do `if (cancelled()) return false;`
        // before the heavy blob tables so a close/teardown can abort the regen
        // between tables instead of blocking on the full multi-second copy.
        auto cancelled = [&]() -> bool {
            if (!cancel || !cancel->load()) return false;
            m_lastError = QStringLiteral("regenToPath: cancelled");
            tmpDb.rollback(); rollbackPg(); tmpDb.close();
            tmpDb = QSqlDatabase();
            QSqlDatabase::removeDatabase(tmpConn); cleanup();
            return true;
        };

        // SP4.5 Stage 2b: centralized teardown for the incremental refresh paths
        // (mirrors the inline error teardown the full-copy blocks use).
        auto bail = [&](const QString& where, const QSqlError& e) -> bool {
            m_lastError = QStringLiteral("regenerate(%1): ").arg(where) + e.text();
            tmpDb.rollback(); rollbackPg(); tmpDb.close();
            tmpDb = QSqlDatabase(); QSqlDatabase::removeDatabase(tmpConn); cleanup();
            return false;
        };
        // Incremental mode copies the prior snapshot, so each refreshed table
        // must be emptied before re-inserting. Returns false (after teardown) on
        // error so callers can `if (incremental && !deleteAll("t")) return false;`.
        auto deleteAll = [&](const char* table) -> bool {
            QSqlQuery d(tmpDb);
            if (!d.exec(QStringLiteral("DELETE FROM %1").arg(table)))
                return bail(QStringLiteral("DELETE %1").arg(table), d.lastError());
            return true;
        };

        // H14: column-arity gate for one table's SELECT / INSERT / bind-loop
        // triple. A drift silently mis-copies data, so it aborts the regen the
        // same way a SQL error does rather than writing a plausible-looking but
        // wrong snapshot: a stale-but-correct previous snapshot is strictly
        // better than a silently mis-copied fresh one, and the atomic
        // tmp-then-promote write below guarantees the old snapshot survives an
        // aborted regen (nothing touches prodPath until the final MoveFileExW).
        // checkColumnArity has already logged at critical severity and filled
        // m_lastError by the time this returns false.
        auto arityOk = [&](QSqlQuery& sel, int insertPlaceholders, int loopBound,
                           const char* table) -> bool {
            if (checkColumnArity(sel, insertPlaceholders, loopBound, table,
                                 &m_lastError))
                return true;
            tmpDb.rollback(); rollbackPg(); tmpDb.close();
            tmpDb = QSqlDatabase(); QSqlDatabase::removeDatabase(tmpConn); cleanup();
            return false;
        };

        // v3 Phase 3c: the three long-format tables copy with exactly the same
        // SELECT / INSERT / bind-loop triple as every block below, with no blob
        // or diff special-casing, so they share one helper instead of a ninth
        // and tenth verbatim copy of it. The eight pre-existing explicit blocks
        // are deliberately NOT refactored onto it: 3c's premise is a near-zero
        // regression surface, and rewriting a working copy path buys nothing.
        //
        // checkColumnArity (via arityOk) runs for each table, same as everywhere
        // else. These triples are hand-maintained by definition, and H14 is
        // precisely about them drifting.
        auto copyTable = [&](const char* table, const char* selectSql,
                             const char* insertSql, int expectCols,
                             const QString& label) -> bool {
            QSqlQuery src(pg);
            if (!src.exec(selectSql))
                return bail(QStringLiteral("SELECT %1").arg(table), src.lastError());
            const int kCols = src.record().count();
            if (!arityOk(src, expectCols, kCols, table)) return false;
            step(label);
            if (incremental && !deleteAll(table)) return false;
            QSqlQuery dst(tmpDb);
            dst.prepare(insertSql);
            while (src.next()) {
                for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec())
                    return bail(QStringLiteral("INSERT %1").arg(table), dst.lastError());
            }
            return true;
        };

        // SP4.5 Stage 2b: the live content fingerprint, captured INSIDE the PG
        // REPEATABLE READ txn so it matches the data copied below. Reused by the
        // image diff (which image table changed) and persisted into _snapshot_meta.
        // It also reports whether this server has the v3 long-format tables --
        // one probe, used for both, and it must happen before any statement that
        // would fail on their absence and poison the transaction.
        bool longTables = false;
        const QString liveFp = snapshotContentFingerprint(pg, &longTables);

        // Image-table refresh. Full rebuild: straight copy. Incremental: if this
        // table's fingerprint segment is unchanged, the copied blobs are already
        // correct -> skip (0 bytes pulled from PG, the whole point). Otherwise
        // diff by (id, updated_at) and pull ONLY new/changed blobs. image_data is
        // column index 4 in every image SELECT.
        auto refreshImages = [&](const char* table, const char* selectAllSql,
                                 const char* insertSql, int expectCols,
                                 const QString& label) -> bool {
            step(label);
            if (!incremental) {
                QSqlQuery src(pg);
                if (!src.exec(selectAllSql))
                    return bail(QStringLiteral("SELECT %1").arg(table), src.lastError());
                const int kCols = src.record().count();
                if (!arityOk(src, expectCols, kCols, table)) return false;
                QSqlQuery dst(tmpDb); dst.prepare(insertSql);
                while (src.next()) {
                    for (int c = 0; c < kCols; ++c) {
                        if (c == 4) dst.bindValue(c, src.value(c).toByteArray());
                        else        dst.bindValue(c, src.value(c));
                    }
                    if (!dst.exec())
                        return bail(QStringLiteral("INSERT %1").arg(table), dst.lastError());
                }
                return true;
            }
            // Incremental + segment unchanged: the copied blobs are already right.
            if (!segmentChanged(priorFp, liveFp, table)) return true;
            // Changed: diff (id, updated_at) -- delete removed, pull new/changed.
            QHash<qint64, QString> pgRows, tmpRows;
            {
                QSqlQuery q(pg);
                if (!q.exec(QStringLiteral("SELECT id, updated_at FROM %1").arg(table)))
                    return bail(QStringLiteral("diff-pg %1").arg(table), q.lastError());
                while (q.next()) pgRows.insert(q.value(0).toLongLong(), q.value(1).toString());
            }
            {
                QSqlQuery q(tmpDb);
                if (!q.exec(QStringLiteral("SELECT id, updated_at FROM %1").arg(table)))
                    return bail(QStringLiteral("diff-tmp %1").arg(table), q.lastError());
                while (q.next()) tmpRows.insert(q.value(0).toLongLong(), q.value(1).toString());
            }
            // Deletes: rows in tmp no longer in pg.
            for (auto it = tmpRows.constBegin(); it != tmpRows.constEnd(); ++it) {
                if (!pgRows.contains(it.key())) {
                    QSqlQuery d(tmpDb);
                    d.prepare(QStringLiteral("DELETE FROM %1 WHERE id = ?").arg(table));
                    d.addBindValue(it.key());
                    if (!d.exec())
                        return bail(QStringLiteral("diff-del %1").arg(table), d.lastError());
                }
            }
            // Pull: ids new in pg, or whose updated_at differs.
            QVariantList toPull;
            for (auto it = pgRows.constBegin(); it != pgRows.constEnd(); ++it)
                if (!tmpRows.contains(it.key()) || tmpRows.value(it.key()) != it.value())
                    toPull << it.key();
            if (toPull.isEmpty()) return true;
            QStringList ph;
            for (int i = 0; i < toPull.size(); ++i) ph << QStringLiteral("?");
            QSqlQuery src(pg);
            src.prepare(QString::fromUtf8(selectAllSql).replace(
                QLatin1String("ORDER BY id"),
                QStringLiteral("WHERE id IN (%1) ORDER BY id").arg(ph.join(QLatin1Char(',')))));
            for (const QVariant& id : toPull) src.addBindValue(id);
            if (!src.exec())
                return bail(QStringLiteral("diff-sel %1").arg(table), src.lastError());
            const int kCols = src.record().count();
            if (!arityOk(src, expectCols, kCols, table)) return false;
            QSqlQuery del(tmpDb);
            del.prepare(QStringLiteral("DELETE FROM %1 WHERE id = ?").arg(table));
            QSqlQuery dst(tmpDb); dst.prepare(insertSql);
            while (src.next()) {
                del.bindValue(0, src.value(0));   // positional: reused each row
                if (!del.exec())
                    return bail(QStringLiteral("diff-repdel %1").arg(table), del.lastError());
                for (int c = 0; c < kCols; ++c) {
                    if (c == 4) dst.bindValue(c, src.value(c).toByteArray());
                    else        dst.bindValue(c, src.value(c));
                }
                if (!dst.exec())
                    return bail(QStringLiteral("diff-ins %1").arg(table), dst.lastError());
                if (outStats) outStats->imageRowsPulledFromPg++;
            }
            return true;
        };

        // ---- files ---------------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, file_path, file_name, loaded_at, template_version, "
                          "sheet_count, sample_count, added_at, updated_at, updated_by, version, "
                          "app_version "
                          "FROM files ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT files): ")
                              + src.lastError().text();
                tmpDb.rollback(); rollbackPg(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            // Drive the bind loop from the SELECT's actual column count so the
            // loop bound can't silently drift from the checked arity: a future
            // edit that changes the SELECT (and thus kCols) must also keep the
            // INSERT placeholder literal in step or the guard aborts the regen.
            const int kCols = src.record().count();
            if (!arityOk(src, 12, kCols, "files")) return false;
            step(QStringLiteral("Copying files"));
            if (incremental && !deleteAll("files")) return false;
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO files (id, file_path, file_name, loaded_at, "
                        "template_version, sheet_count, sample_count, "
                        "added_at, updated_at, updated_by, version, app_version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT files): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); rollbackPg(); tmpDb.close();
                    tmpDb = QSqlDatabase();
                    QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
                }
            }
        }

        // ---- tests ---------------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, file_id, sheet_name, template_version, "
                          "overall_avg_tpm, overall_stddev_tpm, is_raw_table, "
                          "sort_order, updated_at, updated_by, version, "
                          "raw_grid::text, from_inferred_schema "
                          "FROM tests ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT tests): ")
                              + src.lastError().text();
                tmpDb.rollback(); rollbackPg(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            const int kCols = src.record().count();
            if (!arityOk(src, 13, kCols, "tests")) return false;
            step(QStringLiteral("Copying tests"));
            if (incremental && !deleteAll("tests")) return false;
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO tests (id, file_id, sheet_name, template_version, "
                        "overall_avg_tpm, overall_stddev_tpm, is_raw_table, sort_order, "
                        "updated_at, updated_by, version, raw_grid, from_inferred_schema) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT tests): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); rollbackPg(); tmpDb.close();
                    tmpDb = QSqlDatabase();
                    QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
                }
            }
        }

        // ---- samples -------------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, test_id, sort_order, sample_name, sample_id, date, "
                          "tester, media, viscosity, resistance, voltage, power, "
                          "heating_technology, puffing_regime, initial_oil_mass, "
                          "average_tpm, stddev_tpm, avg_power_density, "
                          "efficiency_percent, total_oil_consumed, total_puffs, "
                          "normalized_tpm, burn_status, clog_status, leak_status, "
                          "updated_at, updated_by, version "
                          "FROM samples ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT samples): ")
                              + src.lastError().text();
                tmpDb.rollback(); rollbackPg(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            const int kCols = src.record().count();
            if (!arityOk(src, 28, kCols, "samples")) return false;
            step(QStringLiteral("Copying samples"));
            if (incremental && !deleteAll("samples")) return false;
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO samples (id, test_id, sort_order, sample_name, "
                        "sample_id, date, tester, media, viscosity, resistance, voltage, "
                        "power, heating_technology, puffing_regime, initial_oil_mass, "
                        "average_tpm, stddev_tpm, avg_power_density, efficiency_percent, "
                        "total_oil_consumed, total_puffs, normalized_tpm, "
                        "burn_status, clog_status, leak_status, "
                        "updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, "
                        "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT samples): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); rollbackPg(); tmpDb.close();
                    tmpDb = QSqlDatabase();
                    QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
                }
            }
        }

        if (cancelled()) return false;

        // ---- data_rows -----------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, sample_id, sort_order, puffs, before_weight, "
                          "after_weight, draw_pressure, resistance, smell, clog, notes, "
                          "tpm, tpm_power_density, variation_tpm, oil_consumed, "
                          "puffing_regime, updated_at, updated_by, version "
                          "FROM data_rows ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT data_rows): ")
                              + src.lastError().text();
                tmpDb.rollback(); rollbackPg(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            const int kCols = src.record().count();
            if (!arityOk(src, 19, kCols, "data_rows")) return false;
            step(QStringLiteral("Copying data rows"));
            if (incremental && !deleteAll("data_rows")) return false;
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO data_rows (id, sample_id, sort_order, puffs, "
                        "before_weight, after_weight, draw_pressure, resistance, smell, "
                        "clog, notes, tpm, tpm_power_density, variation_tpm, oil_consumed, "
                        "puffing_regime, updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT data_rows): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); rollbackPg(); tmpDb.close();
                    tmpDb = QSqlDatabase();
                    QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
                }
            }
        }

        if (cancelled()) return false;

        // ---- images (BYTEA -> BLOB) ----------------------------------------
        if (!refreshImages("images",
                "SELECT id, sample_id, sort_order, file_name, image_data, "
                "layout_x, layout_y, layout_w, layout_h, "
                "crop_x, crop_y, crop_w, crop_h, "
                "updated_at, updated_by, version "
                "FROM images ORDER BY id",
                "INSERT INTO images (id, sample_id, sort_order, file_name, "
                "image_data, layout_x, layout_y, layout_w, layout_h, "
                "crop_x, crop_y, crop_w, crop_h, "
                "updated_at, updated_by, version) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                16, QStringLiteral("Copying images"))) return false;

        // ---- sensory_sessions (JSONB -> TEXT) ------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, session_name, tester_name, assessor_name, media, "
                          "puff_length, date, timestamp, "
                          "json_data::text, layout_json::text, "
                          "updated_at, updated_by, version, "
                          "app_version "
                          "FROM sensory_sessions ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT sensory_sessions): ")
                              + src.lastError().text();
                tmpDb.rollback(); rollbackPg(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            const int kCols = src.record().count();
            if (!arityOk(src, 14, kCols, "sensory_sessions")) return false;
            step(QStringLiteral("Copying sensory sessions"));
            if (incremental && !deleteAll("sensory_sessions")) return false;
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO sensory_sessions (id, session_name, tester_name, "
                        "assessor_name, media, puff_length, date, timestamp, "
                        "json_data, layout_json, updated_at, updated_by, version, "
                        "app_version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT sensory_sessions): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); rollbackPg(); tmpDb.close();
                    tmpDb = QSqlDatabase();
                    QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
                }
            }
        }

        if (cancelled()) return false;

        // ---- sensory_images ------------------------------------------------
        if (!refreshImages("sensory_images",
                "SELECT id, session_id, sort_order, file_name, image_data, "
                "layout_x, layout_y, layout_w, layout_h, "
                "crop_x, crop_y, crop_w, crop_h, "
                "updated_at, updated_by, version "
                "FROM sensory_images ORDER BY id",
                "INSERT INTO sensory_images (id, session_id, sort_order, "
                "file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                "crop_x, crop_y, crop_w, crop_h, "
                "updated_at, updated_by, version) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                16, QStringLiteral("Copying sensory images"))) return false;

        // ---- detailed_sensory_sessions -------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, session_name, tester_name, assessor_name, media, "
                          "date, timestamp, json_data::text, "
                          "updated_at, updated_by, version, "
                          "app_version "
                          "FROM detailed_sensory_sessions ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT detailed_sensory_sessions): ")
                              + src.lastError().text();
                tmpDb.rollback(); rollbackPg(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            const int kCols = src.record().count();
            if (!arityOk(src, 12, kCols, "detailed_sensory_sessions")) return false;
            step(QStringLiteral("Copying detailed sensory sessions"));
            if (incremental && !deleteAll("detailed_sensory_sessions")) return false;
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO detailed_sensory_sessions (id, session_name, "
                        "tester_name, assessor_name, media, date, timestamp, json_data, "
                        "updated_at, updated_by, version, app_version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT detailed_sensory_sessions): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); rollbackPg(); tmpDb.close();
                    tmpDb = QSqlDatabase();
                    QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
                }
            }
        }

        if (cancelled()) return false;

        // ---- detailed_sensory_images ---------------------------------------
        if (!refreshImages("detailed_sensory_images",
                "SELECT id, session_id, sort_order, file_name, image_data, "
                "layout_x, layout_y, layout_w, layout_h, "
                "crop_x, crop_y, crop_w, crop_h, "
                "updated_at, updated_by, version "
                "FROM detailed_sensory_images ORDER BY id",
                "INSERT INTO detailed_sensory_images (id, session_id, sort_order, "
                "file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                "crop_x, crop_y, crop_w, crop_h, "
                "updated_at, updated_by, version) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                16, QStringLiteral("Copying detailed sensory images"))) return false;

        // ---- settings ------------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT key, value, updated_at, updated_by, version "
                          "FROM settings ORDER BY key")) {
                m_lastError = QStringLiteral("regenerate(SELECT settings): ")
                              + src.lastError().text();
                tmpDb.rollback(); rollbackPg(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            const int kCols = src.record().count();
            if (!arityOk(src, 5, kCols, "settings")) return false;
            step(QStringLiteral("Copying settings"));
            if (incremental && !deleteAll("settings")) return false;
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO settings (key, value, updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT settings): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); rollbackPg(); tmpDb.close();
                    tmpDb = QSqlDatabase();
                    QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
                }
            }
        }

        // ---- v3 Phase 3c: the long-format tables (snapshot schema v4) ------
        // Open metrics (DataRow::extra / SampleResult::extra) live ONLY here --
        // they have no wide-column equivalent anywhere - so without these three
        // copies every custom column silently disappears the moment the NAS is
        // unreachable, which is the failure the whole sub-phase exists to stop.
        //
        // metric_defs first: it is the vocabulary the other two join to, and
        // copying it in full (not just the keys this database happens to use)
        // keeps the mirror a faithful projection.
        if (longTables) {
            if (!copyTable("metric_defs",
                    "SELECT id, kind, key, display_name, value_type, unit, role, "
                    "tags::text, updated_at, updated_by, version "
                    "FROM metric_defs ORDER BY id",
                    "INSERT INTO metric_defs (id, kind, key, display_name, "
                    "value_type, unit, role, tags, updated_at, updated_by, version) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    11, QStringLiteral("Copying metric definitions"))) return false;

            if (!copyTable("measurements",
                    "SELECT id, sample_id, metric_id, sort_order, value_num, "
                    "value_text, updated_at, updated_by, version "
                    "FROM measurements ORDER BY id",
                    "INSERT INTO measurements (id, sample_id, metric_id, sort_order, "
                    "value_num, value_text, updated_at, updated_by, version) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    9, QStringLiteral("Copying measurements"))) return false;

            if (!copyTable("sample_headers",
                    "SELECT id, sample_id, field_id, value_num, value_text, "
                    "updated_at, updated_by, version "
                    "FROM sample_headers ORDER BY id",
                    "INSERT INTO sample_headers (id, sample_id, field_id, "
                    "value_num, value_text, updated_at, updated_by, version) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    8, QStringLiteral("Copying sample headers"))) return false;
        } else if (incremental) {
            // No long tables on this server, but the prior snapshot we copied
            // may hold rows from one that had them. The mirror follows the
            // source: empty them rather than leaving stale extras behind.
            for (const char* t : kLongFormatTables)
                if (!deleteAll(t)) return false;
        }

        // R7: capture the freshness stamp from the SERVER clock, not the
        // client clock. The client's wall clock may be NTP-skewed (and the
        // offline-era display is most meaningful relative to the authoritative
        // database). We read now() while the REPEATABLE READ transaction is
        // still open, so the stamp is the transaction-start instant -- exactly
        // the point-in-time this snapshot represents. Fall back to the client
        // clock only if the server read fails (better an approximate stamp
        // than none; the data copy already succeeded above).
        QString serverStamp;
        {
            QSqlQuery pgNow(pg);
            if (pgNow.exec("SELECT now() AT TIME ZONE 'UTC'") && pgNow.next()) {
                QDateTime dt = pgNow.value(0).toDateTime();
                dt.setTimeZone(QTimeZone::UTC);
                if (dt.isValid()) {
                    if (outServerTimeUtc) *outServerTimeUtc = dt;
                    serverStamp = dt.toString(Qt::ISODateWithMs);
                }
            }
        }

        // SP4.5: capture the content fingerprint from the SAME REPEATABLE READ
        // transaction, so it reflects exactly the data this snapshot copied. It
        // is persisted into _snapshot_meta below and compared on the next close
        // to skip an unchanged full regen.
        const QString contentFp = liveFp;   // SP4.5 2b: computed once, above (in-txn)
        if (outFingerprint) *outFingerprint = contentFp;

        // All PG SELECTs are done -- close the consistency transaction. We
        // can release this early because nothing below reads PG again.
        // COMMIT vs ROLLBACK is equivalent for a READ ONLY transaction
        // (nothing to commit) -- COMMIT is the conventional close.
        {
            QSqlQuery pgEnd(pg);
            pgEnd.exec("COMMIT");
            pgTxnActive = false;
        }

        // ---- _snapshot_meta ------------------------------------------------
        {
            // R7 follow-up (T2 review minor): distrusting the client clock is
            // the whole point of sourcing the stamp from the PG server. If the
            // SELECT now() read above failed (serverStamp empty), we fall back
            // to the client clock so the snapshot still has SOME freshness
            // stamp -- but that fallback must NOT be silent, because a skewed
            // client clock makes the offline-era display misleading. Warn LOUD.
            if (serverStamp.isEmpty()) {
                qWarning() << "OfflineSnapshot::regenerate -- could not read the "
                              "PG server clock; falling back to the (possibly "
                              "NTP-skewed) CLIENT clock for snapshot_taken_at. "
                              "The offline-era display may be inaccurate.";
            }
            const QString stamp = !serverStamp.isEmpty()
                ? serverStamp
                : QDateTime::currentDateTimeUtc().toString(Qt::ISODateWithMs);
            QSqlQuery dst(tmpDb);
            // INSERT OR REPLACE: in incremental mode the copied tmp already has
            // these meta keys (key is the PRIMARY KEY), so a plain INSERT would
            // collide. Harmless for the full path (empty table).
            dst.prepare("INSERT OR REPLACE INTO _snapshot_meta (key, value) VALUES (?, ?)");
            dst.bindValue(0, "snapshot_taken_at");
            dst.bindValue(1, stamp);
            if (!dst.exec()) {
                m_lastError = QStringLiteral("regenerate(INSERT meta): ")
                              + dst.lastError().text();
                tmpDb.rollback(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            dst.bindValue(0, "source_schema_version");
            dst.bindValue(1, QString::number(kSnapshotSchemaVersion));
            if (!dst.exec()) {
                m_lastError = QStringLiteral("regenerate(INSERT meta v): ")
                              + dst.lastError().text();
                tmpDb.rollback(); tmpDb.close();
                tmpDb = QSqlDatabase();
                QSqlDatabase::removeDatabase(tmpConn); cleanup(); return false;
            }
            // SP4.5: persist the content fingerprint so the next close can skip
            // an unchanged regen. A non-fatal failure here just means the next
            // close won't be able to skip (it will regenerate, the safe default).
            dst.bindValue(0, "content_fingerprint");
            dst.bindValue(1, contentFp);
            if (!dst.exec()) {
                qWarning() << "OfflineSnapshot::regenerate -- could not store "
                              "content_fingerprint (next close will regen):"
                           << dst.lastError().text();
            }
        }

        if (!tmpDb.commit()) {
            m_lastError = QStringLiteral("regenerate(commit): ") + tmpDb.lastError().text();
            tmpDb.rollback();
            tmpDb.close();
            tmpDb = QSqlDatabase();
            QSqlDatabase::removeDatabase(tmpConn);
            cleanup();
            return false;
        }

        // R7: fold the WAL back into the main database file and truncate it so
        // the tmp is a SINGLE self-contained file. After this the atomic
        // promotion only needs to move one file -- there are no -wal/-shm
        // siblings that could be left behind (carrying a stale WAL over the
        // production path would corrupt the promoted snapshot). A checkpoint
        // failure is non-fatal: the data is committed; we proceed and rely on
        // the post-move sidecar cleanup as a backstop.
        {
            QSqlQuery cp(tmpDb);
            cp.exec("PRAGMA wal_checkpoint(TRUNCATE)");
        }
        tmpDb.close();
    }
    // tmpDb goes out of scope; remove the named connection BEFORE the file-
    // system swap so Windows lets us delete/rename the .tmp file. (Safe here
    // because the local handle is out of scope; in the failure paths above
    // we cleared it explicitly before this call.)
    QSqlDatabase::removeDatabase(tmpConn);

    // ---- Atomic promotion (R7) ---------------------------------------------
    // SQLite on Windows holds a file handle until removeDatabase() completes,
    // which is why the temp connection is removed above.
    //
    // The OLD path here was crash-UNSAFE: it QFile::remove(prodPath) and THEN
    // QFile::rename(tmp, prod). A crash (or power loss) in the gap between the
    // delete and the rename left the user with NO snapshot at all -- exactly
    // the failure this store exists to prevent. (QFile::rename can't overwrite
    // on Windows, which is why the delete-first was there.)
    //
    // MoveFileExW with MOVEFILE_REPLACE_EXISTING does the replace in one
    // syscall -- there is never a moment where the destination doesn't exist.
    // MOVEFILE_WRITE_THROUGH flushes the rename through to disk before
    // returning, so a crash can't leave a dangling directory entry pointing at
    // a half-promoted file. On failure the previous snapshot is untouched and
    // the tmp is cleaned up.
#ifdef Q_OS_WIN
    const std::wstring tmpW  =
        QDir::toNativeSeparators(tmpPath).toStdWString();
    const std::wstring prodW =
        QDir::toNativeSeparators(prodPath).toStdWString();
    if (!MoveFileExW(tmpW.c_str(), prodW.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        m_lastError = QStringLiteral("regenerate: atomic replace failed (err %1) for %2")
                          .arg(GetLastError()).arg(prodPath);
        cleanup();   // success is still false -> removes the tmp
        return false;
    }
#else
    // Non-Windows path (not shipped; keeps the file compilable elsewhere).
    QFile::remove(prodPath);
    if (!QFile::rename(tmpPath, prodPath)) {
        m_lastError = QStringLiteral("regenerate: cannot rename tmp to ") + prodPath;
        cleanup();
        return false;
    }
#endif

    // After a successful move, sweep away any STALE sidecars belonging to the
    // OLD production snapshot. The freshly-promoted file was checkpointed +
    // closed (single self-contained file), so a leftover prod -wal/-shm from a
    // previous run would be a phantom journal SQLite must not pick up.
    QFile::remove(prodPath + "-wal");
    QFile::remove(prodPath + "-shm");

    if (progress) progress(kRegenPhases, kRegenPhases, QStringLiteral("Done"));
    success = true;
    cleanup();
    return true;
}

bool OfflineSnapshot::regenerate(PostgresConnection* live) {
    return regenerate(live, RegenProgress{});
}

bool OfflineSnapshot::regenerate(PostgresConnection* live, const RegenProgress& progress) {
    m_lastError.clear();
    if (!live || !live->isOpen()) {
        m_lastError = QStringLiteral("regenerate: PostgresConnection is not open");
        return false;
    }
    // The instance owns m_db (the read-only snapshot handle); release it before
    // regenToPath promotes a new prod file underneath us.
    if (m_open) close();
    QString fp, err;
    QDateTime srv;
    const bool ok = regenToPath(live, path(), /*cancel*/nullptr, &fp, &srv, &err, progress);
    if (ok && srv.isValid()) m_lastRegenServerTime = srv;
    if (!ok) m_lastError = err;
    return ok;
}

// ----------------------------------------------------------------------------
// openReadOnly
// ----------------------------------------------------------------------------
bool OfflineSnapshot::openReadOnly() {
    m_lastError.clear();
    m_lastOpenWasDecodeFailure = false;
    if (m_open) return true;

    const QString p = path();
    if (!QFile::exists(p)) {
        m_lastError = QStringLiteral("openReadOnly: snapshot file does not exist: ") + p;
        return false;
    }

    // R5 (MIP resilience): if the snapshot file is MIP/AIP-encrypted at rest
    // (the "%TSD-Header-###%" marker in its first bytes), SQLite cannot open it.
    // Unlike the JSON recovery store (which reads plaintext over a python STDOUT
    // pipe), QSqlDatabase needs a real file PATH, so a temp is unavoidable here:
    // the MIP-allowlisted python WRITES the plaintext temp and we open THAT
    // read-only. Per this project's MIP convention a python-written file is
    // UNLABELED -- the same reliance the whole MIP strategy rests on -- and
    // decryptToTempViaPython re-checks the temp is not still ciphertext before
    // returning. If work-machine testing ever shows this SQLite temp re-labels,
    // escalate to a python-side row export (read rows over the pipe and rebuild
    // a local db); do NOT build that speculatively now. On ANY failure this is a
    // DECODE FAILURE (not mere absence) -- flag it so the caller warns loudly.
    QString openPath = p;
    if (QFile::exists(m_decryptedSnapshotTmp)) {
        QFile::remove(m_decryptedSnapshotTmp);
        m_decryptedSnapshotTmp.clear();
    }
    if (looksEncrypted(p)) {
        QString err;
        const QString tmp = decryptToTempViaPython(p, resolveBundledPython(), err);
        if (tmp.isEmpty()) {
            m_lastError = QStringLiteral("openReadOnly: snapshot is MIP-encrypted "
                                         "and could not be decrypted (")
                          + p + QStringLiteral("): ") + err;
            m_lastOpenWasDecodeFailure = true;
            qWarning().noquote() << "OfflineSnapshot::openReadOnly --" << m_lastError;
            return false;
        }
        m_decryptedSnapshotTmp = tmp;
        openPath = tmp;
    }

    if (QSqlDatabase::contains(m_connName)) {
        QSqlDatabase::removeDatabase(m_connName);
    }

    m_db = QSqlDatabase::addDatabase("QSQLITE", m_connName);
    m_db.setDatabaseName(openPath);
    // Qt 6's QSQLITE driver expects flag-style options without "=1"; the
    // "QSQLITE_OPEN_READONLY=1" form is rejected as "Unsupported option".
    m_db.setConnectOptions("QSQLITE_OPEN_READONLY");
    if (!m_db.open()) {
        m_lastError = QStringLiteral("openReadOnly: ") + m_db.lastError().text();
        m_db = QSqlDatabase();
        QSqlDatabase::removeDatabase(m_connName);
        if (!m_decryptedSnapshotTmp.isEmpty()) {
            QFile::remove(m_decryptedSnapshotTmp);
            m_decryptedSnapshotTmp.clear();
        }
        return false;
    }

    // Validate -- if the file isn't a snapshot we built, _snapshot_meta is
    // missing and the SELECT errors out.
    {
        QSqlQuery q(m_db);
        if (!q.exec("SELECT 1 FROM _snapshot_meta LIMIT 1")) {
            m_lastError = QStringLiteral("openReadOnly: not a DataViewer snapshot (")
                          + q.lastError().text() + ")";
            m_db.close();
            m_db = QSqlDatabase();
            QSqlDatabase::removeDatabase(m_connName);
            return false;
        }
    }

    // R7: reject a stale source_schema_version. A snapshot written by an older
    // build (e.g. a v2.4.1 install that wrote schema v2) may have a DIFFERENT
    // table layout than this build expects; trusting it would mis-read columns
    // silently. We treat a mismatch as "no usable snapshot": close, set
    // m_lastError, return false. The caller already handles false as
    // absence-of-snapshot (it falls through to online mode and regenerates a
    // fresh v3 snapshot on the next clean online close), so a v2->v3 upgrade
    // degrades gracefully -- no crash, just one offline session without the
    // local cache until the next clean close refreshes it.
    {
        QSqlQuery q(m_db);
        q.prepare("SELECT value FROM _snapshot_meta WHERE key = ?");
        q.addBindValue("source_schema_version");
        if (q.exec() && q.next()) {
            const int fileVer = q.value(0).toInt();
            if (fileVer != kSnapshotSchemaVersion) {
                m_lastError = QStringLiteral(
                    "openReadOnly: snapshot schema v%1, expected v%2 "
                    "-- ignoring stale snapshot (will regenerate on next "
                    "clean online close)")
                        .arg(fileVer).arg(kSnapshotSchemaVersion);
                qWarning() << "OfflineSnapshot::openReadOnly --" << m_lastError;
                m_db.close();
                m_db = QSqlDatabase();
                QSqlDatabase::removeDatabase(m_connName);
                return false;
            }
        }
        // A missing source_schema_version row is tolerated here (pre-versioned
        // snapshots from before the field existed); only a PRESENT-and-
        // mismatching value rejects. Such legacy files predate v3 and will be
        // refreshed on the next clean close regardless.
    }

    m_open = true;

    // R7: staleness is a WARN, never a reject. A stale-but-readable offline
    // cache is strictly better than none when the NAS is down, so we open it
    // and merely log loudly (the offline banner can echo this). Only a schema
    // mismatch (above) is fatal.
    {
        const QDateTime taken = snapshotTakenAt();
        if (taken.isValid()) {
            const qint64 ageDays =
                taken.daysTo(QDateTime::currentDateTimeUtc());
            if (ageDays > kSnapshotStaleWarnDays) {
                qWarning().nospace()
                    << "OfflineSnapshot::openReadOnly -- snapshot is "
                    << ageDays << " days old (taken " << taken.toString(Qt::ISODate)
                    << "); offline data may be out of date. Opening anyway "
                       "(stale-but-usable beats no snapshot).";
            }
        }
    }

    return true;
}

void OfflineSnapshot::close() {
    if (m_open) {
        m_db.close();
    }
    m_db = QSqlDatabase();
    if (QSqlDatabase::contains(m_connName)) {
        QSqlDatabase::removeDatabase(m_connName);
    }
    m_open = false;

    // Tear down the writable queue connection alongside the read-only
    // snapshot connection — both are owned by this instance's lifetime.
    if (m_queueOpen) {
        m_queueDb.close();
    }
    m_queueDb = QSqlDatabase();
    if (QSqlDatabase::contains(m_queueConnName)) {
        QSqlDatabase::removeDatabase(m_queueConnName);
    }
    m_queueOpen = false;
    m_queueDegraded = false;   // IMPORTANT 3: stale degraded state must not survive a close

    // R5: drop any decrypted plaintext temp copies now that their connections
    // are released. We never leave a decrypted snapshot lying around longer
    // than the open lifetime.
    if (!m_decryptedSnapshotTmp.isEmpty()) {
        QFile::remove(m_decryptedSnapshotTmp);
        m_decryptedSnapshotTmp.clear();
    }
    if (!m_decryptedQueueTmp.isEmpty()) {
        QFile::remove(m_decryptedQueueTmp);
        m_decryptedQueueTmp.clear();
    }
}

QDateTime OfflineSnapshot::snapshotTakenAt() const {
    if (!m_open) {
        m_lastError = QStringLiteral("snapshotTakenAt: not open");
        return QDateTime();
    }
    QSqlQuery q(m_db);
    q.prepare("SELECT value FROM _snapshot_meta WHERE key = ?");
    q.addBindValue("snapshot_taken_at");
    if (!q.exec() || !q.next()) return QDateTime();
    const QString stamp = q.value(0).toString();
    // The stamp is an ISO-8601 string written by regenerate() (R7: sourced
    // from the PG server clock as a UTC-tagged value, so it carries a 'Z'
    // suffix; the pre-R7 client-clock path also wrote a UTC-tagged string).
    // Either way the offset is encoded in the string and Qt parses the spec
    // from it -- no explicit setTimeZone() needed.
    QDateTime dt = QDateTime::fromString(stamp, Qt::ISODateWithMs);
    if (!dt.isValid()) dt = QDateTime::fromString(stamp, Qt::ISODate);
    return dt;
}

// ----------------------------------------------------------------------------
// Read accessors
// ----------------------------------------------------------------------------
QVector<FileRecord> OfflineSnapshot::listFiles() const {
    QVector<FileRecord> records;
    if (!m_open) {
        m_lastError = QStringLiteral("listFiles: snapshot not open");
        return records;
    }
    QSqlQuery q(m_db);
    // F6: include added_at and order by it (mirrors DatabaseManager::listFiles)
    // so versioned re-adds of one path appear newest-first in the DB browser.
    if (!q.exec("SELECT id, file_path, file_name, loaded_at, template_version, "
                "sheet_count, sample_count, added_at, app_version "
                "FROM files ORDER BY added_at DESC, id DESC")) {
        m_lastError = QStringLiteral("listFiles(SELECT): ") + q.lastError().text();
        return records;
    }
    while (q.next()) {
        FileRecord r;
        r.id              = q.value(0).toInt();
        r.filePath        = q.value(1).toString();
        r.fileName        = q.value(2).toString();
        r.loadedAt        = q.value(3).toString();
        r.templateVersion = q.value(4).toString();
        r.sheetCount      = q.value(5).toInt();
        r.sampleCount     = q.value(6).toInt();
        r.addedAt         = q.value(7).toString();
        r.appVersion      = q.value(8).toString();
        records.append(r);
    }
    return records;
}

FileResult OfflineSnapshot::loadFileByPath(const QString& filePath) const {
    FileResult result;
    if (!m_open) {
        m_lastError = QStringLiteral("loadFileByPath: snapshot not open");
        return result;
    }
    QSqlQuery q(m_db);
    // F6: a path may have several versioned rows; return the most recently
    // added one (mirrors DatabaseManager::loadFileByPath). added_at is TEXT in
    // the snapshot (ISO-8601), so a lexical DESC sort is also chronological;
    // id DESC is the tiebreak for legacy rows whose added_at backfilled NULL.
    q.prepare("SELECT id FROM files WHERE file_path = ? "
              "ORDER BY added_at DESC, id DESC LIMIT 1");
    q.addBindValue(filePath);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadFileByPath(SELECT): ") + q.lastError().text();
        return result;
    }
    if (q.next()) return loadFile(q.value(0).toInt());
    return result;
}

FileResult OfflineSnapshot::loadFile(int id) const {
    FileResult result;
    if (!m_open) {
        m_lastError = QStringLiteral("loadFile: snapshot not open");
        return result;
    }

    {
        QSqlQuery q(m_db);
        q.prepare("SELECT id, version, file_path, file_name, template_version "
                  "FROM files WHERE id = ?");
        q.addBindValue(id);
        if (!q.exec()) {
            m_lastError = QStringLiteral("loadFile(SELECT files): ") + q.lastError().text();
            return result;
        }
        if (!q.next()) return result;
        result.id              = q.value(0).toInt();
        result.version         = q.value(1).toInt();
        result.filePath        = q.value(2).toString();
        result.fileName        = q.value(3).toString();
        result.templateVersion = q.value(4).toString();
    }

    struct TestInfo {
        int id; QString sheetName; QString templateVersion;
        double avgTPM; double stddevTPM; bool isRaw;
        QString rawGrid;  // TEXT column; empty when NULL (non-raw sheets)
        bool fromInferred; // from_inferred_schema (appended last)
    };
    QVector<TestInfo> tests;
    {
        QSqlQuery q(m_db);
        q.prepare("SELECT id, sheet_name, template_version, overall_avg_tpm, "
                  "overall_stddev_tpm, is_raw_table, raw_grid, from_inferred_schema FROM tests "
                  "WHERE file_id = ? ORDER BY sort_order");
        q.addBindValue(id);
        if (!q.exec()) {
            m_lastError = QStringLiteral("loadFile(SELECT tests): ") + q.lastError().text();
            return result;
        }
        while (q.next()) {
            tests.append({q.value(0).toInt(), q.value(1).toString(),
                          q.value(2).toString(), q.value(3).toDouble(),
                          q.value(4).toDouble(), q.value(5).toInt() != 0,
                          q.value(6).toString(), q.value(7).toInt() != 0});
        }
    }

    for (const TestInfo& ti : tests) {
        SheetResult sheet;
        sheet.sheetName        = ti.sheetName;
        sheet.templateVersion  = ti.templateVersion;
        sheet.overallAvgTPM    = ti.avgTPM;
        sheet.overallStdDevTPM = ti.stddevTPM;
        sheet.isRawTable       = ti.isRaw;
        sheet.fromInferredSchema = ti.fromInferred;
        result.sheetNames.append(ti.sheetName);
        // Reconstruct raw grid from TEXT (no-op when ti.rawGrid is empty/null).
        if (ti.isRaw)
            rawGridFromJson(ti.rawGrid, sheet.rawHeaders, sheet.rawRows);

        struct SampleInfo {
            int id; QString name, sampleID, date, tester, media;
            double visc, res, volt, pwr; QString heatingTech, puffRegime;
            double initOil, avgTPM, stdDev, avgPD, effPct, totOil;
            int totPuffs; double normTPM;
            QString burn, clog, leak;
        };
        QVector<SampleInfo> sampleInfos;
        {
            QSqlQuery q(m_db);
            q.prepare("SELECT id, sample_name, sample_id, date, tester, media, viscosity, "
                      "resistance, voltage, power, heating_technology, puffing_regime, "
                      "initial_oil_mass, average_tpm, stddev_tpm, avg_power_density, "
                      "efficiency_percent, total_oil_consumed, total_puffs, normalized_tpm, "
                      "burn_status, clog_status, leak_status "
                      "FROM samples WHERE test_id = ? ORDER BY sort_order");
            q.addBindValue(ti.id);
            if (!q.exec()) {
                // Short-circuit so the SELECT samples error doesn't get
                // overwritten by a later loadFile(SELECT data_rows) error
                // (or by m_lastError churn in subsequent iterations).
                m_lastError = QStringLiteral("loadFile(SELECT samples): ")
                              + q.lastError().text();
                return result;
            }
            while (q.next()) {
                SampleInfo si;
                si.id         = q.value(0).toInt();
                si.name       = q.value(1).toString();
                si.sampleID   = q.value(2).toString();
                si.date       = q.value(3).toString();
                si.tester     = q.value(4).toString();
                si.media      = q.value(5).toString();
                si.visc       = q.value(6).toDouble();
                si.res        = q.value(7).toDouble();
                si.volt       = q.value(8).toDouble();
                si.pwr        = q.value(9).toDouble();
                si.heatingTech= q.value(10).toString();
                si.puffRegime = q.value(11).toString();
                si.initOil    = q.value(12).toDouble();
                si.avgTPM     = q.value(13).toDouble();
                si.stdDev     = q.value(14).toDouble();
                si.avgPD      = q.value(15).toDouble();
                si.effPct     = q.value(16).toDouble();
                si.totOil     = q.value(17).toDouble();
                si.totPuffs   = q.value(18).toInt();
                si.normTPM    = q.value(19).toDouble();
                si.burn       = q.value(20).toString();
                si.clog       = q.value(21).toString();
                si.leak       = q.value(22).toString();
                sampleInfos.append(si);
            }
        }

        bool sheetHasRegime = false;
        for (const SampleInfo& si : sampleInfos) {
            SampleResult sr;
            sr.sampleName          = si.name;
            sr.sampleID            = si.sampleID;
            sr.date                = si.date;
            sr.tester              = si.tester;
            sr.media               = si.media;
            sr.viscosity           = si.visc;
            sr.resistance          = si.res;
            sr.voltage             = si.volt;
            sr.power               = si.pwr;
            sr.heatingTechnology   = si.heatingTech;
            sr.puffingRegime       = si.puffRegime;
            sr.initialOilMass      = si.initOil;
            sr.averageTPM          = si.avgTPM;
            sr.stdDevTPM           = si.stdDev;
            sr.averagePowerDensity = si.avgPD;
            sr.efficiencyPercent   = si.effPct;
            sr.totalOilConsumed    = si.totOil;
            sr.totalPuffs          = si.totPuffs;
            sr.normalizedTPM       = si.normTPM;
            sr.burnStatus          = si.burn;
            sr.clogStatus          = si.clog;
            sr.leakStatus          = si.leak;

            // data rows
            {
                QSqlQuery q(m_db);
                q.prepare("SELECT id, puffs, before_weight, after_weight, draw_pressure, "
                          "resistance, smell, clog, notes, tpm, tpm_power_density, "
                          "variation_tpm, oil_consumed, puffing_regime "
                          "FROM data_rows WHERE sample_id = ? ORDER BY sort_order");
                q.addBindValue(si.id);
                if (!q.exec()) {
                    // Short-circuit: don't let m_lastError get clobbered by
                    // the next sample's images query.
                    m_lastError = QStringLiteral("loadFile(SELECT data_rows): ")
                                  + q.lastError().text();
                    return result;
                }
                while (q.next()) {
                    DataRow dr;
                    dr.id              = q.value(0).toInt();
                    dr.puffs           = q.value(1).toDouble();
                    dr.beforeWeight    = q.value(2).toDouble();
                    dr.afterWeight     = q.value(3).toDouble();
                    dr.drawPressure    = q.value(4).toDouble();
                    dr.resistance      = q.value(5).toDouble();
                    dr.smell           = q.value(6).toString();
                    dr.clog            = q.value(7).toString();
                    dr.notes           = q.value(8).toString();
                    dr.tpm             = q.value(9).toDouble();
                    dr.tpmPowerDensity = q.value(10).toDouble();
                    dr.variationTPM    = q.value(11).toDouble();
                    dr.oilConsumed     = q.value(12).toDouble();
                    const QVariant pr  = q.value(13);
                    if (!pr.isNull()) { dr.puffingRegime = pr.toString(); sheetHasRegime = true; }
                    sr.rows.append(dr);
                }
            }

            // v3 Phase 3c: this sample's open metrics, the offline mirror of
            // the two supplementary SELECTs DatabaseManager::loadFile issues.
            //
            // REUSES MetricDefCache::decodeValue -- the inverse of the
            // encodeValue the save path calls -- rather than restating the
            // value_type <-> (value_num, value_text) mapping a third time. Two
            // halves of that contract disagreeing is silent data corruption,
            // not a compile error, so they are not allowed to live apart:
            // extend decodeValue if a new type appears, never copy it.
            //
            // Keyed (sample_id, sort_order): `measurements` has no referential
            // link to `data_rows` at all (hazard H22), so the row ORDINAL is the
            // only join there is. The rows above were read ORDER BY sort_order
            // and the save path binds both sort_orders from the same index, so
            // sort_order N is the Nth row. Out-of-range ordinals are dropped
            // rather than trusted - a stray measurement from a deleted row must
            // not re-bind onto a survivor.
            //
            // NO kind / key FILTER, mirroring the Postgres read: in 3c
            // `measurements` holds EXCLUSIVELY open metrics. When 3d puts the
            // standard metrics in there too, this starts pulling `tpm` /
            // `before_weight` / ... into `extra` alongside the wide columns read
            // above, and 3d must resolve that here as well as in
            // DatabaseManager (hazard H25).
            //
            // A failure is logged and skipped, and m_lastError is deliberately
            // NOT set: schema v4 always creates these tables, and a snapshot
            // regenerated against a server without them simply has none - the
            // file comes back complete, minus extras it could not have had.
            {
                QSqlQuery q(m_db);
                q.prepare("SELECT m.sort_order, md.key, md.value_type, "
                          "m.value_num, m.value_text "
                          "FROM measurements m "
                          "JOIN metric_defs md ON md.id = m.metric_id "
                          "WHERE m.sample_id = ? ORDER BY m.sort_order");
                q.addBindValue(si.id);
                if (q.exec()) {
                    while (q.next()) {
                        const int ord = q.value(0).toInt();
                        if (ord < 0 || ord >= sr.rows.size()) continue;
                        const QVariant v = MetricDefCache::decodeValue(
                            q.value(2).toString(), q.value(3), q.value(4));
                        // Both columns NULL. encodeValue never writes that
                        // (index D2: a value with nothing to say is not stored),
                        // so this only shields against a row some other tool put
                        // in the source database.
                        if (!v.isValid()) continue;
                        sr.rows[ord].extra.insert(q.value(1).toString(), v);
                    }
                } else {
                    qWarning().noquote()
                        << "OfflineSnapshot::loadFile -- open metrics not read "
                           "for sample" << si.id << ":" << q.lastError().text();
                }
            }
            {
                // sample_headers is keyed (sample_id, field_id) - a header field
                // is per-sample, so there is no sort_order dimension.
                QSqlQuery q(m_db);
                q.prepare("SELECT md.key, md.value_type, sh.value_num, sh.value_text "
                          "FROM sample_headers sh "
                          "JOIN metric_defs md ON md.id = sh.field_id "
                          "WHERE sh.sample_id = ?");
                q.addBindValue(si.id);
                if (q.exec()) {
                    while (q.next()) {
                        const QVariant v = MetricDefCache::decodeValue(
                            q.value(1).toString(), q.value(2), q.value(3));
                        if (!v.isValid()) continue;
                        sr.extra.insert(q.value(0).toString(), v);
                    }
                } else {
                    qWarning().noquote()
                        << "OfflineSnapshot::loadFile -- open header metrics not "
                           "read for sample" << si.id << ":" << q.lastError().text();
                }
            }

            // images -- materialise BLOBs to disk so imagePaths works
            {
                QSqlDatabase nonConstDb = m_db;
                const QString imageCacheDir =
                    QFileInfo(path()).absolutePath() + "/ImageCache";
                loadImagesForSnapshot(nonConstDb, "images", si.id, "sample_id",
                                      imageCacheDir,
                                      &sr.imagePaths, &sr.imageLayouts, &sr.imageCrops,
                                      &m_lastError);
            }

            sheet.samples.append(sr);
        }
        sheet.hasPerRowRegime = sheetHasRegime;

        for (const SampleResult& sr : sheet.samples) {
            for (const DataRow& dr : sr.rows) {
                if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
                sheet.tpmTrend.append(dr.tpm);
                sheet.puffCounts.append(dr.puffs);
            }
        }

        // Detect incomplete DB data: a raw sheet with no grid, or a normal
        // sheet with aggregate-TPM samples that have no per-row data (legacy /
        // partially-migrated records). Consumed by the Task 8 banner.
        {
            bool incomplete = false;
            if (sheet.isRawTable) {
                incomplete = sheet.rawHeaders.isEmpty();
            } else {
                for (const SampleResult& s : sheet.samples) {
                    if (s.averageTPM > 0.0 && s.rows.isEmpty()) {
                        incomplete = true;
                        break;
                    }
                }
            }
            sheet.dbDataIncomplete = incomplete;
        }

        result.sheets.append(sheet);
    }

    return result;
}

QVector<SensoryRecord> OfflineSnapshot::listSensoryRecords() const {
    QVector<SensoryRecord> result;
    if (!m_open) {
        m_lastError = QStringLiteral("listSensoryRecords: snapshot not open");
        return result;
    }
    // SQLite has no jsonb_array_length / ->> operators by default, so we read
    // json_data into memory and parse it client-side to derive test_title,
    // tester_name, and sample_count. Listings are small (< thousands typically),
    // so this is cheap enough.
    QSqlQuery q(m_db);
    if (!q.exec("SELECT id, session_name, assessor_name, media, date, json_data, app_version "
                "FROM sensory_sessions ORDER BY id DESC")) {
        m_lastError = QStringLiteral("listSensoryRecords(SELECT): ") + q.lastError().text();
        return result;
    }
    while (q.next()) {
        SensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();

        const QByteArray jsonBytes = q.value(5).toString().toUtf8();
        const QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
        if (doc.isObject()) {
            const QJsonObject root = doc.object();
            rec.testTitle   = root.value("test_title").toString();
            rec.testerName  = root.value("tester_name").toString();
            rec.sampleCount = root.value("samples").toArray().size();
        }
        rec.appVersion = q.value(6).toString();   // SP4 A4 (legacy-string flag is online-only)
        result.append(rec);
    }
    return result;
}

SensorySession OfflineSnapshot::loadSensorySession(int id) const {
    SensorySession sess;
    if (!m_open) {
        m_lastError = QStringLiteral("loadSensorySession: snapshot not open");
        return sess;
    }
    QSqlQuery q(m_db);
    q.prepare("SELECT version, json_data FROM sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadSensorySession(SELECT): ") + q.lastError().text();
        return sess;
    }
    if (!q.next()) return sess;

    const int rowVersion = q.value(0).toInt();
    if (!deserializeSensoryJsonLocal(q.value(1).toString().toUtf8(), sess)) return sess;
    sess.id      = id;
    sess.version = rowVersion;

    QSqlDatabase nonConstDb = m_db;
    const QString imageCacheDir =
        QFileInfo(path()).absolutePath() + "/ImageCache";
    loadImagesForSnapshot(nonConstDb, "sensory_images", id, "session_id",
                          imageCacheDir,
                          &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops,
                          &m_lastError);
    return sess;
}

QVector<DetailedSensoryRecord> OfflineSnapshot::listDetailedSensoryRecords() const {
    QVector<DetailedSensoryRecord> result;
    if (!m_open) {
        m_lastError = QStringLiteral("listDetailedSensoryRecords: snapshot not open");
        return result;
    }
    QSqlQuery q(m_db);
    if (!q.exec("SELECT id, session_name, assessor_name, media, date, json_data, app_version "
                "FROM detailed_sensory_sessions ORDER BY id DESC")) {
        m_lastError = QStringLiteral("listDetailedSensoryRecords(SELECT): ")
                      + q.lastError().text();
        return result;
    }
    while (q.next()) {
        DetailedSensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();

        const QByteArray jsonBytes = q.value(5).toString().toUtf8();
        const QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
        if (doc.isObject()) {
            const QJsonObject root = doc.object();
            rec.testTitle   = root.value("test_title").toString();
            rec.testerName  = root.value("tester_name").toString();
            rec.sampleCount = root.value("samples").toArray().size();
        }
        rec.appVersion = q.value(6).toString();   // SP4 A4 (legacy-string flag is online-only)
        result.append(rec);
    }
    return result;
}

DetailedSensorySession OfflineSnapshot::loadDetailedSensorySession(int id) const {
    DetailedSensorySession sess;
    if (!m_open) {
        m_lastError = QStringLiteral("loadDetailedSensorySession: snapshot not open");
        return sess;
    }
    QSqlQuery q(m_db);
    q.prepare("SELECT version, json_data FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadDetailedSensorySession(SELECT): ")
                      + q.lastError().text();
        return sess;
    }
    if (!q.next()) return sess;

    const int rowVersion = q.value(0).toInt();
    if (!deserializeDetailedSensoryJsonLocal(q.value(1).toString().toUtf8(), sess)) return sess;
    sess.id      = id;
    sess.version = rowVersion;

    QSqlDatabase nonConstDb = m_db;
    const QString imageCacheDir =
        QFileInfo(path()).absolutePath() + "/ImageCache";
    loadImagesForSnapshot(nonConstDb, "detailed_sensory_images", id, "session_id",
                          imageCacheDir,
                          &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops,
                          &m_lastError);
    return sess;
}

// ----------------------------------------------------------------------------
// Settings (added Plan C T3 — DatabaseManager routes offline reads here)
// ----------------------------------------------------------------------------
QString OfflineSnapshot::getSetting(const QString& key, const QString& defaultVal) const
{
    if (!m_open) {
        m_lastError = QStringLiteral("getSetting: snapshot not open");
        return defaultVal;
    }
    QSqlQuery q(m_db);
    q.prepare("SELECT value FROM settings WHERE key = ?");
    q.addBindValue(key);
    if (!q.exec()) {
        m_lastError = QStringLiteral("getSetting(SELECT): ") + q.lastError().text();
        return defaultVal;
    }
    if (q.next()) return q.value(0).toString();
    return defaultVal;
}

// ----------------------------------------------------------------------------
// Pending-edit queue (v2.0.1)
//
// Lives in a SEPARATE SQLite file from the read-only snapshot so the
// snapshot connection can stay QSQLITE_OPEN_READONLY without having to
// reopen it as writable. The queue file path is a sibling of the
// snapshot file (so setOverrideDirForTesting() propagates to it).
//
// Schema:
//   pending_edits(id INTEGER PRIMARY KEY,
//                 target_table TEXT NOT NULL,
//                 row_id       INTEGER NOT NULL,
//                 column_name  TEXT,
//                 value_text   TEXT,
//                 captured_at  TEXT NOT NULL)
//
// The column_name + value_text columns hold a per-cell edit; older
// queue entries with column_name IS NULL are NOT touched by
// drainPendingEdits (reserved for future file-level rollups).
// ----------------------------------------------------------------------------
QString OfflineSnapshot::queuePath() const {
    return QFileInfo(path()).absolutePath() + "/pending_edits.sqlite";
}

bool OfflineSnapshot::ensureQueueOpen() const {
    if (m_queueOpen) return true;

    // R5 (MIP resilience): the pending_edits queue is opened WRITABLE. A missing
    // file is normal (first offline edit creates it), so absence is fine. But if
    // an EXISTING queue file is MIP/AIP-encrypted at rest, SQLite cannot open it
    // and every offline edit would silently fail to enqueue (the exact risk this
    // task closes). As with the read-only snapshot above, QSqlDatabase needs a
    // file PATH, so the MIP-allowlisted python WRITES a plaintext temp .sqlite
    // (python-written => UNLABELED by this project's MIP convention;
    // decryptToTempViaPython re-verifies the temp is not still ciphertext) and we
    // open THAT so queued edits remain replayable this session. If the SQLite
    // temp is ever shown to re-label on the work machine, escalate to a
    // python-side row export -- not built speculatively now. On a decode failure
    // we set m_lastError + m_queueDegraded + return false; the caller (LiveSync)
    // surfaces a loud failure instead of swallowing it.
    const QString realQueuePath = queuePath();
    QString openPath = realQueuePath;
    if (!m_decryptedQueueTmp.isEmpty()) {
        QFile::remove(m_decryptedQueueTmp);
        m_decryptedQueueTmp.clear();
    }
    const bool fileExists = QFile::exists(realQueuePath);
    if (fileExists && looksEncrypted(realQueuePath)) {
        QString err;
        const QString tmp =
            decryptToTempViaPython(realQueuePath, resolveBundledPython(), err);
        if (tmp.isEmpty()) {
            m_lastError = QStringLiteral("queue open: pending_edits is "
                                         "MIP-encrypted and could not be "
                                         "decrypted (") + realQueuePath
                          + QStringLiteral("): ") + err;
            // IMPORTANT 3: the queue file EXISTS but is undecodable. Mark it
            // degraded so pendingEditCount()==0 is NOT read as "empty/drained":
            // there may be offline edits stranded inside the ciphertext.
            m_queueDegraded = true;
            qWarning().noquote() << "OfflineSnapshot::ensureQueueOpen --"
                                 << m_lastError;
            return false;
        }
        m_decryptedQueueTmp = tmp;
        openPath = tmp;
    }

    if (QSqlDatabase::contains(m_queueConnName)) {
        QSqlDatabase::removeDatabase(m_queueConnName);
    }
    m_queueDb = QSqlDatabase::addDatabase("QSQLITE", m_queueConnName);
    m_queueDb.setDatabaseName(openPath);
    if (!m_queueDb.open()) {
        m_lastError = QStringLiteral("queue open: ") + m_queueDb.lastError().text();
        m_queueDb = QSqlDatabase();
        QSqlDatabase::removeDatabase(m_queueConnName);
        // IMPORTANT 3: only an EXISTING-but-unopenable queue (locked / corrupt /
        // undecodable) is "degraded" — that is the case where pendingEditCount()
        // would falsely read 0 while edits sit stranded on disk. A genuinely-
        // absent file that SQLite still failed to create is an environmental
        // open error, NOT a degraded queue: there are no stranded edits, so
        // leaving m_queueDegraded false keeps "absent != degraded" honest and
        // matches the documented queueDegraded() invariant.
        if (fileExists)
            m_queueDegraded = true;
        return false;
    }

    // Create the table if first use. Additive ALTER TABLE block follows
    // so a snapshot installed by an earlier build (without
    // column_name / value_text) still upgrades transparently.
    {
        QSqlQuery q(m_queueDb);
        if (!q.exec("CREATE TABLE IF NOT EXISTS pending_edits ("
                    "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                    "target_table TEXT NOT NULL, "
                    "row_id INTEGER NOT NULL, "
                    "column_name TEXT, "
                    "value_text TEXT, "
                    "captured_at TEXT NOT NULL)")) {
            m_lastError = QStringLiteral("queue CREATE: ") + q.lastError().text();
            return false;
        }
    }
    {
        QSqlQuery q(m_queueDb);
        bool hasColumnName = false, hasValueText = false, hasReplayedAt = false;
        if (q.exec("PRAGMA table_info(pending_edits)")) {
            while (q.next()) {
                const QString name = q.value(1).toString();
                if (name == QLatin1String("column_name")) hasColumnName = true;
                if (name == QLatin1String("value_text"))  hasValueText  = true;
                if (name == QLatin1String("replayed_at")) hasReplayedAt = true;
            }
        }
        if (!hasColumnName) {
            QSqlQuery alter(m_queueDb);
            alter.exec("ALTER TABLE pending_edits ADD COLUMN column_name TEXT");
        }
        if (!hasValueText) {
            QSqlQuery alter(m_queueDb);
            alter.exec("ALTER TABLE pending_edits ADD COLUMN value_text TEXT");
        }
        // C5: replayed_at is a sentinel for rows whose Postgres write
        // succeeded but whose bulk DELETE from the local queue failed. We
        // mark them rather than delete so a subsequent drain doesn't
        // re-apply them (which would silently double-write the cell).
        if (!hasReplayedAt) {
            QSqlQuery alter(m_queueDb);
            alter.exec("ALTER TABLE pending_edits ADD COLUMN replayed_at TEXT");
        }
    }
    {
        // v3 Phase 3d (hazard H9): measurement edits queue with
        // schema_version=2 and a sort_order (the row ordinal half of their
        // natural key; row_id carries the SAMPLE id, column_name the metric
        // key). v1 rows leave both NULL and drain through the legacy paths -
        // see LiveSync::flushPending for the three-generation replay policy.
        QSqlQuery q(m_queueDb);
        bool hasSchemaVersion = false, hasSortOrder = false;
        if (q.exec("PRAGMA table_info(pending_edits)")) {
            while (q.next()) {
                const QString name = q.value(1).toString();
                if (name == QLatin1String("schema_version")) hasSchemaVersion = true;
                if (name == QLatin1String("sort_order"))     hasSortOrder     = true;
            }
        }
        if (!hasSchemaVersion) {
            QSqlQuery alter(m_queueDb);
            alter.exec("ALTER TABLE pending_edits ADD COLUMN schema_version INTEGER");
        }
        if (!hasSortOrder) {
            QSqlQuery alter(m_queueDb);
            alter.exec("ALTER TABLE pending_edits ADD COLUMN sort_order INTEGER");
        }
    }

    m_queueOpen = true;
    m_queueDegraded = false;   // IMPORTANT 3: a clean open clears the degraded flag
    return true;
}

bool OfflineSnapshot::enqueueCellEdit(const QString& table, qint64 rowId,
                                     const QString& column, const QVariant& value)
{
    if (!ensureQueueOpen()) return false;

    QSqlQuery q(m_queueDb);
    q.prepare("INSERT INTO pending_edits "
              "(target_table, row_id, column_name, value_text, captured_at) "
              "VALUES (?, ?, ?, ?, ?)");
    q.addBindValue(table);
    q.addBindValue(static_cast<qlonglong>(rowId));
    q.addBindValue(column);
    q.addBindValue(value.toString());
    q.addBindValue(QDateTime::currentDateTimeUtc().toString(Qt::ISODateWithMs));
    if (!q.exec()) {
        m_lastError = QStringLiteral("enqueueCellEdit: ") + q.lastError().text();
        return false;
    }
    return true;
}

// v3 Phase 3d (hazard H9): the measurement flavor - schema_version=2, keyed
// by the natural (sample id, metric key, row ordinal) identity. row_id holds
// the SAMPLE id and column_name the metric key so the existing drain SELECT,
// count and clear machinery serve both generations unchanged.
bool OfflineSnapshot::enqueueMeasurementEdit(qint64 sampleId, const QString& key,
                                             int sortOrder, const QVariant& value)
{
    if (!ensureQueueOpen()) return false;

    QSqlQuery q(m_queueDb);
    q.prepare("INSERT INTO pending_edits "
              "(target_table, row_id, column_name, value_text, captured_at, "
              " schema_version, sort_order) "
              "VALUES (?, ?, ?, ?, ?, 2, ?)");
    q.addBindValue(QStringLiteral("measurements"));
    q.addBindValue(static_cast<qlonglong>(sampleId));
    q.addBindValue(key);
    q.addBindValue(value.toString());
    q.addBindValue(QDateTime::currentDateTimeUtc().toString(Qt::ISODateWithMs));
    q.addBindValue(sortOrder);
    if (!q.exec()) {
        m_lastError = QStringLiteral("enqueueMeasurementEdit: ") + q.lastError().text();
        return false;
    }
    return true;
}

int OfflineSnapshot::drainPendingEdits(
    std::function<bool(const QString&, qint64, const QString&,
                       const QVariant&, int, int)> apply)
{
    if (!ensureQueueOpen()) return 0;

    // First, snapshot the rows into memory. Holding the SELECT cursor open
    // while the apply callback runs is fragile — the callback may end up
    // touching the same DB connection on retry paths. Read everything,
    // close the cursor, then iterate.
    //
    // v3 Phase 3d: schema_version / sort_order ride along (NULL -> 1 / -1 for
    // v1 rows) so the callback can route the three queue generations - see
    // LiveSync::flushPending.
    struct Row { qint64 id; QString table; qint64 rowId; QString column;
                 QString value; int schemaVersion; int sortOrder; };
    QVector<Row> rows;
    {
        QSqlQuery q(m_queueDb);
        // C5: filter AND replayed_at IS NULL so rows previously marked as
        // replayed (via the UPDATE fallback below) are not re-applied.
        if (!q.exec("SELECT id, target_table, row_id, column_name, value_text, "
                    "COALESCE(schema_version, 1), COALESCE(sort_order, -1) "
                    "FROM pending_edits "
                    "WHERE column_name IS NOT NULL AND replayed_at IS NULL "
                    "ORDER BY id")) {
            m_lastError = QStringLiteral("drainPendingEdits(SELECT): ")
                          + q.lastError().text();
            return 0;
        }
        while (q.next()) {
            Row r;
            r.id            = q.value(0).toLongLong();
            r.table         = q.value(1).toString();
            r.rowId         = q.value(2).toLongLong();
            r.column        = q.value(3).toString();
            r.value         = q.value(4).toString();
            r.schemaVersion = q.value(5).toInt();
            r.sortOrder     = q.value(6).toInt();
            rows.append(r);
        }
    }

    QVector<qint64> applied;
    for (const Row& r : rows) {
        if (apply(r.table, r.rowId, r.column, QVariant(r.value),
                  r.schemaVersion, r.sortOrder)) {
            applied.append(r.id);
        }
    }

    if (!applied.isEmpty()) {
        QStringList idStrs;
        idStrs.reserve(applied.size());
        for (qint64 id : applied) idStrs << QString::number(id);
        const QString idCsv = idStrs.join(QLatin1Char(','));
        QSqlQuery del(m_queueDb);
        if (!del.exec("DELETE FROM pending_edits WHERE id IN (" + idCsv + ")")) {
            // C5 fallback: the applies already landed in Postgres but the
            // local DELETE failed (lock contention, disk full, etc).
            // Rather than report 0 (which caused the caller to mistake
            // success for no-op and double-replay on the next drain), we
            // mark the rows as replayed via UPDATE and return the true
            // applied count so the queue still drains observably.
            const QString delErr = del.lastError().text();
            const QString stamp =
                QDateTime::currentDateTimeUtc().toString(Qt::ISODateWithMs);
            QSqlQuery upd(m_queueDb);
            upd.prepare("UPDATE pending_edits SET replayed_at = ? "
                        "WHERE id IN (" + idCsv + ") AND replayed_at IS NULL");
            upd.addBindValue(stamp);
            if (!upd.exec()) {
                m_lastError = QStringLiteral(
                    "drainPendingEdits(DELETE failed: %1; UPDATE replayed_at failed: %2)")
                              .arg(delErr, upd.lastError().text());
                qWarning() << "OfflineSnapshot::drainPendingEdits — both clear paths"
                              " failed for" << applied.size() << "rows:" << m_lastError;
                return 0;
            }
            m_lastError = QStringLiteral(
                "drainPendingEdits(DELETE failed, fell back to replayed_at sentinel): %1")
                          .arg(delErr);
            qWarning() << "OfflineSnapshot::drainPendingEdits — DELETE failed, marked"
                       << applied.size() << "rows replayed_at=" << stamp;
        }
    }
    return applied.size();
}

int OfflineSnapshot::pendingEditCount() const {
    if (!ensureQueueOpen()) return 0;
    QSqlQuery q(m_queueDb);
    // C5: exclude sentinel-marked rows (DELETE-fallback path) from the
    // "pending" count so the caller can trust pendingEditCount == 0 as
    // a true drain-complete signal.
    if (!q.exec("SELECT COUNT(*) FROM pending_edits "
                "WHERE column_name IS NOT NULL AND replayed_at IS NULL")) {
        m_lastError = QStringLiteral("pendingEditCount: ") + q.lastError().text();
        return 0;
    }
    return q.next() ? q.value(0).toInt() : 0;
}

// ----------------------------------------------------------------------------
// File-static helpers
// ----------------------------------------------------------------------------
namespace {

void loadImagesForSnapshot(QSqlDatabase& db,
                           const QString& tableName,
                           qint64 parentId,
                           const QString& parentCol,
                           const QString& imageCacheDir,
                           QStringList* outPaths,
                           QVector<QRectF>* outLayouts,
                           QVector<QRectF>* outCrops,
                           QString* outError)
{
    QSqlQuery qi(db);
    if (!qi.prepare(QString("SELECT file_name, image_data, layout_x, layout_y, "
                            "layout_w, layout_h, crop_x, crop_y, crop_w, crop_h "
                            "FROM %1 WHERE %2 = ? ORDER BY sort_order")
                        .arg(tableName, parentCol))) {
        if (outError) {
            *outError = QStringLiteral(
                "OfflineSnapshot::loadImagesForSnapshot(%1): prepare failed: %2")
                            .arg(tableName, qi.lastError().text());
        }
        return;
    }
    qi.addBindValue(static_cast<qlonglong>(parentId));
    if (!qi.exec()) {
        if (outError) {
            *outError = QStringLiteral(
                "OfflineSnapshot::loadImagesForSnapshot(%1): exec failed: %2")
                            .arg(tableName, qi.lastError().text());
        }
        return;
    }

    // The cache directory is a sibling of the snapshot file (so test
    // overrides propagate), not %LOCALAPPDATA% directly.
    QDir().mkpath(imageCacheDir);

    int ii = 0;
    while (qi.next()) {
        const QString  fileName = qi.value(0).toString();
        const QByteArray blob   = qi.value(1).toByteArray();
        const QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                            qi.value(4).toDouble(), qi.value(5).toDouble());
        const QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                          qi.value(8).toDouble(), qi.value(9).toDouble());

        const QString tempPath = imageCacheDir + "/snap_" + tableName + "_" +
                                 QString::number(parentId) + "_" +
                                 QString::number(ii++) + "_" + fileName;
        QFile tmpFile(tempPath);
        if (tmpFile.open(QIODevice::WriteOnly)) {
            tmpFile.write(blob);
            tmpFile.close();
        }
        outPaths->append(tempPath);
        outLayouts->append(layout);
        outCrops->append(crop);
    }
}

// Thin wrapper around the canonical pipeline-layer decoder so the local
// snapshot read path stays byte-compatible with the JSONB column written
// by DatabaseManager and the .json files written by SensoryPanel.
bool deserializeSensoryJsonLocal(const QByteArray& bytes, SensorySession& sess)
{
    const QJsonDocument doc = QJsonDocument::fromJson(bytes);
    if (doc.isNull() || !doc.isObject()) return false;
    sess = sensorySessionFromJson(doc.object());
    return true;
}

bool deserializeDetailedSensoryJsonLocal(const QByteArray& bytes,
                                         DetailedSensorySession& sess)
{
    const QJsonDocument doc = QJsonDocument::fromJson(bytes);
    if (doc.isNull() || !doc.isObject()) return false;

    const QJsonObject root = doc.object();
    sess.sessionName        = root["session_name"].toString();
    sess.testTitle          = root["test_title"].toString();
    sess.assessorName       = root["assessor_name"].toString();
    sess.testerName         = root["tester_name"].toString();
    sess.facilitatorName    = root["facilitator_name"].toString();
    sess.facilitatorComment = root["facilitator_comment"].toString();
    sess.media              = root["media"].toString();
    sess.date               = root["date"].toString();
    sess.timestamp          = root["timestamp"].toString();
    sess.oilSmellLiking     = root["oil_smell_liking"].toInt(3);
    sess.clog               = root["clog"].toBool(false);
    sess.clogOilLevel       = root["clog_oil_level"].toString();
    sess.mouthpieceNotes    = root["mouthpiece_notes"].toString();
    sess.deviceReturnDate   = root["device_return_date"].toString();
    sess.viscosity          = root["viscosity"].toString();

    for (const QJsonValue& sv : root["samples"].toArray()) {
        const QJsonObject sObj = sv.toObject();
        DetailedSensorySample sample;
        sample.name              = sObj["name"].toString();
        sample.comments          = sObj["comments"].toString();
        sample.voltage           = sObj["voltage"].toDouble();
        sample.resistance        = sObj["resistance"].toDouble();
        sample.power             = sObj["power"].toDouble();
        sample.heatingTechnology = sObj["heating_technology"].toString();
        for (const QString& metric : kDetailedAllMetrics) {
            const double maxVal = kDetailedMetricMaxScore.value(metric, 9);
            sample.scores[metric] = qBound(1.0, sObj[metric].toDouble(1.0), maxVal);
        }
        sess.samples.append(sample);
    }
    return true;
}

} // anonymous

} // namespace DVE

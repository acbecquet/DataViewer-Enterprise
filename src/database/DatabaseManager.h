#pragma once

#include <QObject>
#include <QString>
#include <QStringList>
#include <QVector>
#include <QVariantMap>
#include <QSqlDatabase>
#include <functional>
#include "ConfigLoader.h"
#include "../pipeline/ReportData.h"
#include "../pipeline/SensoryData.h"
#include "../pipeline/DetailedSensoryData.h"

namespace DVE {

class PostgresConnection;
class IdentityManager;
class OfflineSnapshot;

// ── WriteResult — outcome of a save method that participates in optimistic
//    concurrency control. Defined ahead of class DatabaseManager so callers
//    that include this header can switch on the enum.
//
//    Naming convention: the primary "save with rich result" API is
//    `tryWriteXxx(...) -> WriteResult`. The historical `saveXxx(...) -> bool`
//    shims forward to tryWriteXxx and return `result == Success`. v2.0.1
//    routes per-cell edits through LiveSync (which calls tryWrite* directly),
//    and the bool shims are still used for bulk/manual saves + offline replay.
//
//    `OfflineReadOnly` is a forward declaration for Plan C — no method
//    currently returns it. It exists in the enum now so Plan C doesn't
//    have to re-ship this header.
enum class WriteResult {
    Success,
    VersionMismatch,    // expected version stale; server row has newer version
    RowDeleted,         // row no longer exists at this id
    UniqueViolation,    // hit a unique constraint (e.g., duplicate file_path)
    OfflineReadOnly,    // reserved for Plan C; saves blocked while offline
    OtherError          // generic SQL failure; details in lastError()
};

struct FileRecord {
    int    id;
    QString filePath;
    QString fileName;
    QString loadedAt;
    QString templateVersion;
    int     sheetCount;
    int     sampleCount;
    // F6 (v2.5.0): the add-time stamp that, together with file_path, identifies
    // a version. The DB browser appends it to the file name when a path has more
    // than one version, so historical re-adds are distinguishable. Empty on
    // legacy rows whose added_at backfilled NULL.
    QString addedAt;
    // SP4 A4: server-side era stamp (NULL/empty on pre-v2.4.2 rows).
    QString appVersion;
};

struct SensoryRecord {
    int     id;
    QString sessionName;
    QString testTitle;
    QString assessorName;
    QString testerName;
    QString media;
    QString date;
    int     sampleCount;
    // SP4 A4: era stamp + a server-derived health signal. hasLegacyStringScores
    // is true when any score key still holds a string-typed value (old-version
    // write); it is computed server-side (the C++ reader coerces strings to
    // doubles, so the distinction is invisible here) and is false on the offline
    // snapshot path (SQLite has no jsonb_typeof).
    QString appVersion;
    bool    hasLegacyStringScores = false;
};

struct DetailedSensoryRecord {
    int     id;
    QString sessionName;
    QString testTitle;
    QString assessorName;
    QString testerName;
    QString media;
    QString date;
    int     sampleCount;
    QString appVersion;
    bool    hasLegacyStringScores = false;
};

class DatabaseManager : public QObject {
    Q_OBJECT

public:
    explicit DatabaseManager(QObject* parent = nullptr);
    ~DatabaseManager() override;

    bool open(const DbConfig& cfg, IdentityManager* identity);
    // v2.1.0: tear down the underlying QSqlDatabase and reopen with the
    // cfg captured during the last successful open(). Used by MainWindow
    // when ConnectionMonitor reports the network is back — Qt's QSqlDatabase
    // doesn't poll, so a transient outage leaves the driver-level connection
    // in a half-dead state that fails the next query with "server closed the
    // connection unexpectedly". Returns false (and leaves m_online=false) if
    // the new open() fails or if open() was never called successfully.
    bool reopen();
    void close();
    bool isOpen() const;

    // Additive, idempotent schema reconciliation. Brings a live database that
    // was created from an OLDER init.sql up to the current column set by adding
    // any post-baseline additive column that's missing. Run automatically on
    // every successful open()/reopen(); also safe to call directly (e.g. a
    // future "repair" action). Each column is catalog-checked first so the
    // common case (already present) never takes the brief ACCESS EXCLUSIVE lock
    // an ALTER TABLE would — important on a live multi-user DB. Never drops or
    // renames anything, so it cannot corrupt or lose data. Keep the column list
    // in sync with deploy/postgres/migrations/*.sql (ADD COLUMN) + init.sql.
    void ensureSchema();

    // v3 Phase 3d (index D-3d-2): connect-time gate - data_rows must be the
    // post-cutover name-holder VIEW. Sets m_lastError and returns false
    // against a pre-cutover database; open()/reopen() then refuse.
    bool verifyCutoverSchema();

    QString currentPath() const;

    // ── Hierarchical file storage ────────────────────────────────────────────
    // Stores the full file hierarchy (tests, samples, data rows, images).
    //
    // Optimistic concurrency: if result.id != -1 and result.version > 0,
    // the file row is UPDATEd with WHERE id = ? AND version = ?. A zero
    // rowcount triggers a follow-up SELECT to distinguish VersionMismatch
    // (row exists, version stale) from RowDeleted (row no longer there).
    // INSERT path returns UniqueViolation on SQLSTATE 23505 (duplicate
    // file_path). On Success, children (tests/samples/data_rows/images)
    // are wiped via DELETE WHERE file_id = ? and re-inserted.
    WriteResult tryWriteFile(const FileResult& result);
    // Mutable overload — identical to the const-ref version except that on
    // WriteResult::Success it writes the post-save id + version back into
    // `result`. Callers that need to round-trip the file across multiple
    // saves (e.g., MainWindow's recreate handler after a RowDeleted) MUST
    // use this overload; the const-ref version above is fire-and-forget.
    WriteResult tryWriteFile(FileResult& result);
    // Bool shim - returns true iff tryWriteFile returned Success.
    //
    // 3a/H7: NO production caller remains. It used to be the DB-browser load's
    // save (MainWindow's onOpenDatabaseBrowser), which is the last call the
    // H7 fix moved to the mutable-ref overload so the id/version writeback
    // lands; the const-ref tryWriteFile above is in turn reachable only
    // through this shim. What keeps both alive is tst_databasemanager, which
    // drives ~25 call sites through them plus one save-integrity scenario -
    // churning those is not worth it here. Retire the pair in Phase 3d, when
    // the write path is being touched anyway.
    bool saveFile(const FileResult& result);

    // Quick existence check
    bool hasFile(const QString& filePath) const;

    // Load a full FileResult from the database (inverse of saveFile).
    // Returns an empty FileResult (filePath.isEmpty()) on failure.
    FileResult loadFile(int id) const;
    FileResult loadFileByPath(const QString& filePath) const;

    // ── File listing / removal ───────────────────────────────────────────────
    QVector<FileRecord> listFiles() const;
    bool removeFile(int id);

    // Remove duplicates: keep only the N most recent entries per file_name.
    // Also deletes entries with "unknown" template version (corrupt/empty).
    int deduplicateFiles(int keepPerName = 3);

    // SP4 A5: run the server-side legacy-score normalizer NOW — the same
    // idempotent, lossless string->numeric coercion the nightly cron performs
    // (dve_normalize_legacy_json), invoked from the Database Browser's "Repair
    // Legacy Scores" button. Acts on EVERY damaged row in both session tables,
    // not a selection. Returns the number of rows rewritten, or -1 on error
    // (offline, or SQL failure — details in lastError()).
    int normalizeLegacyScores();

    // Recent files (last 20)
    QStringList recentFilePaths() const;

    // ── Sensory sessions ─────────────────────────────────────────────────────
    //
    // Optimistic concurrency: identical scheme to tryWriteFile. The by-ref
    // overload also writes the post-save id and version back into the
    // struct so the next save round-trips them.
    WriteResult tryWriteSensorySession(const SensorySession& s);
    WriteResult tryWriteSensorySession(SensorySession& s);
    // v2.5.0 RC4 — collision-tolerant wrapper over tryWriteSensorySession(byRef).
    // On WriteResult::UniqueViolation it bumps BOTH s.sessionName AND s.testTitle
    // via OutputPaths::nextSuffixedName (kept in lockstep because buildSession()
    // regenerates sessionName FROM testTitle — suffixing only sessionName would
    // regress on the next save) and retries, up to maxAttempts. On Success the
    // byRef back-fill populates s.id/s.version AND this wrapper sets
    // s.originalSessionName = s.sessionName, so the panel's rename detector can
    // never re-detect the resolved name as a fresh rename and re-collide (the
    // June-10 endless-loop kill). For a non-collision result the behavior is
    // identical to tryWriteSensorySession — it is a strict superset, safe to
    // substitute everywhere the plain call is made. Returns UniqueViolation only
    // if maxAttempts is exhausted; any other non-Success result is returned as-is.
    WriteResult tryWriteSensorySessionAutoSuffix(SensorySession& s, int maxAttempts = 25);
    // Bool shims — return true iff Success. The by-ref overload still
    // populates s.id on Success (and now s.version too).
    bool saveSensorySession(const SensorySession& s);
    // Overload that populates s.id with the autoincrement id assigned by
    // Postgres on insert, plus s.version with the server-assigned version.
    // Post-condition: on true return, s.id is normally > 0 — but if a
    // concurrent writer deletes the row between the INSERT and the
    // lookup-by-natural-key, s.id stays at its prior value (-1 for a
    // fresh struct). Callers that depend on a valid id should verify
    // s.id > 0.
    bool saveSensorySession(SensorySession& s);
    QVector<SensorySession> loadSensorySessions() const;
    SensorySession loadSensorySession(int id) const;
    QVector<SensoryRecord> listSensoryRecords() const;
    bool removeSensorySession(int id);
    QString nextDefaultTestName() const;

    // ── Detailed sensory sessions ───────────────────────────────────────────
    WriteResult tryWriteDetailedSensorySession(const DetailedSensorySession& s);
    // Mutable overload — symmetric to tryWriteFile(FileResult&) and
    // tryWriteSensorySession(SensorySession&). On WriteResult::Success it writes
    // the post-save id + version back into `s`. Required for any caller that
    // round-trips the struct across saves (recreate flow, etc.).
    WriteResult tryWriteDetailedSensorySession(DetailedSensorySession& s);
    // v2.5.0 RC4 — detailed twin of tryWriteSensorySessionAutoSuffix. Suffixes
    // s.sessionName + s.testTitle in lockstep on UniqueViolation and retries.
    // DetailedSensorySession has no originalSessionName (no in-place rename
    // branch), so there is nothing to re-baseline here. Strict superset of
    // tryWriteDetailedSensorySession for non-collision results.
    WriteResult tryWriteDetailedSensorySessionAutoSuffix(DetailedSensorySession& s,
                                                         int maxAttempts = 25);
    bool saveDetailedSensorySession(const DetailedSensorySession& s);
    QVector<DetailedSensorySession> loadDetailedSensorySessions() const;
    DetailedSensorySession loadDetailedSensorySession(int id) const;
    QVector<DetailedSensoryRecord> listDetailedSensoryRecords() const;
    bool removeDetailedSensorySession(int id);

    // ── Layout JSON persistence (sensory report preview) ─────────────────────
    QString loadSensoryLayout(int sessionId) const;
    bool    saveSensoryLayout(int sessionId, const QString& layoutJson);

    QString loadCumulativeLayout() const;
    bool    saveCumulativeLayout(const QString& layoutJson);

    // ── Natural-key lookup for imported sessions (v2.0.6) ────────────────────
    // Returns (id, version) of sessions matching the natural keys
    // (session_name, tester_name, date). Misses get (-1, 0). Used by
    // SensoryPanel::inheritExistingIdsAndVersions to convert a fresh-import
    // INSERT into an in-place UPDATE so re-importing the same Excel/JSON
    // file doesn't UNIQUE-violate against the prior import's DB row.
    //
    // One round-trip regardless of session count (chunked internally at 200
    // keys to stay well below libpq's parameter ceiling). Each input key is
    // matched literally — caller is responsible for trimming whitespace.
    struct SessionKey {
        qint64 id      = -1;
        int    version = 0;
    };
    struct NaturalKey {
        QString sessionName;
        QString testerName;
        QString date;
    };
    struct SessionKeyMatch {
        QString sessionName;
        QString testerName;
        QString date;
        qint64  id      = -1;
        int     version = 0;
    };
    QVector<SessionKeyMatch>
        findSensorySessionsByKeys(const QVector<NaturalKey>& keys) const;
    QVector<SessionKeyMatch>
        findDetailedSensorySessionsByKeys(const QVector<NaturalKey>& keys) const;

    // ── Sensory header presets (v2.0.4 QoL) ──────────────────────────────────
    // Saves the current Test Title / Media / per-sample names into a shared
    // pool so other users can pick them from a dropdown instead of retyping.
    // Idempotent: re-saving the same trio is a no-op (INSERT ... ON CONFLICT
    // DO NOTHING on the (kind, value) unique constraint). Empty / whitespace
    // values are silently skipped — the table-level CHECK constraint would
    // otherwise reject them and abort the batch.
    bool saveSensoryHeaderPresets(const QString& testName,
                                  const QString& media,
                                  const QStringList& sampleNames);

    // Returns the alphabetised list of saved preset values for the given
    // kind ('test_name' | 'media' | 'sample_name'). Returns empty if the
    // table doesn't exist yet (pre-migration installs) or the kind is
    // unknown — callers should fall back to a plain text edit on empty.
    QStringList loadSensoryHeaderPresets(const QString& kind) const;

    // DATAVIEWER-2: sample-name presets are scoped per test. Returns the
    // alphabetised sample names saved under one test_name (the sample_name
    // rows whose test_name column equals testName). test_name/media presets
    // remain global and are still read via loadSensoryHeaderPresets() above.
    QStringList loadSampleNamesForTest(const QString& testName) const;

    // SP4.5 v2.4.8: remove one saved preset from the shared pool (the per-entry
    // ✕ in the dropdown). kind is 'test_name'|'media'|'sample_name'. testName
    // scopes sample_name deletes to one test; pass empty for the global kinds
    // (test_name/media, whose rows carry a NULL test_name). Like the save, this
    // commits on its OWN isolated connection so it can't be dropped by the
    // shared connection's transaction state.
    bool deleteSensoryHeaderPreset(const QString& kind, const QString& value,
                                   const QString& testName = QString());

    // ── Settings key/value store ─────────────────────────────────────────────
    bool setSetting(const QString& key, const QString& value);
    QString getSetting(const QString& key, const QString& defaultVal = "") const;

    // ── Offline mode (Plan C) ──────────────────────────────────────────────
    //
    // DatabaseManager owns a soft "online" flag separate from the underlying
    // PostgresConnection's hard isOpen() state. ConnectionMonitor (C3) toggles
    // this when ping detects a disconnect / reconnect. Independent because the
    // PG connection may still be technically open (Qt's QSqlDatabase doesn't
    // poll) while the server is unreachable.
    //
    // While offline:
    //   - All tryWrite*/save*/remove*/dedup*/setSetting/saveLayout methods
    //     refuse with WriteResult::OfflineReadOnly (or false for bool shims)
    //     and set m_lastError to "DatabaseManager is offline (read-only mode)".
    //   - All read methods (listFiles, loadFile, loadFileByPath, loadSensory*,
    //     listSensory*, getSetting, recentFilePaths, loadSensoryLayout,
    //     loadCumulativeLayout) route to the OfflineSnapshot when one is set
    //     AND open. Without a snapshot, reads return empty/default values
    //     and set m_lastError to
    //     "DatabaseManager is offline and no snapshot is set".
    //
    // Default state after open() succeeds: m_online = true, m_snapshot = nullptr.
    void setOfflineSnapshot(OfflineSnapshot* snap);
    bool isOnline() const { return m_online; }
    void setOnline(bool b);

    QString lastError() const { return m_lastError; }

    // Test-only (v2.4.8): drive an ambient transaction on the SHARED connection
    // so a test can prove QoL writes (header presets) persist on their own
    // isolated connection even when the shared connection later rolls back —
    // the exact production failure where a preset INSERT rode an uncommitted
    // transaction and was dropped on reconnect. No production callers.
    bool beginSharedTransactionForTesting();
    bool rollbackSharedTransactionForTesting();

private:
    PostgresConnection* m_pg = nullptr;
    IdentityManager*    m_identity = nullptr;
    // v2.1.0: remember the cfg from the last successful open() so reopen()
    // can rebuild the QSqlDatabase without having to plumb cfg back through
    // MainWindow on every reconnect. Empty DbConfig = open() never succeeded.
    DbConfig            m_cfg;
    bool                m_haveCfg = false;
    // Read methods (loadFile, listFiles, getSetting, etc.) are const but
    // still need to clear/set m_lastError on entry/error. mutable lets them
    // do that without breaking the const contract; m_lastError is purely a
    // diagnostic side-channel.
    mutable QString     m_lastError;
    bool                m_open = false;
    OfflineSnapshot*    m_snapshot = nullptr;  // not owned; lifetime managed by MainWindow
    bool                m_online   = false;    // set to true on successful open()

    void logDebug(const QString& msg) const;

    // SP4.5 v2.4.8: run `body` on a short-lived ISOLATED QPSQL connection opened
    // from m_cfg (fresh = autocommit), so QoL writes (header preset save/delete)
    // commit durably and immediately, independent of the shared connection's
    // transaction/reconnect state. Returns body's result; false (m_lastError set)
    // if the isolated connection can't be opened. Mirrors PersistWorker's connect.
    bool runIsolatedWrite(const std::function<bool(QSqlDatabase&)>& body) const;
};

} // namespace DVE

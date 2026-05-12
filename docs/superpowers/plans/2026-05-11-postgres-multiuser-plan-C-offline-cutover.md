# Postgres Multi-User — Plan C: Offline Mode + Final Cutover Implementation Plan

> **For agentic workers:** Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax. See [INDEX](2026-05-11-postgres-multiuser-INDEX.md).

**Goal:** Add the hybrid online/offline behavior — when the Synology is unreachable, the app falls into read-only mode against a local SQLite snapshot with a polite banner. Reconnect detection with jittered retry. In-flight-edit pending badge. Delete any remaining `<dbPath>.lock` code that Plan B didn't already remove. Final v2.0 release.

**Architecture:** New `OfflineSnapshot` class (local SQLite mirror at `%LOCALAPPDATA%\DataViewer\snapshot.sqlite`, atomic write on clean app close). New `ConnectionMonitor` that owns the 30s ping + jittered retry. Banner widget at the top of MainWindow. `DatabaseManager` gains a fallback path to OfflineSnapshot when PostgresConnection is unhealthy.

**Spec:** [docs/superpowers/specs/2026-05-11-postgres-multiuser-design.md](../specs/2026-05-11-postgres-multiuser-design.md) — see "Online / offline flow" section.

---

## Tasks

### Phase 1 — OfflineSnapshot (Tasks 1–3)

#### Task 1: `OfflineSnapshot.h`/`.cpp`

**Files (new, Python pattern):**
- `src/database/OfflineSnapshot.h`
- `src/database/OfflineSnapshot.cpp`

**Header:**

```cpp
#pragma once

#include <QObject>
#include <QString>
#include <QSqlDatabase>
#include <QDateTime>
#include "../pipeline/ReportData.h"
#include "../pipeline/SensoryData.h"
#include "../pipeline/DetailedSensoryData.h"

namespace DVE {

class PostgresConnection;

class OfflineSnapshot : public QObject {
    Q_OBJECT
public:
    explicit OfflineSnapshot(QObject* parent = nullptr);
    ~OfflineSnapshot() override;

    // Path: %LOCALAPPDATA%\DataViewer\snapshot.sqlite (resolves on first use).
    QString path() const;

    // Generates a snapshot by reading from the live Postgres connection
    // and writing into a new SQLite file. Atomic: writes to .tmp, then
    // renames over the production path (MoveFileEx REPLACE_EXISTING).
    // Returns false on any error; the previous snapshot is preserved.
    bool regenerate(PostgresConnection* live);

    // Opens the snapshot read-only. Subsequent loadFile/listFiles/... go
    // through this connection.
    bool openReadOnly();
    void close();
    bool isOpen() const { return m_open; }

    // Read-only accessors (mirror DatabaseManager's read methods).
    QVector<FileRecord>           listFiles() const;
    FileResult                    loadFileByPath(const QString& filePath) const;
    QVector<SensoryRecord>        listSensoryRecords() const;
    SensorySession                loadSensorySession(int id) const;
    QVector<DetailedSensoryRecord> listDetailedSensoryRecords() const;
    DetailedSensorySession        loadDetailedSensorySession(int id) const;

    // Provenance: when was this snapshot last regenerated?
    QDateTime snapshotTakenAt() const;

private:
    QSqlDatabase m_db;
    QString      m_path;
    QString      m_connName;
    bool         m_open = false;
};

} // namespace DVE
```

**.cpp:** create the SQLite snapshot DB with the same schema as Plan A's `init.sql` (TPM hierarchy + sensory + detailed sensory + settings; presence + schema_meta NOT mirrored), populate by streaming rows from the live PG connection in dependency order. Use BLOB binding for images. Track `snapshot_taken_at` in a small `_snapshot_meta` SQLite table.

Atomic write:
```cpp
const QString tmpPath = m_path + ".tmp";
// ... write to tmpPath ...
QFile::remove(m_path);  // Windows MoveFile won't overwrite an existing file without this
if (!QFile::rename(tmpPath, m_path)) return false;
```

Or use `MoveFileExW(... MOVEFILE_REPLACE_EXISTING)` via Win32 for true atomicity on Windows.

Register in `.pro`. Commit: `feat(db): OfflineSnapshot with atomic regenerate + read-only access`.

#### Task 2: `tst_offlinesnapshot`

- Regenerate from a populated Postgres → snapshot file exists with right row counts.
- Atomic write: corrupt the .tmp midway → original snapshot preserved.
- `openReadOnly` returns SELECTable results; INSERT errors out (sqlite read-only flag).
- `snapshotTakenAt` returns sensible timestamp.

Commit: `test(db): tst_offlinesnapshot regenerate + atomic + read-only`.

#### Task 3: `DatabaseManager` integration with snapshot

`DatabaseManager` adds:
```cpp
void setOfflineSnapshot(OfflineSnapshot* snap);
bool isOnline() const;
void setOnline(bool b);  // called by ConnectionMonitor
```

Read methods (`listFiles`, `loadFileByPath`, etc.) route to PG when online, snapshot when offline. Write methods refuse with a specific error code when offline:

```cpp
WriteResult DatabaseManager::saveFile(const FileResult& f) {
    if (!m_online) return WriteResult::OfflineReadOnly;
    // ... existing PG path ...
}
```

`WriteResult` enum gains `OfflineReadOnly`. MainWindow translates this to a status bar message ("Working offline — cannot save").

Commit: `feat(db): DatabaseManager online/offline routing`.

### Phase 2 — ConnectionMonitor (Tasks 4–5)

#### Task 4: `ConnectionMonitor.h`/`.cpp`

```cpp
namespace DVE {
class ConnectionMonitor : public QObject {
    Q_OBJECT
public:
    ConnectionMonitor(PostgresConnection* conn, const DbConfig& cfg,
                      QObject* parent = nullptr);

    void start();
    void stop();
    bool isOnline() const { return m_online; }

signals:
    void wentOffline();
    void cameOnline();

private slots:
    void onPing();
    void onReconnectAttempt();

private:
    PostgresConnection* m_conn;
    DbConfig            m_cfg;
    QTimer              m_pingTimer;       // 30s while online
    QTimer              m_reconnectTimer;  // jittered while offline
    bool                m_online = true;
};
}
```

`start()`: launches the 30s ping while online. On ping fail → `wentOffline()` signal, switches timers. While offline, every ~30s + jitter, attempt reconnect; on success → `cameOnline()`.

Disconnect detection signals (spec section 3):
1. NOTIFY socket dies (driver loses connection).
2. Query returns SQLSTATE 08006 / 57P01 (connection-level errors).
3. 30s `SELECT 1` ping fails.

ConnectionMonitor handles #3; QSqlDriver signal `notification` going silent + (1) is monitored by NotificationListener.

Commit: `feat(db): ConnectionMonitor with 30s ping + jittered reconnect`.

#### Task 5: tst_connectionmonitor

Spin Postgres up, ConnectionMonitor reports online. Stop container, ping fails within 30s, `wentOffline` fires. Start container, reconnect within ~30s, `cameOnline` fires. Use jitter cap so test completes in <60s.

Commit: `test(db): tst_connectionmonitor offline/online transitions`.

### Phase 3 — Offline UI (Tasks 6–8)

#### Task 6: `OfflineBanner` widget

A QFrame at the top of MainWindow's central area:

```
┌──────────────────────────────────────────────────────────────────────┐
│ ⚠  Working offline — read-only. Last synced 2 hours ago. [Retry]    │
└──────────────────────────────────────────────────────────────────────┘
```

Visible only when `m_db->isOnline() == false`. Styled with `background: #fef3c7; border-bottom: 1px solid #f59e0b;`.

Wire `cameOnline` → hide banner (with a brief "Reconnected" toast). `wentOffline` → show banner.

Commit: `feat(ui): OfflineBanner top-of-window`.

#### Task 7: Pending-edit badge

If a save attempt while offline is queued in memory (per spec: "edits stay in memory, nothing silently discarded"), show a small yellow dot on the affected cell. Banner gains second line: "1 unsaved change. Will retry when reconnected."

When ConnectionMonitor signals `cameOnline`, flush the pending edits via optimistic concurrency. If a queued edit conflicts, surface ConflictResolver dialog.

Commit: `feat(ui): in-flight-edit pending badge + retry on reconnect`.

#### Task 8: Wire ConnectionMonitor into MainWindow

```cpp
m_monitor = new ConnectionMonitor(m_pgConn, cfg, this);
connect(m_monitor, &ConnectionMonitor::wentOffline, this, [this]() {
    m_db->setOnline(false);
    m_offlineBanner->setVisible(true);
});
connect(m_monitor, &ConnectionMonitor::cameOnline, this, [this]() {
    m_db->setOnline(true);
    m_offlineBanner->setVisible(false);
    // ... show "Reconnected" toast, retry pending edits ...
});
m_monitor->start();
```

Also: on app startup, if PG connect fails AND snapshot exists, fall back to snapshot read-only mode. If PG connect fails AND no snapshot, show modal: "Cannot reach database and no offline copy available."

Commit: `feat(ui): MainWindow wires ConnectionMonitor + offline modal fallback`.

### Phase 4 — Snapshot lifecycle (Task 9)

#### Task 9: Refresh snapshot on clean close + manual refresh menu item

```cpp
// In MainWindow destructor / closeEvent:
if (m_db->isOnline() && m_snapshot) {
    m_snapshot->regenerate(m_pgConn);
}
```

Also add `File → Refresh Offline Snapshot` menu item that calls `regenerate` synchronously with a progress dialog.

Commit: `feat(ui): snapshot regenerate on clean close + manual menu item`.

### Phase 5 — Final cutover (Tasks 10–12)

#### Task 10: Delete all remaining lock-file code

Search `src/` for any remaining references to `LockInfo`, `forceReleaseLock`, `writeLockFile`, `pathLooksCloudSynced`, `m_lockPath`, `m_lockInfo`, `acquireLock`. Remove them. Plan B's DatabaseManager rewrite likely got most; this is the cleanup sweep.

Also remove the "Force Release Lock" menu item from MainWindow if it still exists.

Commit: `chore(db): delete remaining lock-file infrastructure`.

#### Task 11: Deployment self-test additions

`tests/deployment/Test-Deployment.ps1` gains offline scenario coverage in `SelfTest`:
- `testOfflineSnapshotExists` (check `%LOCALAPPDATA%\DataViewer\snapshot.sqlite`)
- `testOfflineFallback` (stop PG mid-`--self-test`, verify fallback path)

`src/utils/SelfTest.cpp` adds `testOfflineSnapshot()` and `testReconnectDetection()` cases.

Commit: `feat(selftest): offline snapshot + reconnect diagnostics`.

#### Task 12: CLAUDE.md final update

Replace any remaining transitional language ("being phased out", "scheduled across three plans") with definitive statements ("uses PostgreSQL", "lock-file code is gone").

Commit: `docs(claude): finalize Postgres-only language after v2 ship`.

### Phase 6 — Release (Tasks 13–15)

#### Task 13: Plan C checkpoint test

Full test sweep including:
- All Plan A/B test suites
- Plan C: tst_offlinesnapshot, tst_connectionmonitor
- E2E: simulate NAS down → banner appears → resume → banner clears → pending edit flushes

Commit: `test: Plan C end-to-end checkpoint`.

#### Task 14: Bump version to v2.0.0

`tools/version.txt` or wherever the version lives. Update installer wizard title. Commit: `chore(release): bump version to 2.0.0`.

#### Task 15: Plan C closeout

Update INDEX status (Plans A/B/C all "Complete and verified"). Tag `v2.0.0-rc.1` for the work machine to build the installer. Commit: `docs(plans): mark Plan C complete`.

---

## Plan C Checkpoint Criteria

1. Pull cable / stop PG container → banner appears within 30s with reasonable copy.
2. Reconnect → "Reconnected" toast, banner clears, presence reappears.
3. While offline: file list still browseable from snapshot, save attempts politely refused.
4. Mid-edit when offline: edit retained, "pending" badge shown. On reconnect → either save succeeds or ConflictResolver dialog fires.
5. `<dbPath>.lock` code is GONE — `grep -r LockInfo src/` and `grep -r writeLockFile src/` find nothing.
6. Deployment self-test reports all diagnostics passing including new offline cases.
7. `Test-Deployment.ps1` end-to-end pass with `-PreMigrationSqlite` providing real migrated data.
8. Manual two-machine test: simulate NAS reboot, verify both clients ride through gracefully.

After this checkpoint, the Postgres multi-user database initiative is complete. Tag v2.0.0, build installer, ship to Synology release folder, run real-world install on a couple workstations.

# Postgres Multi-User — Plan B: Runtime Cutover Implementation Plan

> **For agentic workers:** Use superpowers:subagent-driven-development. Steps use checkbox (`- [ ]`) syntax. See [INDEX](2026-05-11-postgres-multiuser-INDEX.md) for context.

**Goal:** Switch the DataViewer runtime from SQLite to PostgreSQL using the infrastructure Plan A delivered. Wire up live NOTIFY-driven UI updates, presence indicators, optimistic concurrency conflict resolution, and the "don't yank in-progress edits" rule. Keep SQLite as the offline-snapshot mechanism (Plan C handles that).

**Architecture:** New components in `src/database/` (`NotificationListener`, `PresenceManager`, `ConflictResolver`) plus three conflict dialog classes. `DatabaseManager` rewired to use `PostgresConnection` under the hood — public signatures stay the same. Identity (Plan A's `IdentityManager`) supplies `updated_by` on every write. `MainWindow` wires everything together.

**Spec:** [docs/superpowers/specs/2026-05-11-postgres-multiuser-design.md](../specs/2026-05-11-postgres-multiuser-design.md)

**Cross-cutting:** MIP encryption mitigation (new `.cpp`/`.h` files via Python pattern). Qt 6.10.1 at home / 6.10.2 at work. `vendor/libpq-16` on PATH for tests. `DVE_TEST_PG_CONN` env var for ephemeral Postgres (use `tests/start-test-postgres.ps1`).

---

## Tasks

### Phase 1 — NotificationListener (Tasks 1–2)

#### Task 1: `NotificationListener.h`/`.cpp`

**Files (new, Python pattern):**
- `src/database/NotificationListener.h`
- `src/database/NotificationListener.cpp`

**Header:**

```cpp
#pragma once

#include <QObject>
#include <QString>
#include <QUuid>

class QSqlDriver;

namespace DVE {

class PostgresConnection;

struct RowChange {
    QString table;
    QString op;          // "INSERT" | "UPDATE" | "DELETE"
    qint64  id   = -1;
    QString updatedBy;   // UUID string (matches IdentityManager::uuid().toString(WithoutBraces))
};

struct PresenceChange {
    QString op;            // "INSERT" | "UPDATE" | "DELETE"
    QUuid   userUuid;
    QString resourceType;  // "file" | "sensory_session" | "detailed_sensory_session"
    qint64  resourceId = -1;
    QString intent;        // "viewing" | "editing"
};

class NotificationListener : public QObject {
    Q_OBJECT
public:
    explicit NotificationListener(PostgresConnection* conn, QObject* parent = nullptr);
    ~NotificationListener() override;

    bool subscribe();         // subscribes to both channels
    void unsubscribe();

signals:
    void rowChanged(const DVE::RowChange& change);
    void presenceChanged(const DVE::PresenceChange& change);

private slots:
    void onNotification(const QString& name, QSqlDriver::NotificationSource source,
                        const QVariant& payload);

private:
    PostgresConnection* m_conn;
    bool                m_subscribed = false;
};

} // namespace DVE
```

**.cpp body:**

```cpp
#include "NotificationListener.h"
#include "PostgresConnection.h"

#include <QSqlDriver>
#include <QJsonDocument>
#include <QJsonObject>

namespace DVE {

NotificationListener::NotificationListener(PostgresConnection* conn, QObject* parent)
    : QObject(parent), m_conn(conn) {}

NotificationListener::~NotificationListener() { unsubscribe(); }

bool NotificationListener::subscribe() {
    if (!m_conn || !m_conn->isOpen()) return false;
    QSqlDriver* drv = m_conn->listenDb().driver();
    if (!drv) return false;
    bool ok = drv->subscribeToNotification("dataviewer_changes");
    ok = drv->subscribeToNotification("dataviewer_presence") && ok;
    if (ok) {
        connect(drv, &QSqlDriver::notification,
                this, &NotificationListener::onNotification);
        m_subscribed = true;
    }
    return ok;
}

void NotificationListener::unsubscribe() {
    if (!m_subscribed || !m_conn || !m_conn->isOpen()) return;
    QSqlDriver* drv = m_conn->listenDb().driver();
    if (drv) {
        drv->unsubscribeFromNotification("dataviewer_changes");
        drv->unsubscribeFromNotification("dataviewer_presence");
        disconnect(drv, &QSqlDriver::notification, this, nullptr);
    }
    m_subscribed = false;
}

void NotificationListener::onNotification(const QString& name,
                                          QSqlDriver::NotificationSource,
                                          const QVariant& payload) {
    const QJsonDocument doc = QJsonDocument::fromJson(payload.toString().toUtf8());
    if (!doc.isObject()) return;
    const QJsonObject o = doc.object();

    if (name == "dataviewer_changes") {
        RowChange c;
        c.table     = o.value("table").toString();
        c.op        = o.value("op").toString();
        c.id        = o.value("id").toVariant().toLongLong();
        c.updatedBy = o.value("updated_by").toString();
        emit rowChanged(c);
    } else if (name == "dataviewer_presence") {
        PresenceChange p;
        p.op           = o.value("op").toString();
        p.userUuid     = QUuid(o.value("user_uuid").toString());
        p.resourceType = o.value("resource_type").toString();
        p.resourceId   = o.value("resource_id").toVariant().toLongLong();
        p.intent       = o.value("intent").toString();
        emit presenceChanged(p);
    }
}

} // namespace DVE
```

Register in `DataViewerEnterprise.pro`. Commit: `feat(db): NotificationListener for dataviewer_changes + dataviewer_presence channels`.

#### Task 2: `tst_notificationlistener` end-to-end test

**Files (new):**
- `tests/tst_notificationlistener/tst_notificationlistener.pro`
- `tests/tst_notificationlistener/tst_notificationlistener.cpp`
- modify `tests/tests.pro` SUBDIRS

**Test cpp:**

```cpp
#include <QtTest/QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSignalSpy>
#include "../../src/database/NotificationListener.h"
#include "../../src/database/PostgresConnection.h"
#include "../../src/database/ConfigLoader.h"

using namespace DVE;

namespace {
DbConfig pgConfig() {
    DbConfig c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
    for (const QString& part : env.split(' ', Qt::SkipEmptyParts)) {
        const QStringList kv = part.split('=');
        if (kv.size() != 2) continue;
        const QString k = kv[0], v = kv[1];
        if      (k == "host")     c.host = v;
        else if (k == "port")     c.port = v.toInt();
        else if (k == "dbname")   c.database = v;
        else if (k == "user")     c.user = v;
        else if (k == "password") c.password = v;
    }
    return c;
}
}

class TstNotificationListener : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
            QSKIP("DVE_TEST_PG_CONN not set");
    }

    void subscribe_returnsTrue_whenConnected() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        NotificationListener nl(&pg);
        QVERIFY(nl.subscribe());
    }

    void rowChange_emits_on_insert() {
        PostgresConnection listener;
        QVERIFY(listener.open(pgConfig()));
        NotificationListener nl(&listener);
        QVERIFY(nl.subscribe());

        QSignalSpy spy(&nl, &NotificationListener::rowChanged);

        // Write from a separate connection
        PostgresConnection writer;
        QVERIFY(writer.open(pgConfig()));
        QSqlQuery q(writer.queryDb());
        QVERIFY(q.exec("INSERT INTO settings(key, value, updated_by) "
                       "VALUES ('nltest', 'a', 'bob') ON CONFLICT (key) DO UPDATE "
                       "SET value=EXCLUDED.value, updated_by=EXCLUDED.updated_by"));
        QVERIFY(spy.wait(2000));
        QVERIFY(spy.count() >= 1);

        // Cleanup
        q.exec("DELETE FROM settings WHERE key='nltest'");
    }

    void presenceChange_emits_on_insert() {
        PostgresConnection listener;
        QVERIFY(listener.open(pgConfig()));
        NotificationListener nl(&listener);
        QVERIFY(nl.subscribe());

        QSignalSpy spy(&nl, &NotificationListener::presenceChanged);

        PostgresConnection writer;
        QVERIFY(writer.open(pgConfig()));
        QSqlQuery q(writer.queryDb());
        const QString uuid = "00000000-0000-0000-0000-000000000001";
        QVERIFY(q.exec(QString(
            "INSERT INTO presence(user_uuid, user_name, user_color, "
            "resource_type, resource_id, intent) VALUES "
            "('%1'::uuid, 'B', '#3b82f6', 'file', 1, 'viewing')").arg(uuid)));

        QVERIFY(spy.wait(2000));
        QVERIFY(spy.count() >= 1);

        q.exec(QString("DELETE FROM presence WHERE user_uuid='%1'").arg(uuid));
    }
};

QTEST_MAIN(TstNotificationListener)
#include "tst_notificationlistener.moc"
```

`.pro`: standard pattern with QT += core sql testlib, SOURCES list includes NotificationListener.cpp + PostgresConnection.cpp + ConfigLoader.cpp.

Build + test with ephemeral Postgres. Commit: `test(db): tst_notificationlistener end-to-end NOTIFY round-trip`.

### Phase 2 — PresenceManager (Tasks 3–4)

#### Task 3: `PresenceManager.h`/`.cpp`

**Files (new, Python pattern):**
- `src/database/PresenceManager.h`
- `src/database/PresenceManager.cpp`

**Header:**

```cpp
#pragma once

#include <QObject>
#include <QTimer>
#include <QString>
#include <QUuid>
#include <QHash>

namespace DVE {

class PostgresConnection;
class IdentityManager;

struct PresenceRow {
    QUuid   userUuid;
    QString userName;
    QString userColor;
    QString resourceType;
    qint64  resourceId;
    QString intent;
};

class PresenceManager : public QObject {
    Q_OBJECT
public:
    PresenceManager(PostgresConnection* conn, IdentityManager* identity,
                    QObject* parent = nullptr);
    ~PresenceManager() override;

    // Switch the active resource for this user. Inserts a new presence row
    // and deletes any previous active row in the same resourceType. Starts
    // the heartbeat timer if not already running.
    bool activate(const QString& resourceType, qint64 resourceId,
                  const QString& intent = "viewing");

    // Update intent for the current active resource ("viewing" → "editing").
    bool setIntent(const QString& intent);

    // Remove this user's presence for the current active resource.
    bool deactivate();

    // List active presence rows for a resource (for UI presence dots).
    QVector<PresenceRow> activeFor(const QString& resourceType, qint64 resourceId) const;

private slots:
    void heartbeat();

private:
    PostgresConnection* m_conn;
    IdentityManager*    m_identity;
    QTimer              m_timer;
    QString             m_activeType;
    qint64              m_activeId = -1;
    QString             m_activeIntent;
};

} // namespace DVE
```

**.cpp body (key methods):**

```cpp
#include "PresenceManager.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"

#include <QSqlQuery>
#include <QSqlError>

namespace DVE {

PresenceManager::PresenceManager(PostgresConnection* conn,
                                 IdentityManager* identity, QObject* parent)
    : QObject(parent), m_conn(conn), m_identity(identity) {
    m_timer.setInterval(10000);  // 10s heartbeat per spec
    connect(&m_timer, &QTimer::timeout, this, &PresenceManager::heartbeat);
}

PresenceManager::~PresenceManager() {
    if (m_activeId >= 0) deactivate();
}

bool PresenceManager::activate(const QString& resourceType, qint64 resourceId,
                                const QString& intent) {
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;

    QSqlDatabase& db = m_conn->queryDb();
    if (!db.transaction()) return false;

    // Delete any prior active row in the same resourceType for this user
    QSqlQuery del(db);
    del.prepare("DELETE FROM presence WHERE user_uuid = ? AND resource_type = ?");
    del.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    del.addBindValue(resourceType);
    del.exec();

    // Upsert the new active row
    QSqlQuery ins(db);
    ins.prepare("INSERT INTO presence(user_uuid, user_name, user_color, "
                "resource_type, resource_id, intent, last_heartbeat) "
                "VALUES (?, ?, ?, ?, ?, ?, now()) "
                "ON CONFLICT (user_uuid, resource_type, resource_id) "
                "DO UPDATE SET intent = EXCLUDED.intent, last_heartbeat = now()");
    ins.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    ins.addBindValue(m_identity->displayName());
    ins.addBindValue(m_identity->color());
    ins.addBindValue(resourceType);
    ins.addBindValue(qlonglong(resourceId));
    ins.addBindValue(intent);
    if (!ins.exec()) {
        db.rollback();
        return false;
    }
    if (!db.commit()) return false;

    m_activeType   = resourceType;
    m_activeId     = resourceId;
    m_activeIntent = intent;
    if (!m_timer.isActive()) m_timer.start();
    return true;
}

bool PresenceManager::setIntent(const QString& intent) {
    if (m_activeId < 0 || !m_conn || !m_conn->isOpen()) return false;
    QSqlQuery q(m_conn->queryDb());
    q.prepare("UPDATE presence SET intent = ?, last_heartbeat = now() "
              "WHERE user_uuid = ? AND resource_type = ? AND resource_id = ?");
    q.addBindValue(intent);
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    q.addBindValue(m_activeType);
    q.addBindValue(qlonglong(m_activeId));
    if (!q.exec()) return false;
    m_activeIntent = intent;
    return true;
}

bool PresenceManager::deactivate() {
    if (m_activeId < 0) { m_timer.stop(); return true; }
    QSqlQuery q(m_conn->queryDb());
    q.prepare("DELETE FROM presence WHERE user_uuid = ? AND resource_type = ? "
              "AND resource_id = ?");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    q.addBindValue(m_activeType);
    q.addBindValue(qlonglong(m_activeId));
    const bool ok = q.exec();
    m_activeType.clear();
    m_activeId = -1;
    m_activeIntent.clear();
    m_timer.stop();
    return ok;
}

void PresenceManager::heartbeat() {
    if (m_activeId < 0 || !m_conn || !m_conn->isOpen()) return;
    QSqlQuery q(m_conn->queryDb());
    q.prepare("UPDATE presence SET last_heartbeat = now() "
              "WHERE user_uuid = ? AND resource_type = ? AND resource_id = ?");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    q.addBindValue(m_activeType);
    q.addBindValue(qlonglong(m_activeId));
    q.exec();
}

QVector<PresenceRow> PresenceManager::activeFor(const QString& resourceType,
                                                 qint64 resourceId) const {
    QVector<PresenceRow> result;
    if (!m_conn || !m_conn->isOpen()) return result;
    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT user_uuid::text, user_name, user_color, intent "
              "FROM presence WHERE resource_type = ? AND resource_id = ? "
              "AND last_heartbeat > now() - INTERVAL '30 seconds'");
    q.addBindValue(resourceType);
    q.addBindValue(qlonglong(resourceId));
    if (!q.exec()) return result;
    while (q.next()) {
        PresenceRow r;
        r.userUuid     = QUuid(q.value(0).toString());
        r.userName     = q.value(1).toString();
        r.userColor    = q.value(2).toString();
        r.resourceType = resourceType;
        r.resourceId   = resourceId;
        r.intent       = q.value(3).toString();
        result.push_back(r);
    }
    return result;
}

} // namespace DVE
```

Register in `.pro`. Commit: `feat(db): PresenceManager (heartbeat + single-resource-per-user)`.

#### Task 4: `tst_presencemanager` integration test

Covers `activate` (inserts + deletes prior), `setIntent`, `deactivate`, `heartbeat` refreshes `last_heartbeat`, `activeFor` filters by 30s window. Use per-slot `init()` that wipes presence table. Commit: `test(db): tst_presencemanager integration coverage`.

### Phase 3 — Optimistic concurrency in DatabaseManager (Tasks 5–9)

This is the invasive surgery. `DatabaseManager` switches from QSQLITE backend to QPSQL via `PostgresConnection`. Public method signatures stay the same so all callers in `MainWindow` keep compiling.

#### Task 5: Backend swap (open/close)

**File:** `src/database/DatabaseManager.cpp` + `.h`

Replace the SQLite-backed `m_db` with a `PostgresConnection`. Replace `open(path)` with `open(DbConfig)`. The old SQLite path is preserved as `OfflineSnapshot` (Plan C) — for Plan B, only Postgres.

Header changes:

```cpp
// Replace:
//   QSqlDatabase m_db;
// With:
class PostgresConnection;
namespace DVE { class IdentityManager; class NotificationListener; class PresenceManager; }

private:
    PostgresConnection*  m_pg = nullptr;
    IdentityManager*     m_identity = nullptr;
    QString              m_lastError;
    bool                 m_open = false;

    // Old lock-related members removed:
    // m_lockInfo, m_lockPath  — DELETED here, code removed in Plan C
```

`open(...)` signature changes:

```cpp
bool open(const DbConfig& cfg, IdentityManager* identity);
```

Implementation rewires to `m_pg->open(cfg)`, removes the SQLite path entirely. The old `pathLooksCloudSynced` and lock-file code paths get deleted (this also fulfills part of Plan C's deletion goal — moved up).

Commit: `refactor(db): DatabaseManager backend swap from QSQLITE to Postgres`.

#### Task 6: Rewrite `saveFile`/`loadFile`/`listFiles` against Postgres

The hierarchical save (file → tests → samples → data_rows → images) now uses Postgres `BIGSERIAL` IDs via `RETURNING id` instead of SQLite's `lastInsertId()`. JSONB column writes use string binding (Qt converts QString → text automatically).

Wrap in transaction. On INSERT, use `OVERRIDING SYSTEM VALUE` when re-saving with existing ids; on UPDATE, include `WHERE version = $expected` for optimistic concurrency.

The pattern:

```cpp
bool DatabaseManager::saveFile(const FileResult& result) {
    if (!m_pg || !m_pg->isOpen()) return false;
    QSqlDatabase& db = m_pg->queryDb();
    if (!db.transaction()) return false;

    // ... INSERT/UPDATE files row, fetch returned id ...
    // ... for each test, sample, data_row, image: INSERT/UPDATE ...
    //
    // If any UPDATE returns rowcount == 0 (version mismatch or row deleted),
    //   rollback and surface a conflict via m_lastError + return false.
    //   The MainWindow caller will translate this into a ConflictResolver dialog.

    if (!db.commit()) return false;
    return true;
}
```

Update all save callers in MainWindow to pass identity's UUID via the chain (the trigger uses the row's `updated_by` column directly).

Commit: `refactor(db): saveFile/loadFile/listFiles rewired to Postgres`.

#### Task 7: Rewrite sensory + detailed sensory + settings

Same pattern. Sensory sessions use the `(session_name, tester_name, date)` natural-key UNIQUE constraint. ON CONFLICT (session_name, tester_name, date) DO UPDATE SET ... (the existing INSERT OR REPLACE upsert pattern translates).

Commit: `refactor(db): sensory and settings methods rewired to Postgres`.

#### Task 8: Add `version` round-trip + optimistic UPDATE

For every editable struct (`FileResult`, `SensorySession`, `DetailedSensorySession`, `SampleResult`, `DataRow`, etc.), add a `version` field. Read it via `SELECT ..., version`. Pass it to UPDATEs via `WHERE id = ? AND version = ?`. Detect rowcount=0 → conflict.

The `lastError` semantics expand: distinguish VERSION_MISMATCH vs ROW_DELETED via a follow-up `SELECT version FROM ... WHERE id = ?`.

New enum:

```cpp
namespace DVE {
enum class WriteResult {
    Success,
    VersionMismatch,    // expected version stale, server has newer
    RowDeleted,         // row no longer exists
    UniqueViolation,    // hit a unique constraint
    OtherError
};
}
```

Modify save returns from `bool` to `WriteResult` (callers in MainWindow need to handle the conflict cases). The `bool` overloads can stay as adapter wrappers that map non-Success to `false`.

Commit: `feat(db): WriteResult enum + optimistic UPDATE with version check`.

#### Task 9: tst_databasemanager rewrite

Existing 24 tests assumed SQLite backend. About a third assumed the lock file. Rewrite the test fixture to use the ephemeral Postgres pattern (init.sql preloaded). Adapt the assertions for `WriteResult` enum. Test the optimistic conflict path (race two clients on the same row).

Commit: `test(db): retarget tst_databasemanager at Postgres backend + add conflict tests`.

### Phase 4 — ConflictResolver + dialogs (Tasks 10–13)

#### Task 10: `ConflictResolver.h`/`.cpp`

```cpp
namespace DVE {
class ConflictResolver : public QObject {
    Q_OBJECT
public:
    // Returns the user's resolution choice for a version-mismatch conflict.
    enum Choice { KeepMine, TakeTheirs, ApplySelection, Cancel };

    Choice resolveVersionMismatch(const QString& table, qint64 id,
                                  const QVariantMap& mine,
                                  const QVariantMap& theirs,
                                  QVariantMap* merged,
                                  QWidget* parent = nullptr);

    enum DeletedChoice { Recreate, Discard };
    DeletedChoice resolveDeleted(const QString& table, qint64 id,
                                 const QVariantMap& mine,
                                 QWidget* parent = nullptr);

    enum UniqueChoice { OpenExisting, RenameMine, CancelUnique };
    UniqueChoice resolveUniqueViolation(const QString& table,
                                        const QString& conflictDescription,
                                        QString* renamedName,
                                        QWidget* parent = nullptr);
};
}
```

Each method pops up the corresponding dialog (Task 11–13) and returns the user's choice. The dialog widgets are dumb — they show data and capture clicks; the resolver wraps them.

Commit: `feat(db): ConflictResolver dispatching three dialog types`.

#### Task 11: `VersionMismatchDialog`

QDialog with a 4-column table: Field | Their value | Your value | Pick (radio: same|theirs|yours). Bulk buttons at bottom: "Keep all mine" / "Take all theirs" / "Apply selection". Identical fields display as `●same` and are not selectable.

Cell-by-cell field comparison: iterate over the mine/theirs QVariantMap, compare values, build rows.

Commit: `feat(ui): VersionMismatchDialog with per-field selection`.

#### Task 12: `RowDeletedDialog`

QDialog showing the row's table + id, a list of the user's unsaved edits, two buttons: "Recreate as new" / "Discard my changes".

Commit: `feat(ui): RowDeletedDialog`.

#### Task 13: `UniqueViolationDialog`

QDialog showing the conflicting key, the existing creator, three buttons: "Open existing" / "Rename mine to <suggestion>" / "Cancel". For rename, suggest `<name> (2)`.

Commit: `feat(ui): UniqueViolationDialog with rename suggestion`.

### Phase 5 — Presence UI (Tasks 14–16)

#### Task 14: Presence dots in nav widgets

`MainWindow.cpp` populates the TPM file tree, sensory list, and detailed sensory list. Add a custom `QTreeWidgetItem`/`QListWidgetItem` subclass (or use `setData` with a custom role) to track presence. Render a colored dot next to the name.

When `PresenceManager::activeFor()` returns rows, paint colored dots. Hover tooltip lists names + intents.

```cpp
// In MainWindow's resource-list-refresh routine:
auto rows = m_presence->activeFor("file", file.id);
QString tooltipParts;
QStringList colors;
for (const PresenceRow& r : rows) {
    tooltipParts += QString("%1 (%2)\n").arg(r.userName, r.intent);
    colors << r.userColor;
}
item->setToolTip(0, tooltipParts.trimmed());
item->setData(0, Qt::UserRole + 1, colors);
// Custom paint delegate uses the UserRole+1 colors to draw stacked dots.
```

A `PresenceDotsDelegate : public QStyledItemDelegate` overrides `paint` to render the colored dots after the text. Apply via `tree->setItemDelegateForColumn(0, new PresenceDotsDelegate)`.

Commit: `feat(ui): PresenceDotsDelegate + nav tree integration`.

#### Task 15: Avatar row in central editor

`MainWindow.cpp` adds a small `PresenceAvatarBar` widget above the central widget, top-right corner. Each active user is a circle with their initial. Hover shows full name + intent. Click opens a small popup.

```cpp
class PresenceAvatarBar : public QWidget {
    Q_OBJECT
public:
    void setPresence(const QVector<PresenceRow>& rows, const QUuid& selfUuid);
};
```

Self-rendered with a thicker ring; other users plain.

Commit: `feat(ui): PresenceAvatarBar top-right of central editor`.

#### Task 16: Wire NotificationListener::presenceChanged → UI refresh

```cpp
connect(m_notify, &NotificationListener::presenceChanged, this,
        [this](const PresenceChange& c) {
            // Refresh the nav row + avatar bar for c.resourceType / c.resourceId
            refreshPresenceFor(c.resourceType, c.resourceId);
        });
```

`refreshPresenceFor` queries `PresenceManager::activeFor(...)` and updates the corresponding QTreeWidgetItem's userRole color list + the avatar bar.

Commit: `feat(ui): live presence updates via NOTIFY → UI refresh`.

### Phase 6 — Don't-yank-in-progress-edit (Tasks 17–19)

#### Task 17: Cell-level "in-progress" tracking in TPM data table

Add a `dve_editing` property on the QTableWidget's focused cell when the user is typing and the value differs from the loaded baseline. The TPM data table widget is already populated by `MainWindow`; identify the relevant connect points.

```cpp
connect(tableWidget, &QTableWidget::itemChanged, this, [](QTableWidgetItem* it) {
    const QVariant baseline = it->data(Qt::UserRole + 2);
    const bool dirty = it->text() != baseline.toString();
    it->setData(Qt::UserRole + 3, dirty);  // dve_editing
});
```

Commit: `feat(ui): cell-level dve_editing flag for in-progress edits`.

#### Task 18: Yellow border + tooltip on remote cell change

`NotificationListener::rowChanged` arrives. Look up the current TPM table for that table/id. If the user has `dve_editing=true` on that cell:

```cpp
// Decorate the cell with a yellow border + tooltip "changed by Sarah just now"
item->setBackground(QBrush(QColor("#fef3c7")));
item->setToolTip(QString("changed by %1 — click to take their value").arg(remoteUserName));
item->setData(Qt::UserRole + 4, /*pending remote value*/);
```

If clicked while pending: replace the user's text with the remote value, clear the decoration.

Commit: `feat(ui): yellow-border decoration for remote changes during in-progress edit`.

#### Task 19: Row-deleted toast

`NotificationListener::rowChanged` with `op="DELETE"` for a row currently open: show a non-modal toast (a `QFrame` styled as a banner) saying "This row was deleted by <name>. Click to recreate."

Commit: `feat(ui): row-deleted toast for currently-open rows`.

### Phase 7 — Wire it all together (Tasks 20–22)

#### Task 20: `MainWindow` startup wiring

After `IdentityManager` (already wired by Plan A), construct:

```cpp
// In MainWindow constructor, after m_identity = new IdentityManager(this):
m_pgConn = new PostgresConnection(this);
DbConfig cfg;
QString err;
const QString confPath = QString::fromLocal8Bit(qgetenv("PROGRAMDATA"))
                          + "/DataViewer/db.conf";
if (!ConfigLoader::load(confPath, cfg, &err)) {
    QMessageBox::critical(this, tr("Config error"), err);
    return;
}
if (!m_pgConn->open(cfg)) {
    QMessageBox::critical(this, tr("Database error"),
                          tr("Could not connect to Postgres: ") + m_pgConn->lastError());
    return;
}

m_notify = new NotificationListener(m_pgConn, this);
m_notify->subscribe();

m_presence = new PresenceManager(m_pgConn, m_identity, this);
m_conflict = new ConflictResolver(this);

// Repoint the existing DatabaseManager:
m_db->open(cfg, m_identity);
```

Order matters: open Postgres before NotificationListener, before PresenceManager.

Commit: `feat(ui): MainWindow constructs PostgresConnection + NotificationListener + PresenceManager`.

#### Task 21: Save-path conflict handling

Update every save call site in MainWindow (e.g., `flushExcelWrites`, save buttons, etc.) to handle `WriteResult` and route to ConflictResolver. Pattern:

```cpp
const auto result = m_db->saveFile(fileResult);
if (result == WriteResult::VersionMismatch) {
    // Reload fresh from server, diff, show ConflictResolver
    FileResult fresh = m_db->loadFileByPath(fileResult.filePath);
    auto choice = m_conflict->resolveVersionMismatch(...);
    if (choice == ConflictResolver::TakeTheirs) {
        // Just discard local edits and reload
    } else if (choice == ConflictResolver::KeepMine) {
        // Retry with fresh.version
        fileResult.version = fresh.version;
        m_db->saveFile(fileResult);  // expected to succeed now
    } else if (choice == ConflictResolver::ApplySelection) {
        // ... use merged values ...
    }
} else if (result == WriteResult::RowDeleted) { ... }
else if (result == WriteResult::UniqueViolation) { ... }
```

Commit: `feat(ui): save-path conflict handling via ConflictResolver`.

#### Task 22: Filter own-UUID echoes in NotificationListener consumers

In MainWindow:

```cpp
connect(m_notify, &NotificationListener::rowChanged, this,
    [this](const RowChange& c) {
        if (c.updatedBy == m_identity->uuid().toString(QUuid::WithoutBraces))
            return;  // It's my own write echoing back; ignore.
        // ... refresh UI ...
    });
```

Commit: `feat(ui): filter own-UUID NOTIFY echoes to avoid self-refresh`.

### Phase 8 — Integration tests + checkpoint (Tasks 23–25)

#### Task 23: Two-client e2e test

`tests/tst_twoclient_e2e/tst_twoclient_e2e.cpp` spawns two PostgresConnections, exercises:
- Client A inserts a row; client B's NotificationListener fires within 1s.
- Client A updates; client B's UI sees the change.
- Both clients try to update the same row at the same version; one succeeds, one gets VersionMismatch.

Commit: `test(e2e): two-client NOTIFY + conflict round-trip`.

#### Task 24: Update CLAUDE.md "DatabaseManager" mention

Now that DatabaseManager is Postgres-backed, refresh the Architecture section.

Commit: `docs(claude): DatabaseManager is now Postgres-backed`.

#### Task 25: Plan B checkpoint

Run the full test suite. Manual smoke: install build, two clients edit same file, verify presence dots + conflict dialog + live refresh. Update INDEX status to "Complete and verified".

Commit: `docs(plans): mark Plan B complete`.

---

## Plan B Checkpoint Criteria

1. Two `DataViewer.exe` instances on different machines can edit the same TPM file simultaneously; NOTIFY events propagate within 1s.
2. Conflict dialogs (all three types) surface on collision.
3. Presence dots visible in nav for shared files.
4. Self-test passes: `--self-test` reports `postgres_connection: passed` plus new presence/notify cases.
5. All test suites pass including the new tst_notificationlistener, tst_presencemanager, tst_databasemanager (retargeted), tst_twoclient_e2e.

Once verified, proceed to Plan C.

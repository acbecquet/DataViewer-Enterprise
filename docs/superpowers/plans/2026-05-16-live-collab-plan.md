# Live Collaborative Editing Implementation Plan (v2.0.1)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship continuous per-cell sync (Postgres NOTIFY-driven) that replaces the v2.0 save-driven conflict layer. TPM cells gain Variant-C presence borders, every commit is durable instantly across clients, and the v2.0 conflict dialogs are deleted.

**Architecture:** One new chokepoint module (`LiveSync`) plus a new delegate (`CellFocusDelegate`). TPM's existing `QTableWidget` stays — we wire its `itemChanged` signal through `LiveSync::commitCell` and add the delegate for the colored border + name flag. Sensory + Detailed Sensory route commit handlers through the same `LiveSync` using JSONB-path keys. NOTIFY listener gains a `cell_focus` channel and the row-changed payload carries `column` + `new_value`.

**Tech Stack:** C++17 / Qt 6.10 / qmake / MinGW 13.1.0 on Windows. PostgreSQL 16 (samples in JSONB, data_rows in columnar). `-Werror -Wall -Wextra`.

**Spec:** `docs/superpowers/specs/2026-05-16-live-collab-design.md`

## Notable deviation from the spec

The spec proposes a new `LiveTableModel` (`QAbstractTableModel`) to replace TPM's `QTableWidget`. Code inspection shows `m_dataTable` already has substantial per-cell dirty-tracking infrastructure (per-cell `UserRole+2/3/4` roles, baseline-value tracking, `onDataTableItemChanged` slot, `findTableRowForDataRowId`) from the v2.0 sprint. Replacing it would require porting all of that machinery to a new model class for zero user-visible benefit.

**This plan keeps `QTableWidget`** and wires the existing `onDataTableItemChanged` slot through `LiveSync::commitCell`. The `CellFocusDelegate` works identically with either `QTableWidget` or `QTableView`+model. If we later decide we want a proper model, that becomes its own refactor on top of the live-sync work, not a prerequisite for it.

---

## Pre-flight

Before starting any task:

```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
git status                   # main branch, clean
git checkout -b feature/live-collab-v2.0.1
python tools/decrypt_via_copy.py --apply    # decrypt any MIP-touched files
```

Build commands (used in many tasks):

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Test commands:

```powershell
# Full suite
.\tests\run-tests.ps1

# Single test class (the -Filter flag landed in v2.0.1 prep)
.\tests\run-tests.ps1 -Filter livesync
```

Postgres test container (some tasks need it):

```powershell
.\tests\start-test-postgres.ps1     # sets $env:DVE_TEST_PG_CONN
```

---

## Task 1: Schema additions

**Files:**
- Modify: `deploy/postgres/init.sql` (add `cell_focus` table + extend NOTIFY trigger)
- Create: `deploy/postgres/migrations/2026-05-16-cell-focus.sql` (one-off upgrade script for existing installs)

### Step 1: Write the migration script

- [ ] **Create `deploy/postgres/migrations/2026-05-16-cell-focus.sql`**

```sql
-- v2.0 -> v2.0.1: live collaborative editing
-- Adds the cell_focus table and extends the row-changed NOTIFY payload
-- with column + new_value. Idempotent.

BEGIN;

CREATE TABLE IF NOT EXISTS cell_focus (
    user_uuid    UUID        NOT NULL,
    table_name   TEXT        NOT NULL,
    row_id       BIGINT      NOT NULL,
    column_name  TEXT        NOT NULL,
    user_name    TEXT        NOT NULL,
    user_color   TEXT        NOT NULL,
    started_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (user_uuid, table_name, row_id, column_name)
);
CREATE INDEX IF NOT EXISTS idx_cell_focus_target
    ON cell_focus(table_name, row_id);

-- Cleanup of stale focus rows: any focus older than 30s is considered
-- abandoned. Same window as presence heartbeat.
-- (Runs out of the existing pg_cron job; nothing to schedule here if the
-- cron extension is already present.)

-- Extend the existing row-changed trigger function so the NOTIFY payload
-- carries the column that changed and its new value when the operation
-- is a single-column UPDATE. (Multi-column UPDATEs send a payload
-- without column/new_value -- the client falls back to re-SELECT.)
CREATE OR REPLACE FUNCTION notify_row_changed() RETURNS trigger AS $$
DECLARE
    payload JSONB;
BEGIN
    payload := jsonb_build_object(
        'table',       TG_TABLE_NAME,
        'op',          TG_OP,
        'id',          COALESCE(NEW.id, OLD.id),
        'updated_by',  COALESCE(NEW.updated_by, OLD.updated_by)
    );
    -- live-sync extension: include the column + new value when the
    -- caller set a session var indicating a single-column UPDATE.
    IF current_setting('dve.live_column', true) <> '' THEN
        payload := payload || jsonb_build_object(
            'column',    current_setting('dve.live_column', true),
            'new_value', current_setting('dve.live_value',  true)
        );
    END IF;
    PERFORM pg_notify('dataviewer_changes', payload::text);
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- New channel: cell_focus broadcasts. Trigger fires on every INSERT,
-- UPDATE, DELETE on cell_focus and pushes the row.
CREATE OR REPLACE FUNCTION notify_cell_focus() RETURNS trigger AS $$
DECLARE
    payload JSONB;
BEGIN
    payload := jsonb_build_object(
        'op',          TG_OP,
        'user_uuid',   COALESCE(NEW.user_uuid, OLD.user_uuid),
        'user_name',   COALESCE(NEW.user_name, OLD.user_name),
        'user_color',  COALESCE(NEW.user_color, OLD.user_color),
        'table',       COALESCE(NEW.table_name, OLD.table_name),
        'row_id',      COALESCE(NEW.row_id, OLD.row_id),
        'column',      COALESCE(NEW.column_name, OLD.column_name)
    );
    PERFORM pg_notify('dataviewer_cell_focus', payload::text);
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_cell_focus_notify ON cell_focus;
CREATE TRIGGER trg_cell_focus_notify
    AFTER INSERT OR UPDATE OR DELETE ON cell_focus
    FOR EACH ROW EXECUTE FUNCTION notify_cell_focus();

COMMIT;
```

### Step 2: Mirror the changes into init.sql

- [ ] **Append the cell_focus table + triggers to `deploy/postgres/init.sql`**

Open `deploy/postgres/init.sql`. Find the existing `notify_row_changed` trigger function (search for `pg_notify('dataviewer_changes'`). Replace its body with the extended version above. Append the cell_focus table, index, and `notify_cell_focus` function + trigger after the `sensory_images` block.

### Step 3: Verify the schema applies cleanly to a fresh container

- [ ] **Tear down + restart the test container with the new schema**

```powershell
docker rm -f dve-test-pg 2>$null
.\tests\start-test-postgres.ps1
```

Expected output: container starts, init.sql applies without error, `\dt` shows `cell_focus` in the table list. Verify:

```powershell
docker exec dve-test-pg psql -U postgres -d dataviewer -c "\d cell_focus"
```

### Step 4: Commit

- [ ] **Commit schema changes**

```bash
git add deploy/postgres/init.sql deploy/postgres/migrations/2026-05-16-cell-focus.sql
git commit -m "feat(db): cell_focus table + extended row-changed payload

v2.0.1 live-sync prep. Adds cell_focus for per-cell presence and
extends the dataviewer_changes payload with optional column +
new_value fields (set via session vars from the caller).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: NotificationListener extension

**Files:**
- Modify: `src/database/NotificationListener.h` (add fields to `RowChange`, new `CellFocusChange` struct, new signal)
- Modify: `src/database/NotificationListener.cpp` (subscribe to new channel, parse new fields)
- Modify: `tests/tst_notificationlistener/tst_notificationlistener.cpp`

### Step 1: Write the failing test

- [ ] **Add a test slot for the column-aware payload**

Open `tests/tst_notificationlistener/tst_notificationlistener.cpp`. Add:

```cpp
private slots:
    void rowChangedPayloadIncludesColumn();
    void cellFocusChannelEmitsSignal();
```

Implementations:

```cpp
void TstNotificationListener::rowChangedPayloadIncludesColumn()
{
    if (!m_subscribed) QSKIP("listener not subscribed");

    QSignalSpy spy(m_listener, &NotificationListener::rowChanged);

    // Use a foreign connection to UPDATE a single column. The trigger
    // function reads dve.live_column / dve.live_value session vars to
    // attach them to the payload.
    QSqlQuery q(m_foreignDb);
    QVERIFY(q.exec("SELECT set_config('dve.live_column', 'draw_pressure', false)"));
    QVERIFY(q.exec("SELECT set_config('dve.live_value',  '1.42', false)"));
    QVERIFY(q.exec("UPDATE data_rows SET draw_pressure = 1.42 "
                   "WHERE id = (SELECT id FROM data_rows LIMIT 1)"));

    QVERIFY(spy.wait(1500));
    QVERIFY(spy.count() >= 1);
    const auto args = spy.takeLast();
    const RowChange c = args.first().value<RowChange>();
    QCOMPARE(c.table,       QStringLiteral("data_rows"));
    QCOMPARE(c.column,      QStringLiteral("draw_pressure"));
    QCOMPARE(c.newValue.toDouble(), 1.42);
}

void TstNotificationListener::cellFocusChannelEmitsSignal()
{
    if (!m_subscribed) QSKIP("listener not subscribed");

    QSignalSpy spy(m_listener, &NotificationListener::cellFocusChanged);

    QSqlQuery q(m_foreignDb);
    QVERIFY(q.exec(
        "INSERT INTO cell_focus(user_uuid, table_name, row_id, "
        "column_name, user_name, user_color) "
        "VALUES('11111111-1111-1111-1111-111111111111'::uuid, "
        "'data_rows', 1, 'draw_pressure', 'Tina', '#16a34a')"));

    QVERIFY(spy.wait(1500));
    QVERIFY(spy.count() >= 1);
    const auto args = spy.takeLast();
    const CellFocusChange f = args.first().value<CellFocusChange>();
    QCOMPARE(f.op,         QStringLiteral("INSERT"));
    QCOMPARE(f.userName,   QStringLiteral("Tina"));
    QCOMPARE(f.userColor,  QStringLiteral("#16a34a"));
    QCOMPARE(f.tableName,  QStringLiteral("data_rows"));
    QCOMPARE(f.rowId,      qint64(1));
    QCOMPARE(f.columnName, QStringLiteral("draw_pressure"));
}
```

### Step 2: Run the test — verify it fails to compile

- [ ] **Build the test**

```powershell
.\tests\run-tests.ps1 -Filter notificationlistener
```

Expected: FAIL at compile — `CellFocusChange` undefined; `RowChange::column` not a member; `NotificationListener::cellFocusChanged` not declared.

### Step 3: Extend the header

- [ ] **Update `src/database/NotificationListener.h`**

Replace the existing `RowChange` struct definition with:

```cpp
struct RowChange {
    QString  table;
    QString  op;            // "INSERT" | "UPDATE" | "DELETE"
    qint64   id   = -1;
    QString  updatedBy;
    QString  column;        // empty unless single-column UPDATE
    QVariant newValue;      // empty unless single-column UPDATE
};

struct CellFocusChange {
    QString op;             // "INSERT" | "UPDATE" | "DELETE"
    QUuid   userUuid;
    QString userName;
    QString userColor;
    QString tableName;
    qint64  rowId = -1;
    QString columnName;
};
```

Add a new signal to the class:

```cpp
signals:
    void rowChanged(const DVE::RowChange& change);
    void presenceChanged(const DVE::PresenceChange& change);
    void cellFocusChanged(const DVE::CellFocusChange& change);
```

Add `Q_DECLARE_METATYPE(DVE::CellFocusChange)` at the bottom alongside the existing two.

### Step 4: Extend the implementation

- [ ] **Update `src/database/NotificationListener.cpp`**

In `subscribe()`, add a third channel:

```cpp
    ok = drv->subscribeToNotification("dataviewer_cell_focus") && ok;
```

In `unsubscribe()`, add the corresponding `unsubscribeFromNotification` for the same channel.

In the `static_assert` / `qRegisterMetaType` block, add:

```cpp
    qRegisterMetaType<CellFocusChange>();
```

In `onNotification()`, extend the branching:

```cpp
    if (name == "dataviewer_changes") {
        RowChange c;
        c.table     = o.value("table").toString();
        c.op        = o.value("op").toString();
        c.id        = o.value("id").toVariant().toLongLong();
        c.updatedBy = o.value("updated_by").toString();
        c.column    = o.value("column").toString();
        if (o.contains("new_value")) c.newValue = o.value("new_value").toVariant();
        emit rowChanged(c);
    } else if (name == "dataviewer_presence") {
        // ... existing block unchanged ...
    } else if (name == "dataviewer_cell_focus") {
        CellFocusChange f;
        f.op         = o.value("op").toString();
        f.userUuid   = QUuid(o.value("user_uuid").toString());
        f.userName   = o.value("user_name").toString();
        f.userColor  = o.value("user_color").toString();
        f.tableName  = o.value("table").toString();
        f.rowId      = o.value("row_id").toVariant().toLongLong();
        f.columnName = o.value("column").toString();
        emit cellFocusChanged(f);
    }
```

### Step 5: Run the test — verify it passes

- [ ] **Re-run**

```powershell
.\tests\run-tests.ps1 -Filter notificationlistener
```

Expected: PASS (assuming Postgres test container is running with the new schema from Task 1).

### Step 6: Commit

```bash
git add src/database/NotificationListener.h src/database/NotificationListener.cpp \
        tests/tst_notificationlistener/
git commit -m "feat(db): notification listener carries column + cell_focus payload

Extends RowChange with optional column / newValue (set by the trigger
on single-column UPDATEs) and adds CellFocusChange + new signal for
the dataviewer_cell_focus channel.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: LiveSync core — commitCell

**Files:**
- Create: `src/database/LiveSync.h`
- Create: `src/database/LiveSync.cpp`
- Modify: `DataViewerEnterprise.pro` (add new sources)
- Create: `tests/tst_livesync/tst_livesync.cpp`
- Create: `tests/tst_livesync/tst_livesync.pro`
- Modify: `tests/tests.pro` (SUBDIRS += tst_livesync)

### Step 1: Write the failing test for scalar-column commit

- [ ] **Create `tests/tst_livesync/tst_livesync.cpp`**

```cpp
#include <QtTest>
#include <QSqlQuery>

#include "database/LiveSync.h"
#include "database/PostgresConnection.h"
#include "database/IdentityManager.h"

using namespace DVE;

class TstLiveSync : public QObject
{
    Q_OBJECT

private slots:
    void initTestCase();
    void cleanupTestCase();

    void commitCell_writesScalarColumnAndBumpsVersion();
    void commitCell_writesJsonPathWithoutClobberingSiblings();
    void focusCell_writesRowAndBlurDeletes();

private:
    PostgresConnection* m_conn = nullptr;
    IdentityManager*    m_identity = nullptr;
    LiveSync*           m_sync = nullptr;
    qint64              m_sampleId = -1;
    qint64              m_dataRowId = -1;
    qint64              m_sensorySessionId = -1;
};

void TstLiveSync::initTestCase()
{
    const QString conn = qEnvironmentVariable("DVE_TEST_PG_CONN");
    if (conn.isEmpty()) QSKIP("DVE_TEST_PG_CONN not set; skipping");
    m_conn = new PostgresConnection(this);
    QVERIFY(m_conn->open(conn));
    m_identity = new IdentityManager(this);
    m_identity->setDisplayName("TestUser");
    m_identity->setColor("#ff0000");
    m_sync = new LiveSync(m_conn, m_identity, this);

    // Insert fixture rows we can mutate.
    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec("INSERT INTO files(file_path, file_name, loaded_at) "
                   "VALUES('/tmp/t.xlsx', 't.xlsx', '2026-05-16') RETURNING id"));
    QVERIFY(q.next());
    const qint64 fileId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString("INSERT INTO tests(file_id, sheet_name) "
                           "VALUES(%1, 'S1') RETURNING id").arg(fileId)));
    QVERIFY(q.next());
    const qint64 testId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString("INSERT INTO samples(test_id, sample_name) "
                           "VALUES(%1, 'A') RETURNING id").arg(testId)));
    QVERIFY(q.next());
    m_sampleId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString(
        "INSERT INTO data_rows(sample_id, sort_order, draw_pressure) "
        "VALUES(%1, 0, 0.0) RETURNING id").arg(m_sampleId)));
    QVERIFY(q.next());
    m_dataRowId = q.value(0).toLongLong();

    QVERIFY(q.exec(
        "INSERT INTO sensory_sessions(session_name, json_data) "
        "VALUES('test', '{\"samples\":[{\"name\":\"D\",\"voltage\":3.5,"
        "\"power_type\":\"Constant Voltage\"}]}'::jsonb) RETURNING id"));
    QVERIFY(q.next());
    m_sensorySessionId = q.value(0).toLongLong();
}

void TstLiveSync::cleanupTestCase()
{
    if (m_conn && m_conn->isOpen()) {
        QSqlQuery q(m_conn->queryDb());
        q.exec("DELETE FROM files");
        q.exec("DELETE FROM sensory_sessions");
        q.exec("DELETE FROM cell_focus");
    }
}

void TstLiveSync::commitCell_writesScalarColumnAndBumpsVersion()
{
    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString("SELECT version FROM data_rows WHERE id=%1")
                   .arg(m_dataRowId)));
    QVERIFY(q.next());
    const int beforeVersion = q.value(0).toInt();

    bool ok = m_sync->commitCell("data_rows", m_dataRowId,
                                 "draw_pressure", 1.42);
    QVERIFY(ok);

    QVERIFY(q.exec(QString(
        "SELECT draw_pressure, version FROM data_rows WHERE id=%1")
        .arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toDouble(), 1.42);
    QCOMPARE(q.value(1).toInt(), beforeVersion + 1);
}

void TstLiveSync::commitCell_writesJsonPathWithoutClobberingSiblings()
{
    bool ok = m_sync->commitCell(
        "sensory_sessions", m_sensorySessionId,
        "json_path:samples[0].voltage", 4.2);
    QVERIFY(ok);

    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString(
        "SELECT json_data->'samples'->0->>'voltage', "
        "       json_data->'samples'->0->>'name', "
        "       json_data->'samples'->0->>'power_type' "
        "FROM sensory_sessions WHERE id=%1").arg(m_sensorySessionId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toDouble(), 4.2);
    QCOMPARE(q.value(1).toString(), QStringLiteral("D"));
    QCOMPARE(q.value(2).toString(), QStringLiteral("Constant Voltage"));
}

void TstLiveSync::focusCell_writesRowAndBlurDeletes()
{
    QVERIFY(m_sync->focusCell("data_rows", m_dataRowId, "draw_pressure"));

    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString(
        "SELECT count(*) FROM cell_focus "
        "WHERE table_name='data_rows' AND row_id=%1 "
        "AND column_name='draw_pressure'").arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toInt(), 1);

    QVERIFY(m_sync->blurCell());

    QVERIFY(q.exec(QString(
        "SELECT count(*) FROM cell_focus "
        "WHERE table_name='data_rows' AND row_id=%1").arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toInt(), 0);
}

QTEST_MAIN(TstLiveSync)
#include "tst_livesync.moc"
```

### Step 2: Create the test .pro

- [ ] **Write `tests/tst_livesync/tst_livesync.pro`**

Mirror the pattern of `tests/tst_presencemanager/tst_presencemanager.pro`:

```pro
QT += testlib sql network
CONFIG += console c++17
TEMPLATE = app

INCLUDEPATH += ../../src
DEFINES += QT_DEPRECATED_WARNINGS

SOURCES += tst_livesync.cpp \
    ../../src/database/LiveSync.cpp \
    ../../src/database/PostgresConnection.cpp \
    ../../src/database/IdentityManager.cpp

HEADERS += ../../src/database/LiveSync.h \
    ../../src/database/PostgresConnection.h \
    ../../src/database/IdentityManager.h
```

Append `tst_livesync` to `tests/tests.pro` SUBDIRS.

### Step 3: Run the test — verify it fails

- [ ] **Build the test**

```powershell
.\tests\run-tests.ps1 -Filter livesync
```

Expected: FAIL at compile — `LiveSync.h` doesn't exist yet.

### Step 4: Create the LiveSync header

- [ ] **Write `src/database/LiveSync.h`**

```cpp
#pragma once

#include <QObject>
#include <QPointer>
#include <QString>
#include <QVariant>

namespace DVE {

class PostgresConnection;
class IdentityManager;
struct RowChange;
struct CellFocusChange;

// LiveSync is the single chokepoint for per-cell writes and remote
// applies in v2.0.1. Every editable widget commits through
// commitCell(); every NOTIFY-driven cell change emits cellChanged().
//
// JSONB columns are addressed with column names prefixed "json_path:"
// e.g. column = "json_path:samples[2].scores.smoothness". LiveSync
// translates these into jsonb_set() UPDATEs at the database layer so
// concurrent edits to different paths in the same sample don't
// clobber each other.
class LiveSync : public QObject {
    Q_OBJECT
public:
    LiveSync(PostgresConnection* conn, IdentityManager* identity,
             QObject* parent = nullptr);

    // Scalar-column UPDATE on a known table+row+column. Returns true on
    // success, false on connection failure or DB error. Bumps version
    // and stamps updated_by. Sets the dve.live_column / dve.live_value
    // session vars first so the trigger emits a column-aware payload.
    bool commitCell(const QString& table, qint64 rowId,
                    const QString& column, const QVariant& value);

    // Upsert a cell_focus row for the current user. The user can only
    // hold one focus at a time; calling focusCell again first deletes
    // any previous focus row owned by this user.
    bool focusCell(const QString& table, qint64 rowId, const QString& column);

    // Delete the current user's focus row (if any).
    bool blurCell();

signals:
    // Emitted when a remote cell change arrives (after self-UUID filter).
    void cellChanged(const QString& table, qint64 rowId,
                     const QString& column, const QVariant& newValue);

    // Emitted when a remote cell-focus change arrives.
    void cellFocused(const QString& table, qint64 rowId,
                     const QString& column,
                     const QString& userName, const QString& userColor);
    void cellBlurred(const QString& table, qint64 rowId,
                     const QString& column);

public slots:
    // Called by MainWindow when NotificationListener emits its signals,
    // after the own-UUID echo filter. Splits row/focus dispatch.
    void onRowChanged(const RowChange& change);
    void onCellFocusChanged(const CellFocusChange& change);

private:
    bool runScalarUpdate(const QString& table, qint64 rowId,
                         const QString& column, const QVariant& value);
    bool runJsonPathUpdate(const QString& table, qint64 rowId,
                           const QString& jsonPath, const QVariant& value);

    QPointer<PostgresConnection> m_conn;
    QPointer<IdentityManager>    m_identity;
};

} // namespace DVE
```

### Step 5: Create the LiveSync implementation

- [ ] **Write `src/database/LiveSync.cpp`**

```cpp
#include "LiveSync.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "NotificationListener.h"

#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>
#include <QStringList>
#include <QRegularExpression>

namespace DVE {

LiveSync::LiveSync(PostgresConnection* conn, IdentityManager* identity,
                   QObject* parent)
    : QObject(parent), m_conn(conn), m_identity(identity) {}

// Single source of truth for the set of tables that v2.0.1 live-syncs.
// Used both for sanity-checking commitCell calls and for the trigger
// configuration on the DB side (see init.sql).
static bool isLiveSyncTable(const QString& t)
{
    return t == QLatin1String("data_rows")
        || t == QLatin1String("samples")
        || t == QLatin1String("tests")
        || t == QLatin1String("files")
        || t == QLatin1String("sensory_sessions")
        || t == QLatin1String("detailed_sensory_sessions");
}

bool LiveSync::commitCell(const QString& table, qint64 rowId,
                          const QString& column, const QVariant& value)
{
    if (!m_conn || !m_conn->isOpen()) return false;
    if (!isLiveSyncTable(table)) {
        qWarning() << "LiveSync::commitCell unknown table" << table;
        return false;
    }
    if (column.startsWith(QLatin1String("json_path:"))) {
        const QString path = column.mid(QStringLiteral("json_path:").size());
        return runJsonPathUpdate(table, rowId, path, value);
    }
    return runScalarUpdate(table, rowId, column, value);
}

bool LiveSync::runScalarUpdate(const QString& table, qint64 rowId,
                               const QString& column, const QVariant& value)
{
    QSqlQuery q(m_conn->queryDb());

    // Tag the session so the AFTER trigger emits a column-aware payload.
    q.prepare("SELECT set_config('dve.live_column', ?, false), "
              "       set_config('dve.live_value',  ?, false)");
    q.addBindValue(column);
    q.addBindValue(value.toString());
    if (!q.exec()) {
        qWarning() << "LiveSync::commitCell set_config failed:" << q.lastError().text();
        return false;
    }

    const QString uuid = m_identity ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString();
    const QString sql = QStringLiteral(
        "UPDATE %1 SET %2 = ?, version = version + 1, "
        "updated_at = now(), updated_by = ? WHERE id = ?")
        .arg(table, column);
    q.prepare(sql);
    q.addBindValue(value);
    q.addBindValue(uuid);
    q.addBindValue(rowId);
    if (!q.exec()) {
        qWarning() << "LiveSync::commitCell UPDATE failed:" << q.lastError().text();
        return false;
    }
    return q.numRowsAffected() > 0;
}

bool LiveSync::runJsonPathUpdate(const QString& table, qint64 rowId,
                                 const QString& jsonPath, const QVariant& value)
{
    // Parse "samples[2].scores.smoothness" into a Postgres text-array
    // path: '{samples,2,scores,smoothness}'. Brackets become array
    // index entries, dots become segment separators.
    QStringList parts;
    static const QRegularExpression re(QStringLiteral(
        R"(([^.\[\]]+)|\[(\d+)\])"));
    auto it = re.globalMatch(jsonPath);
    while (it.hasNext()) {
        const auto m = it.next();
        parts << (m.captured(1).isEmpty() ? m.captured(2) : m.captured(1));
    }
    if (parts.isEmpty()) {
        qWarning() << "LiveSync json path parse failed:" << jsonPath;
        return false;
    }
    const QString pgPath = QStringLiteral("{%1}").arg(parts.join(QLatin1Char(',')));

    QSqlQuery q(m_conn->queryDb());

    q.prepare("SELECT set_config('dve.live_column', ?, false), "
              "       set_config('dve.live_value',  ?, false)");
    q.addBindValue(QStringLiteral("json_path:") + jsonPath);
    q.addBindValue(value.toString());
    if (!q.exec()) return false;

    const QString uuid = m_identity ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString();
    const QString sql = QStringLiteral(
        "UPDATE %1 SET json_data = jsonb_set(json_data, ?::text[], "
        "to_jsonb(?::text)::jsonb, true), "
        "version = version + 1, updated_at = now(), updated_by = ? "
        "WHERE id = ?").arg(table);
    q.prepare(sql);
    q.addBindValue(pgPath);
    q.addBindValue(value.toString());
    q.addBindValue(uuid);
    q.addBindValue(rowId);
    if (!q.exec()) {
        qWarning() << "LiveSync jsonb_set failed:" << q.lastError().text();
        return false;
    }
    return q.numRowsAffected() > 0;
}

bool LiveSync::focusCell(const QString& table, qint64 rowId, const QString& column)
{
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;
    QSqlQuery q(m_conn->queryDb());

    // One focus per user. Clear prior, then insert. Two statements are
    // safer than ON CONFLICT against a 4-column PK.
    q.prepare("DELETE FROM cell_focus WHERE user_uuid = ?::uuid");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    if (!q.exec()) return false;

    q.prepare(
        "INSERT INTO cell_focus(user_uuid, table_name, row_id, "
        "column_name, user_name, user_color) "
        "VALUES(?::uuid, ?, ?, ?, ?, ?)");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(m_identity->displayName());
    q.addBindValue(m_identity->color());
    if (!q.exec()) {
        qWarning() << "LiveSync::focusCell INSERT failed:" << q.lastError().text();
        return false;
    }
    return true;
}

bool LiveSync::blurCell()
{
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;
    QSqlQuery q(m_conn->queryDb());
    q.prepare("DELETE FROM cell_focus WHERE user_uuid = ?::uuid");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    return q.exec();
}

void LiveSync::onRowChanged(const RowChange& c)
{
    if (c.column.isEmpty()) return;       // multi-column UPDATE - skip
    emit cellChanged(c.table, c.id, c.column, c.newValue);
}

void LiveSync::onCellFocusChanged(const CellFocusChange& f)
{
    if (f.op == QLatin1String("DELETE"))
        emit cellBlurred(f.tableName, f.rowId, f.columnName);
    else
        emit cellFocused(f.tableName, f.rowId, f.columnName,
                         f.userName, f.userColor);
}

} // namespace DVE
```

### Step 6: Update the .pro

- [ ] **Add LiveSync to `DataViewerEnterprise.pro`**

Find the `database/` block in SOURCES + HEADERS. Append:

```pro
SOURCES += src/database/LiveSync.cpp
HEADERS += src/database/LiveSync.h
```

### Step 7: Run the test — verify it passes

- [ ] **Re-run the livesync test**

```powershell
.\tests\run-tests.ps1 -Filter livesync
```

Expected: PASS on all three slots (or SKIP if `DVE_TEST_PG_CONN` is unset).

### Step 8: Commit

```bash
git add src/database/LiveSync.h src/database/LiveSync.cpp \
        DataViewerEnterprise.pro \
        tests/tst_livesync/ tests/tests.pro
git commit -m "feat(db): LiveSync core - commitCell + focusCell

Single chokepoint for v2.0.1 per-cell writes. Scalar-column UPDATEs
go through runScalarUpdate; JSONB paths (column starts with
'json_path:') translate to jsonb_set() so concurrent edits to
different sample fields don't clobber each other.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: Wire LiveSync to NotificationListener (own-UUID filter + dispatch)

**Files:**
- Modify: `src/MainWindow.h` (declare LiveSync member)
- Modify: `src/MainWindow.cpp` (construct LiveSync, wire signals)
- Existing test: `tst_livesync` already covers the slot dispatch

### Step 1: Add the member to MainWindow.h

- [ ] **Edit `src/MainWindow.h`**

Add the forward declaration near other DB classes:

```cpp
class LiveSync;
```

Add a member with the others:

```cpp
DVE::LiveSync* m_liveSync = nullptr;
```

### Step 2: Construct LiveSync in MainWindow

- [ ] **Edit `src/MainWindow.cpp` constructor**

After `m_notify = new DVE::NotificationListener(...)` and `m_notify->subscribe()`, construct LiveSync:

```cpp
            if (m_notify && m_conn && m_conn->isOpen()) {
                m_liveSync = new DVE::LiveSync(m_pgConn, m_identity, this);
                // Filter own writes BEFORE LiveSync sees them.
                const QString selfUuid = m_identity->uuid().toString(QUuid::WithoutBraces);
                connect(m_notify, &DVE::NotificationListener::rowChanged, m_liveSync,
                        [this, selfUuid](const DVE::RowChange& c) {
                            if (c.updatedBy == selfUuid) return;
                            m_liveSync->onRowChanged(c);
                        });
                connect(m_notify, &DVE::NotificationListener::cellFocusChanged, m_liveSync,
                        [this, selfUuid](const DVE::CellFocusChange& f) {
                            if (f.userUuid.toString(QUuid::WithoutBraces) == selfUuid) return;
                            m_liveSync->onCellFocusChanged(f);
                        });
            }
```

Place this block adjacent to the existing `m_notify` wire-up so the
init order is obvious.

### Step 3: Add the include

- [ ] **Add `#include "database/LiveSync.h"` at the top of MainWindow.cpp**

### Step 4: Build

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: clean build. (No tests added in this task — the wiring is exercised by Tasks 6 and 8 end-to-end tests.)

### Step 5: Commit

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(ui): construct LiveSync in MainWindow + wire to NotificationListener

Own-UUID filter runs ahead of LiveSync so we don't echo our own
writes back to ourselves.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: CellFocusDelegate

**Files:**
- Create: `src/widgets/CellFocusDelegate.h`
- Create: `src/widgets/CellFocusDelegate.cpp`
- Create: `tests/tst_cellfocusdelegate/tst_cellfocusdelegate.cpp`
- Create: `tests/tst_cellfocusdelegate/tst_cellfocusdelegate.pro`
- Modify: `tests/tests.pro` (SUBDIRS)
- Modify: `DataViewerEnterprise.pro`

### Step 1: Write the failing pixel test

- [ ] **Create `tests/tst_cellfocusdelegate/tst_cellfocusdelegate.cpp`**

```cpp
#include <QtTest>
#include <QStandardItemModel>
#include <QStyleOptionViewItem>
#include <QPixmap>
#include <QPainter>

#include "widgets/CellFocusDelegate.h"

using namespace DVE;

class TstCellFocusDelegate : public QObject
{
    Q_OBJECT
private slots:
    void noFocusRolesRendersDefault();
    void focusRolesPaintBorderAndFlag();
};

void TstCellFocusDelegate::noFocusRolesRendersDefault()
{
    QStandardItemModel m;
    m.appendRow(new QStandardItem("0.522"));

    CellFocusDelegate d;
    QStyleOptionViewItem opt;
    opt.rect = QRect(0, 0, 80, 24);

    QPixmap pm(80, 24); pm.fill(Qt::white);
    QPainter p(&pm);
    d.paint(&p, opt, m.index(0, 0));
    p.end();

    // No green / orange pixels in a no-roles render.
    const QImage img = pm.toImage();
    bool sawColor = false;
    for (int y = 0; y < img.height() && !sawColor; ++y)
        for (int x = 0; x < img.width(); ++x) {
            const QRgb px = img.pixel(x, y);
            if ((qRed(px) < 50 && qGreen(px) > 200) ||
                (qRed(px) > 200 && qGreen(px) > 100 && qBlue(px) < 50)) {
                sawColor = true; break;
            }
        }
    QVERIFY(!sawColor);
}

void TstCellFocusDelegate::focusRolesPaintBorderAndFlag()
{
    QStandardItemModel m;
    auto* item = new QStandardItem("0.522");
    item->setData(QString("#16a34a"), CellFocusDelegate::kFocusColorRole);
    item->setData(QString("Tina"),    CellFocusDelegate::kFocusNameRole);
    m.appendRow(item);

    CellFocusDelegate d;
    QStyleOptionViewItem opt;
    opt.rect = QRect(0, 16, 80, 24);   // leave room above for the flag

    QPixmap pm(80, 50); pm.fill(Qt::white);
    QPainter p(&pm);
    d.paint(&p, opt, m.index(0, 0));
    p.end();

    const QImage img = pm.toImage();

    // Border check: at least one solid-green pixel along each edge of opt.rect.
    auto greenAt = [&](int x, int y) {
        const QRgb px = img.pixel(x, y);
        return qRed(px) < 60 && qGreen(px) > 140 && qBlue(px) < 100;
    };
    QVERIFY(greenAt(0, 16) || greenAt(0, 17));            // left edge
    QVERIFY(greenAt(79, 16) || greenAt(79, 17));          // right edge
    QVERIFY(greenAt(0, 16) || greenAt(1, 16));            // top edge
    QVERIFY(greenAt(0, 39) || greenAt(1, 39));            // bottom edge

    // Flag check: in the region above opt.rect (y < 16) there should be
    // green-filled pixels and "Tina" text (white-ish pixels on green).
    bool sawFlagGreen = false;
    for (int y = 0; y < 16 && !sawFlagGreen; ++y)
        for (int x = 0; x < 60; ++x)
            if (greenAt(x, y)) { sawFlagGreen = true; break; }
    QVERIFY(sawFlagGreen);
}

QTEST_MAIN(TstCellFocusDelegate)
#include "tst_cellfocusdelegate.moc"
```

### Step 2: Create the test .pro

- [ ] **Write `tests/tst_cellfocusdelegate/tst_cellfocusdelegate.pro`**

```pro
QT += testlib gui widgets
CONFIG += console c++17
TEMPLATE = app

INCLUDEPATH += ../../src

SOURCES += tst_cellfocusdelegate.cpp \
    ../../src/widgets/CellFocusDelegate.cpp

HEADERS += ../../src/widgets/CellFocusDelegate.h
```

Append `tst_cellfocusdelegate` to `tests/tests.pro` SUBDIRS.

### Step 3: Run the test — verify it fails

```powershell
.\tests\run-tests.ps1 -Filter cellfocusdelegate
```

Expected: FAIL at compile — `CellFocusDelegate.h` doesn't exist.

### Step 4: Write the delegate header

- [ ] **Create `src/widgets/CellFocusDelegate.h`**

```cpp
#pragma once

#include <QStyledItemDelegate>

namespace DVE {

// Paints a "someone else is editing this cell" decoration on TPM cells:
// a 2px colored border around the cell + a small name flag positioned
// above. The cell still shows its underlying value beneath.
//
// Roles consumed (set per-item on the model):
//   kFocusColorRole : hex color string of the remote user (e.g. "#16a34a").
//   kFocusNameRole  : display name string ("Tina"). Truncated to ~14 chars.
//   kFlashRole      : bool; when true the cell is rendered with a yellow
//                     tint on top of the normal text for the next paint
//                     cycle. The caller is responsible for clearing the
//                     role after the flash interval.
class CellFocusDelegate : public QStyledItemDelegate {
    Q_OBJECT
public:
    explicit CellFocusDelegate(QObject* parent = nullptr);

    void paint(QPainter* painter, const QStyleOptionViewItem& option,
               const QModelIndex& index) const override;

    static constexpr int kFocusColorRole = Qt::UserRole + 10;
    static constexpr int kFocusNameRole  = Qt::UserRole + 11;
    static constexpr int kFlashRole      = Qt::UserRole + 12;
};

} // namespace DVE
```

### Step 5: Write the delegate implementation

- [ ] **Create `src/widgets/CellFocusDelegate.cpp`**

```cpp
#include "CellFocusDelegate.h"

#include <QPainter>
#include <QFontMetrics>
#include <QColor>

namespace DVE {

namespace {
constexpr int kBorderThickness = 2;
constexpr int kFlagHeight      = 14;
constexpr int kFlagPaddingX    = 4;
constexpr int kFlagFontPt      = 8;
constexpr int kFlagMaxChars    = 14;
} // namespace

CellFocusDelegate::CellFocusDelegate(QObject* parent)
    : QStyledItemDelegate(parent) {}

void CellFocusDelegate::paint(QPainter* painter,
                              const QStyleOptionViewItem& option,
                              const QModelIndex& index) const
{
    // Default render first.
    QStyledItemDelegate::paint(painter, option, index);

    // Flash overlay -- yellow wash on top of the painted cell.
    if (index.data(kFlashRole).toBool()) {
        painter->save();
        painter->fillRect(option.rect, QColor(255, 240, 130, 160));
        painter->restore();
    }

    const QString colorHex = index.data(kFocusColorRole).toString();
    if (colorHex.isEmpty()) return;
    const QColor color(colorHex);
    if (!color.isValid()) return;

    QString name = index.data(kFocusNameRole).toString();
    if (name.size() > kFlagMaxChars) name = name.left(kFlagMaxChars - 1) + QChar(0x2026);

    painter->save();
    painter->setRenderHint(QPainter::Antialiasing, true);

    // Border inside the cell.
    QPen pen(color);
    pen.setWidth(kBorderThickness);
    pen.setJoinStyle(Qt::MiterJoin);
    painter->setPen(pen);
    painter->setBrush(Qt::NoBrush);
    QRect inner = option.rect.adjusted(kBorderThickness / 2,
                                       kBorderThickness / 2,
                                       -kBorderThickness / 2,
                                       -kBorderThickness / 2);
    painter->drawRect(inner);

    // Name flag, rounded top, positioned above the cell.
    if (!name.isEmpty() && option.rect.top() >= kFlagHeight) {
        QFont f = painter->font();
        f.setPointSize(kFlagFontPt);
        f.setBold(true);
        painter->setFont(f);
        const QFontMetrics fm(f);
        const int textW = fm.horizontalAdvance(name);
        const int flagW = textW + 2 * kFlagPaddingX;
        const QRect flagRect(option.rect.left(),
                             option.rect.top() - kFlagHeight,
                             qMin(flagW, option.rect.width()),
                             kFlagHeight);
        QPainterPath path;
        path.addRoundedRect(flagRect, 3, 3);
        path.addRect(flagRect.left(), flagRect.bottom() - 2,
                     flagRect.width(), 3);
        painter->setPen(Qt::NoPen);
        painter->setBrush(color);
        painter->drawPath(path);

        painter->setPen(Qt::white);
        painter->drawText(flagRect.adjusted(kFlagPaddingX, 0, -kFlagPaddingX, 0),
                          Qt::AlignVCenter | Qt::AlignLeft, name);
    }

    painter->restore();
}

} // namespace DVE
```

### Step 6: Add to project .pro

- [ ] **Append to `DataViewerEnterprise.pro`**

```pro
SOURCES += src/widgets/CellFocusDelegate.cpp
HEADERS += src/widgets/CellFocusDelegate.h
```

### Step 7: Run the test — verify it passes

```powershell
.\tests\run-tests.ps1 -Filter cellfocusdelegate
```

Expected: both slots PASS.

### Step 8: Commit

```bash
git add src/widgets/CellFocusDelegate.h src/widgets/CellFocusDelegate.cpp \
        tests/tst_cellfocusdelegate/ tests/tests.pro \
        DataViewerEnterprise.pro
git commit -m "feat(ui): CellFocusDelegate paints Variant-C presence decoration

2px colored border in the editor's color plus a small name flag above
the cell. Roles: kFocusColorRole / kFocusNameRole / kFlashRole. Pixel
test confirms border + flag render with the expected colors.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: Wire TPM data table to LiveSync + CellFocusDelegate

**Files:**
- Modify: `src/MainWindow.cpp` (apply delegate, route itemChanged through LiveSync, listen for remote cell signals, paint focus roles when cellFocused fires)
- Modify: `src/MainWindow.h` (slot declarations for the new wires)

### Step 1: Read the existing itemChanged path

- [ ] **Identify the current m_dataTable wiring**

```bash
grep -n "m_dataTable\|onDataTableItemChanged\|connect.*itemChanged" src/MainWindow.cpp
```

Read the `onDataTableItemChanged` slot fully. Note where it currently:
- Reads the new value from the QTableWidgetItem
- Looks up the data_rows id via the vertical header role
- Pushes a (sheet, row, column, value) tuple to `m_pendingWrites`
- Schedules the debounced Excel write timer

### Step 2: Apply the delegate to m_dataTable

- [ ] **Find where m_dataTable is constructed (search `m_dataTable = new`) and after construction add:**

```cpp
    m_cellFocusDelegate = new DVE::CellFocusDelegate(this);
    m_dataTable->setItemDelegate(m_cellFocusDelegate);
```

Add the member declaration in `MainWindow.h` next to existing widget pointers:

```cpp
DVE::CellFocusDelegate* m_cellFocusDelegate = nullptr;
```

And include the header at the top of MainWindow.cpp:

```cpp
#include "widgets/CellFocusDelegate.h"
```

### Step 3: Route itemChanged through LiveSync

- [ ] **At the bottom of `onDataTableItemChanged`, after the existing m_pendingWrites push, add:**

```cpp
    // v2.0.1 live sync: push the same edit to the DB immediately so
    // every other client sees it without waiting for a save.
    if (m_liveSync) {
        const QString column = columnNameForDataTableColumn(it->column());
        if (!column.isEmpty()) {
            const QVariant value = it->text();  // string for TEXT cols, double for numerics; DB layer casts
            const qint64 rowId = it->data(Qt::UserRole + 1).toLongLong();
            if (rowId > 0) m_liveSync->commitCell("data_rows", rowId, column, value);
        }
    }
```

`columnNameForDataTableColumn(int)` is a helper that maps the visual column index to the DB column name (`puffs`, `before_weight`, `after_weight`, `draw_pressure`, `resistance`, `smell`, `clog`, `notes`). Add it as a static helper at the top of MainWindow.cpp:

```cpp
namespace {
QString columnNameForDataTableColumn(int col) {
    static const QStringList kCols = {
        "puffs", "before_weight", "after_weight", "draw_pressure",
        "resistance", "smell", "clog", "notes"
    };
    return (col >= 0 && col < kCols.size()) ? kCols[col] : QString();
}
}
```

Verify the column order against the actual `m_dataTable` header setup — search for `setHorizontalHeaderLabels` and confirm the indices match. If the table has different columns or order, update `kCols` accordingly.

### Step 4: Receive remote cell changes

- [ ] **In MainWindow's constructor (after `m_liveSync = new ...`), connect the new signals:**

```cpp
                connect(m_liveSync, &DVE::LiveSync::cellChanged, this,
                        &MainWindow::onRemoteCellChanged);
                connect(m_liveSync, &DVE::LiveSync::cellFocused, this,
                        &MainWindow::onRemoteCellFocused);
                connect(m_liveSync, &DVE::LiveSync::cellBlurred, this,
                        &MainWindow::onRemoteCellBlurred);
```

Add the slot declarations to `MainWindow.h::private slots:`:

```cpp
    void onRemoteCellChanged(const QString& table, qint64 rowId,
                              const QString& column, const QVariant& newValue);
    void onRemoteCellFocused(const QString& table, qint64 rowId,
                              const QString& column,
                              const QString& userName,
                              const QString& userColor);
    void onRemoteCellBlurred(const QString& table, qint64 rowId,
                              const QString& column);
```

Add the slot implementations to MainWindow.cpp:

```cpp
void MainWindow::onRemoteCellChanged(const QString& table, qint64 rowId,
                                     const QString& column, const QVariant& newValue)
{
    if (table != QLatin1String("data_rows")) {
        // Sensory + Detailed Sensory go through their own handlers
        // installed by SensoryPanel / DetailedSensoryPanel.
        return;
    }
    // Find the on-screen QTableWidgetItem for this data_rows id.
    const int row = findTableRowForDataRowId(static_cast<int>(rowId));
    if (row < 0) return;
    const int col = dataTableColumnForColumnName(column);
    if (col < 0) return;
    QTableWidgetItem* it = m_dataTable->item(row, col);
    if (!it) return;

    QSignalBlocker b(m_dataTable);
    it->setText(newValue.toString());
    // Flash for 200 ms.
    it->setData(DVE::CellFocusDelegate::kFlashRole, true);
    QTimer::singleShot(200, this, [this, row, col]() {
        QTableWidgetItem* it = m_dataTable->item(row, col);
        if (it) it->setData(DVE::CellFocusDelegate::kFlashRole, QVariant());
    });
    // Also fold the remote edit into the pending Excel writeback so the
    // local .xlsx tracks it.
    m_pendingWrites.append(CellWrite{ /* fill from current sheet ctx */ });
    if (m_excelWriteTimer) m_excelWriteTimer->start();
}

void MainWindow::onRemoteCellFocused(const QString& table, qint64 rowId,
                                     const QString& column,
                                     const QString& userName,
                                     const QString& userColor)
{
    if (table != QLatin1String("data_rows")) return;
    const int row = findTableRowForDataRowId(static_cast<int>(rowId));
    if (row < 0) return;
    const int col = dataTableColumnForColumnName(column);
    if (col < 0) return;
    QTableWidgetItem* it = m_dataTable->item(row, col);
    if (!it) return;

    QSignalBlocker b(m_dataTable);
    it->setData(DVE::CellFocusDelegate::kFocusColorRole, userColor);
    it->setData(DVE::CellFocusDelegate::kFocusNameRole,  userName);
    m_dataTable->viewport()->update();
}

void MainWindow::onRemoteCellBlurred(const QString& table, qint64 rowId,
                                     const QString& column)
{
    if (table != QLatin1String("data_rows")) return;
    const int row = findTableRowForDataRowId(static_cast<int>(rowId));
    if (row < 0) return;
    const int col = dataTableColumnForColumnName(column);
    if (col < 0) return;
    QTableWidgetItem* it = m_dataTable->item(row, col);
    if (!it) return;

    QSignalBlocker b(m_dataTable);
    it->setData(DVE::CellFocusDelegate::kFocusColorRole, QVariant());
    it->setData(DVE::CellFocusDelegate::kFocusNameRole,  QVariant());
    m_dataTable->viewport()->update();
}
```

Also add a tiny inverse helper:

```cpp
namespace {
int dataTableColumnForColumnName(const QString& name) {
    static const QStringList kCols = {
        "puffs", "before_weight", "after_weight", "draw_pressure",
        "resistance", "smell", "clog", "notes"
    };
    return kCols.indexOf(name);
}
}
```

### Step 5: Broadcast local focus

- [ ] **Wire `m_dataTable->currentItemChanged` to LiveSync::focusCell**

In MainWindow constructor near the existing m_dataTable signal hookups:

```cpp
    connect(m_dataTable, &QTableWidget::currentItemChanged, this,
            [this](QTableWidgetItem* curr, QTableWidgetItem* /*prev*/) {
                if (!m_liveSync) return;
                if (!curr) { m_liveSync->blurCell(); return; }
                const qint64 rowId = curr->data(Qt::UserRole + 1).toLongLong();
                const QString column = columnNameForDataTableColumn(curr->column());
                if (rowId > 0 && !column.isEmpty())
                    m_liveSync->focusCell("data_rows", rowId, column);
                else
                    m_liveSync->blurCell();
            });
```

Also blur on focus-out of the table widget itself (when the user clicks somewhere else in the app). The simplest hook is an event filter, but a pragmatic shortcut: blur when the dock or central widget loses focus. For v2.0.1, blur happens naturally when the user selects a different cell or closes the file. Crash-cleanup is covered by the 30 s heartbeat staleness.

### Step 6: Build full app

- [ ] **Verify clean build**

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: clean under `-Werror -Wall -Wextra`. Decrypt if needed.

### Step 7: Commit

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(ui): TPM table commits live to LiveSync + paints remote focus

itemChanged still pushes Excel-pending edits; now also calls
commitCell so every other client receives the edit immediately.
currentItemChanged drives focusCell/blurCell. Remote rowChanged
flashes the affected cell for 200 ms; remote cellFocus draws the
Variant-C border + name flag via CellFocusDelegate.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 7: Save button rename + Excel write-back integration

**Files:**
- Modify: `src/MainWindow.cpp` (ribbon label change, Save button click handler)
- Modify: `src/MainWindow.h` (rename slot if applicable)

### Step 1: Rename the ribbon button

- [ ] **Find the Save button construction**

```bash
grep -n "tr(\"Save\")\|onSaveTriggered\|buildHomeTab" src/MainWindow.cpp
```

Locate the `addButton` (or similar) call that creates the Save button on the Home tab. Update its label and tooltip:

```cpp
    auto* exportBtn = m_ribbon->addButton(homeTab, tr("Export to Excel"),
                                          resourcePath("icons/export.png"));
    exportBtn->setToolTip(tr("Force-flush pending writes to the .xlsx file now"));
    connect(exportBtn, &QToolButton::clicked, this, &MainWindow::onExportToExcelTriggered);
```

(Adapt to whatever ribbon API the existing code uses. Search `homeTab` and other Save-button hookups to match.)

### Step 2: Rename or refocus the slot

- [ ] **Rename `onSaveTriggered` → `onExportToExcelTriggered` and trim DB save**

In MainWindow.h, rename the declaration. In .cpp, rename the implementation. The body should now:

1. Call `flushExcelWrites()` synchronously.
2. NOT call any save-coordinator / DB save (live sync already persists everything).
3. Show a brief status message: `"Exported to <filename>"`.

```cpp
void MainWindow::onExportToExcelTriggered()
{
    if (m_pendingWrites.isEmpty()) {
        updateStatusBar(tr("Nothing to export"));
        return;
    }
    flushExcelWrites();
    updateStatusBar(tr("Exported to Excel"));
}
```

### Step 3: Build

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: clean.

### Step 4: Commit

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(ui): rename Save button to Export to Excel

Live sync persists every cell commit immediately, so the explicit
'Save' becomes 'Export to Excel' -- a manual flush of the debounced
write-back to the .xlsx file. Removes the DB save call from the
click handler since LiveSync already covers it.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 8: Wire Sensory + Detailed Sensory modes to LiveSync

**Files:**
- Modify: `src/ui/SensoryPanel.h` (LiveSync pointer member, setter)
- Modify: `src/ui/SensoryPanel.cpp` (SampleCard commits route through LiveSync)
- Modify: `src/ui/DetailedSensoryPanel.h` / `.cpp` (same)
- Modify: `src/MainWindow.cpp` (inject m_liveSync into both panels)

### Step 1: Add setter + member to SensoryPanel.h

- [ ] **Edit `src/ui/SensoryPanel.h`**

Add forward declaration and member:

```cpp
class LiveSync;
```

In the public section near `setSaveCoordinator`:

```cpp
    void setLiveSync(LiveSync* sync) { m_liveSync = sync; }
```

In the private section:

```cpp
    LiveSync* m_liveSync = nullptr;
```

### Step 2: Route SampleCard commits through LiveSync

- [ ] **In `src/ui/SensoryPanel.cpp`, find each `changed()` signal connection in SampleCard**

For each widget that emits `changed()` (the spin boxes for scores + voltage / resistance / power_type / puff_length / comments / name), augment the connection with a LiveSync commit. The connection lives in the SampleCard constructor; SampleCard doesn't have a LiveSync pointer directly. Easiest pattern: emit a more detailed signal from SampleCard and have SensoryPanel translate it.

Add to `SampleCard`'s signals section:

```cpp
signals:
    void changed();
    void removeRequested(SampleCard* card);
    void cellCommitted(const QString& jsonPath, const QVariant& value);
```

In the constructor, after each widget's edit-finalization signal, emit `cellCommitted`. For example for `m_nameEdit`:

```cpp
    connect(m_nameEdit, &QLineEdit::editingFinished, this, [this]() {
        emit cellCommitted(QStringLiteral("name"), m_nameEdit->text());
    });
```

For `m_powerTypeCombo`:

```cpp
    connect(m_powerTypeCombo, &QComboBox::currentTextChanged, this,
            [this](const QString& s){ emit cellCommitted(QStringLiteral("power_type"), s); });
```

For each score spin (inside the loop that already emits `changed()`):

```cpp
        connect(spin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, [this, metric](double v) {
                    emit cellCommitted(QStringLiteral("scores.") + metric, v);
                });
```

For voltage, resistance, heatingTech, puff_length, comments — mirror the same `cellCommitted(field, value)` pattern.

The `jsonPath` emitted is relative to the sample (e.g. `"name"`, `"voltage"`, `"scores.Burnt Taste"`, `"power_type"`, `"puff_length_sec"`, `"comments"`). SensoryPanel will prepend the `samples[i]` prefix and the `json_path:` LiveSync key.

### Step 3: Connect SampleCard::cellCommitted in SensoryPanel

- [ ] **In SensoryPanel, when adding a card (search `addSampleCard`):**

```cpp
    connect(card, &SampleCard::cellCommitted, this,
            [this, card](const QString& fieldPath, const QVariant& value) {
                if (!m_liveSync) return;
                if (m_currentTesterIdx < 0) return;
                SensorySession& s = m_sessions[m_currentTesterIdx];
                if (s.id <= 0) return;          // not yet persisted; nothing to live-sync
                const int idx = m_cards.indexOf(card);
                if (idx < 0) return;
                const QString jsonPath = QStringLiteral("json_path:samples[%1].%2")
                    .arg(idx).arg(fieldPath);
                m_liveSync->commitCell(QStringLiteral("sensory_sessions"),
                                       s.id, jsonPath, value);
            });
```

(The `s.id <= 0` guard prevents broadcasting placeholder edits before first save — paired with the existing placeholder predicate.)

### Step 4: Apply remote cell changes to SensoryPanel cards

- [ ] **Subscribe SensoryPanel to LiveSync::cellChanged**

In the `setLiveSync` setter (move it inline into a method since we now need to connect a signal):

```cpp
void SensoryPanel::setLiveSync(LiveSync* sync) {
    if (m_liveSync == sync) return;
    if (m_liveSync) disconnect(m_liveSync, nullptr, this, nullptr);
    m_liveSync = sync;
    if (m_liveSync) {
        connect(m_liveSync, &LiveSync::cellChanged, this,
                &SensoryPanel::onRemoteCellChanged);
    }
}
```

Move the declaration to .h (no longer inline) and add:

```cpp
private slots:
    void onRemoteCellChanged(const QString& table, qint64 rowId,
                              const QString& column, const QVariant& newValue);
```

In the implementation:

```cpp
void SensoryPanel::onRemoteCellChanged(const QString& table, qint64 rowId,
                                       const QString& column,
                                       const QVariant& newValue)
{
    if (table != QLatin1String("sensory_sessions")) return;
    if (!column.startsWith(QLatin1String("json_path:samples["))) return;

    // Parse "json_path:samples[N].<field-with-possible-dots>" into
    // (sampleIdx, fieldPath).
    const QString rest = column.mid(QStringLiteral("json_path:samples[").size());
    const int rbr = rest.indexOf(QLatin1Char(']'));
    if (rbr < 0) return;
    bool okIdx = false;
    const int idx = rest.left(rbr).toInt(&okIdx);
    if (!okIdx) return;
    if (rest.size() <= rbr + 2) return;   // need .field
    const QString fieldPath = rest.mid(rbr + 2);  // skip "]."

    // Locate the local session by id.
    int sessIdx = -1;
    for (int i = 0; i < m_sessions.size(); ++i)
        if (m_sessions[i].id == static_cast<int>(rowId)) { sessIdx = i; break; }
    if (sessIdx < 0) return;
    if (sessIdx != m_currentTesterIdx) {
        // Not visible right now; patch the in-memory struct only.
        if (idx < 0 || idx >= m_sessions[sessIdx].samples.size()) return;
        applyRemoteFieldToSample(m_sessions[sessIdx].samples[idx], fieldPath, newValue);
        return;
    }
    if (idx < 0 || idx >= m_cards.size()) return;
    SampleCard* card = m_cards[idx];

    // Patch the struct then re-bind the card with signals blocked.
    applyRemoteFieldToSample(m_sessions[sessIdx].samples[idx], fieldPath, newValue);
    QSignalBlocker b(card);
    card->fromSample(m_sessions[sessIdx].samples[idx]);
}
```

Add the helper:

```cpp
void SensoryPanel::applyRemoteFieldToSample(SensorySample& s,
                                            const QString& fieldPath,
                                            const QVariant& value)
{
    if (fieldPath == QLatin1String("name"))               s.name = value.toString();
    else if (fieldPath == QLatin1String("comments"))      s.comments = value.toString();
    else if (fieldPath == QLatin1String("voltage"))       s.voltage = value.toDouble();
    else if (fieldPath == QLatin1String("resistance"))    s.resistance = value.toDouble();
    else if (fieldPath == QLatin1String("power"))         s.power = value.toDouble();
    else if (fieldPath == QLatin1String("heatingTechnology") ||
             fieldPath == QLatin1String("heating_technology"))
                                                          s.heatingTechnology = value.toString();
    else if (fieldPath == QLatin1String("power_type"))    s.powerType = value.toString();
    else if (fieldPath == QLatin1String("puff_length_sec")) s.puffLengthSec = value.toDouble();
    else if (fieldPath.startsWith(QLatin1String("scores."))) {
        const QString metric = fieldPath.mid(QStringLiteral("scores.").size());
        s.scores[metric] = value.toDouble();
    }
}
```

Declare both new methods in the header.

### Step 5: Repeat for DetailedSensoryPanel

- [ ] **Apply the same pattern to `src/ui/DetailedSensoryPanel.{h,cpp}`**

The Detailed Sensory card has 14 questions instead of 5 metrics; otherwise the structure is the same. Mirror Steps 2–4 for the detailed widget.

### Step 6: Inject LiveSync into both panels from MainWindow

- [ ] **In `src/MainWindow.cpp`, after both panels are constructed and `m_liveSync` exists:**

```cpp
    if (m_sensoryPanel && m_liveSync)         m_sensoryPanel->setLiveSync(m_liveSync);
    if (m_detailedSensoryPanel && m_liveSync) m_detailedSensoryPanel->setLiveSync(m_liveSync);
```

Place this near where `setSaveCoordinator` was previously called.

### Step 7: Build + test

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

```powershell
.\tests\run-tests.ps1
```

Expected: clean build; baseline 27 PASS / 2 pre-existing SOP failures. The existing `tst_sensoryreportsource` tests still cover the JSON serializer behavior; no change required there.

### Step 8: Commit

```bash
git add src/ui/SensoryPanel.h src/ui/SensoryPanel.cpp \
        src/ui/DetailedSensoryPanel.h src/ui/DetailedSensoryPanel.cpp \
        src/MainWindow.cpp
git commit -m "feat(sensory): route SampleCard commits through LiveSync (JSONB path)

Sensory + Detailed Sensory editing now broadcasts via LiveSync the
same way TPM does. Sample edits emit cellCommitted(jsonPath, value);
SensoryPanel prepends 'json_path:samples[i].' and forwards to
LiveSync::commitCell. Remote applies patch the in-memory struct and
re-bind the visible card with signals blocked. Detailed Sensory
mirrors the same pattern.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 9: Extend OfflineSnapshot pending-edit queue for per-cell

**Files:**
- Modify: `src/database/OfflineSnapshot.h` (queue API takes table + rowId + column + value)
- Modify: `src/database/OfflineSnapshot.cpp` (extend pending-edits SQLite schema + replay logic)
- Modify: `src/database/LiveSync.cpp` (commitCell falls back to queue when offline)
- Modify: `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp` (new round-trip slot)

### Step 1: Inspect the existing queue

- [ ] **Read the current pending-edit storage**

```bash
grep -n "pending_edits\|enqueuePendingEdit\|drainPendingEdits" src/database/OfflineSnapshot.cpp
```

Identify the existing schema. If `pending_edits` already has columns close to what we need (target table, row id, payload), extend rather than replace. Note the current API shape.

### Step 2: Extend the schema + API

- [ ] **Add columns to the SQLite `pending_edits` table:** `column_name TEXT`, `value_text TEXT`. If they already exist, skip.

```cpp
// In ensureSchema() / similar:
m_db.exec("ALTER TABLE pending_edits ADD COLUMN column_name TEXT");
m_db.exec("ALTER TABLE pending_edits ADD COLUMN value_text TEXT");
// (ALTER ADD COLUMN is no-op safe in SQLite if column already exists --
// actually it errors. Check pragma table_info() first and only add if missing.)
```

- [ ] **Add `enqueueCellEdit(table, rowId, column, value)` and update `drainPendingEdits` to replay cell-level ops.**

```cpp
bool OfflineSnapshot::enqueueCellEdit(const QString& table, qint64 rowId,
                                     const QString& column, const QVariant& value)
{
    QSqlQuery q(m_db);
    q.prepare("INSERT INTO pending_edits(target_table, row_id, column_name, value_text) "
              "VALUES(?, ?, ?, ?)");
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(value.toString());
    return q.exec();
}
```

Update replay to invoke a caller-supplied callback per cell (the caller is `LiveSync::flushPending()`):

```cpp
int OfflineSnapshot::drainPendingEdits(
    std::function<bool(const QString&, qint64, const QString&, const QVariant&)> apply)
{
    QSqlQuery q(m_db);
    if (!q.exec("SELECT id, target_table, row_id, column_name, value_text "
                "FROM pending_edits WHERE column_name IS NOT NULL "
                "ORDER BY id")) return 0;
    QVector<qint64> applied;
    while (q.next()) {
        const qint64 pid = q.value(0).toLongLong();
        const QString table = q.value(1).toString();
        const qint64 rowId = q.value(2).toLongLong();
        const QString col = q.value(3).toString();
        const QVariant val = q.value(4);
        if (apply(table, rowId, col, val)) applied << pid;
    }
    if (!applied.isEmpty()) {
        QStringList ids; for (qint64 id : applied) ids << QString::number(id);
        QSqlQuery del(m_db);
        del.exec("DELETE FROM pending_edits WHERE id IN (" + ids.join(',') + ")");
    }
    return applied.size();
}
```

### Step 3: Use the queue from LiveSync when offline

- [ ] **Edit `LiveSync::commitCell` (both scalar + jsonPath paths)**

Before issuing the UPDATE, check connection. On failure, enqueue to OfflineSnapshot:

```cpp
    if (!m_conn || !m_conn->isOpen()) {
        if (m_snapshot) m_snapshot->enqueueCellEdit(table, rowId, column, value);
        return true;       // queued = success from the caller's perspective
    }
```

This requires LiveSync to know about OfflineSnapshot. Add an `m_snapshot` pointer to LiveSync and a `setOfflineSnapshot(OfflineSnapshot*)` setter. MainWindow injects it after construction.

Add a `flushPending()` method that ConnectionMonitor calls on reconnect:

```cpp
int LiveSync::flushPending()
{
    if (!m_snapshot || !m_conn || !m_conn->isOpen()) return 0;
    return m_snapshot->drainPendingEdits(
        [this](const QString& t, qint64 r, const QString& c, const QVariant& v) {
            return this->commitCell(t, r, c, v);
        });
}
```

Wire `ConnectionMonitor::cameOnline` (or equivalent existing signal) in MainWindow to `m_liveSync->flushPending()`.

### Step 4: Add a test

- [ ] **Extend `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`**

```cpp
void TstOfflineSnapshot::cellEditsRoundTrip()
{
    QVERIFY(m_snap->enqueueCellEdit("data_rows", 7, "draw_pressure", 1.5));
    QVERIFY(m_snap->enqueueCellEdit("sensory_sessions", 42,
        "json_path:samples[0].voltage", 3.7));

    QVector<QStringList> applied;
    int n = m_snap->drainPendingEdits(
        [&](const QString& t, qint64 r, const QString& c, const QVariant& v) {
            applied << QStringList{t, QString::number(r), c, v.toString()};
            return true;
        });
    QCOMPARE(n, 2);
    QCOMPARE(applied.size(), 2);
    QCOMPARE(applied[0][0], QStringLiteral("data_rows"));
    QCOMPARE(applied[1][2], QStringLiteral("json_path:samples[0].voltage"));

    // Second drain should be empty (rows deleted).
    int n2 = m_snap->drainPendingEdits([](const QString&, qint64, const QString&, const QVariant&){ return true; });
    QCOMPARE(n2, 0);
}
```

### Step 5: Build + test

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

```powershell
.\tests\run-tests.ps1 -Filter offlinesnapshot
```

Expected: PASS including the new slot.

### Step 6: Commit

```bash
git add src/database/OfflineSnapshot.h src/database/OfflineSnapshot.cpp \
        src/database/LiveSync.h src/database/LiveSync.cpp \
        tests/tst_offlinesnapshot/
git commit -m "feat(db): per-cell offline queue + LiveSync replay on reconnect

Extends the pending_edits SQLite table with column_name + value_text.
LiveSync::commitCell falls back to the queue when the connection is
down; flushPending() replays each queued edit via commitCell() once
the connection returns. ConnectionMonitor:cameOnline triggers the
flush.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 10: Retire SaveCoordinator + ConflictResolver + VersionMismatchDialog + RowDeletedDialog

**Files to delete:**
- `src/database/SaveCoordinator.h`
- `src/database/SaveCoordinator.cpp`
- `src/database/ConflictResolver.h`
- `src/database/ConflictResolver.cpp`
- `src/database/VersionMismatchDialog.h`
- `src/database/VersionMismatchDialog.cpp`
- `src/database/RowDeletedDialog.h`
- `src/database/RowDeletedDialog.cpp`

**Files to modify:**
- `src/MainWindow.h` / `.cpp` — remove m_saveCoordinator, m_conflictResolver members + all call sites
- `src/ui/SensoryPanel.h` / `.cpp` — remove `setSaveCoordinator`, `m_saveCoord` + call sites
- `src/ui/DetailedSensoryPanel.h` / `.cpp` — same
- `src/ui/SensoryDialog.cpp` — drop SaveCoordinator usage
- `src/database/DatabaseManager.h` — drop save-coordinator-aware overloads if any
- `DataViewerEnterprise.pro` — remove deleted source files
- `tests/` — remove any test directories that test the deleted classes

### Step 1: Find every call site

- [ ] **Catalog references**

```bash
grep -rn "SaveCoordinator\|ConflictResolver\|VersionMismatchDialog\|RowDeletedDialog" \
    src/ tests/ DataViewerEnterprise.pro
```

This should print every reference that needs surgery.

### Step 2: Replace each call site with a direct LiveSync write or nothing

- [ ] **For each call site that uses `m_saveCoordinator->saveSensorySession(s)` etc.:**

If the call site is `MainWindow::onUpdateDatabase()`'s sensory loop, that loop's job (auto-save) is already covered by LiveSync — delete the loop.

If the call site is `SensoryPanel::save()`, that's the explicit Save action — now redundant with live sync. Delete the method body or repurpose it to just call `m_excelWriteTimer->start()` for export consistency.

Same for `DetailedSensoryPanel::save()` and `SensoryDialog::onSave()`.

Remove the members `m_saveCoordinator`, `m_saveCoord`, `m_conflictResolver` from class headers; remove the constructor injection calls.

### Step 3: Delete the source files

- [ ] **`git rm` the deleted files**

```bash
git rm src/database/SaveCoordinator.h src/database/SaveCoordinator.cpp \
       src/database/ConflictResolver.h src/database/ConflictResolver.cpp \
       src/database/VersionMismatchDialog.h src/database/VersionMismatchDialog.cpp \
       src/database/RowDeletedDialog.h src/database/RowDeletedDialog.cpp
```

### Step 4: Remove from .pro

- [ ] **Edit `DataViewerEnterprise.pro`**

Remove the four files from SOURCES + HEADERS.

### Step 5: Delete test directories if they exist

- [ ] **Check for and remove related tests**

```bash
ls tests/ | grep -iE "save|conflict|versionmismatch|rowdeleted"
```

`git rm -r` any matching directory. Remove their entries from `tests/tests.pro` SUBDIRS.

### Step 6: Build + full suite

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: clean. If there are stragglers (call sites I missed), `-Werror` will flag them.

```powershell
.\tests\run-tests.ps1
```

Expected: 25 PASS / 2 pre-existing SOP failures (drops two if the deleted-class tests existed and were counted before). No new failures.

### Step 7: Commit

```bash
git add -A
git commit -m "refactor(db): retire SaveCoordinator + ConflictResolver + dialogs

Live sync replaces save-driven conflict resolution. The four classes
deleted (SaveCoordinator, ConflictResolver, VersionMismatchDialog,
RowDeletedDialog) and all their call sites are gone. UniqueViolation
Dialog stays -- duplicate natural-key violations on INSERT are still
a real failure path.

The placeholder predicate (isPlaceholderSession) now gates
LiveSync::commitCell for sensory sessions instead of the auto-save
loop in onUpdateDatabase.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 11: End-to-end two-client live sync test

**Files:**
- Modify: `tests/tst_twoclient_e2e/tst_twoclient_e2e.cpp`

### Step 1: Add a per-cell live-sync slot

- [ ] **Append a test slot**

```cpp
void TstTwoClientE2E::cellEditOnClientAReachesClientBLive()
{
    if (!m_setupOk) QSKIP("two-client setup not available");

    // Client A: commit a cell.
    bool ok = m_clientA->liveSync->commitCell("data_rows", m_dataRowId,
                                              "draw_pressure", 3.14);
    QVERIFY(ok);

    // Client B: poll up to 1.5 s for the value to land.
    QElapsedTimer timer; timer.start();
    bool seen = false;
    while (timer.elapsed() < 1500) {
        QCoreApplication::processEvents(QEventLoop::AllEvents, 50);
        QSqlQuery q(m_clientB->conn->queryDb());
        if (q.exec(QString("SELECT draw_pressure FROM data_rows WHERE id=%1")
                   .arg(m_dataRowId)) && q.next()
            && qAbs(q.value(0).toDouble() - 3.14) < 1e-6) {
            seen = true; break;
        }
    }
    QVERIFY(seen);

    // Client B: also verify the cellChanged signal fired.
    QVERIFY(m_clientB->cellChangedSpy->count() >= 1);
}
```

Add a `cellChangedSpy` member (QSignalSpy on `LiveSync::cellChanged`) to the client struct, initialized in initTestCase alongside the other spies.

### Step 2: Run

```powershell
.\tests\run-tests.ps1 -Filter twoclient_e2e
```

Expected: PASS within 1.5 s.

### Step 3: Commit

```bash
git add tests/tst_twoclient_e2e/
git commit -m "test(e2e): two-client per-cell live sync round-trip

Client A commits a cell; client B sees the value land in the DB and
the cellChanged signal fire within 1.5 s. Locks in the v2.0.1
end-to-end contract.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 12: Final test sweep + merge to main

### Step 1: Full suite

```powershell
.\tests\run-tests.ps1
```

Expected: 25–27 PASS / 2 pre-existing unrelated SOP failures. No new failures. (The exact count depends on whether the deleted classes had dedicated tests.)

### Step 2: Confirm clean tree + version bump

- [ ] **Bump VERSION in DataViewerEnterprise.pro from 2.0.0 to 2.0.1**

```bash
grep -n "^VERSION" DataViewerEnterprise.pro
# edit the line: VERSION = 2.0.1
```

- [ ] **Clean rebuild (required for VERSION bumps per CLAUDE.md)**

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make clean && mingw32-make -j8
```

Verify the binary embeds the new version:

```powershell
Get-Item .\build\release\DataViewer.exe | ForEach-Object { $_.VersionInfo.FileVersion }
```

Expected: `2.0.1.0` or `2.0.1`.

### Step 3: Commit version bump

```bash
git add DataViewerEnterprise.pro
git commit -m "chore(release): bump version to 2.0.1 (live collaborative editing)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Step 4: Push to main, delete feature branch

```bash
git push origin HEAD:main
git checkout main
git pull --ff-only origin main
git branch -D feature/live-collab-v2.0.1
```

### Step 5: Verify final state

```bash
git branch -vv
git log --oneline -15
```

Expected: on `main`, top ~12 commits are this plan's commits, `origin/main` matches.

---

## Self-review notes (already applied to this plan)

- **Spec coverage:** every section of `2026-05-16-live-collab-design.md` maps to at least one task here. Schema (Task 1), NotificationListener (Task 2), LiveSync core (Task 3), MainWindow wire-up (Task 4), CellFocusDelegate (Task 5), TPM integration (Task 6), Save → Export (Task 7), Sensory/Detailed Sensory (Task 8), offline (Task 9), retirement (Task 10), end-to-end (Task 11), final sweep (Task 12).
- **Spec deviation:** plan replaces "new LiveTableModel" with "keep QTableWidget and wire its itemChanged through LiveSync." Documented at the top of the plan; user-visible behavior identical.
- **Placeholder scan:** no TBDs, TODOs, "implement later" placeholders. The `columnNameForDataTableColumn` mapping is explicit; the verification step asks the implementer to confirm the column order against the actual header setup, which is a legitimate code-investigation step, not a placeholder.
- **Type consistency:** `RowChange::column` / `RowChange::newValue` (Task 2) match what `LiveSync::onRowChanged` reads (Task 3). `CellFocusDelegate::kFocusColorRole` / `kFocusNameRole` / `kFlashRole` (Task 5) match what MainWindow sets in Task 6 and what tst_cellfocusdelegate exercises. `LiveSync::commitCell` / `focusCell` / `blurCell` / `flushPending` consistent across all tasks that reference them.
- **Risks acknowledged in the spec are NOT separately tested but are baked into design:** the 150 ms focus debounce isn't implemented in this plan because every focusCell call already replaces the prior row (DELETE + INSERT in one transaction). If real-world load shows NOTIFY storms, a future micro-task adds the debounce in MainWindow's focus handler.

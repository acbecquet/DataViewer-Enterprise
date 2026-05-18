// Postgres-backed unit tests for the v2.0.2 stored-function migration.
// Targets deploy/postgres/migrations/2026-05-17-v2.0.2-fixes.sql.
//
// Findings covered:
//   C1 - dve_commit_cell / dve_commit_cell_json reject stale-version writes
//        and accept matching-version writes; NULL p_expected_version is the
//        v2.0.1 no-OCC legacy path.
//   M4 - dve_focus_cell scopes its DELETE to the user's other focus cells.
//   M5 - pg_cron job dve_cell_focus_cleanup uses a valid 5-field expression.
//   M6 - version bump owned by bump_version trigger (single +1 per UPDATE).
//   H4 - dve_commit_session_layout returns new version on success, NULL on
//        stale-version mismatch.

#include <QtTest/QtTest>
#include <QSqlQuery>
#include <QSqlError>
#include <QCoreApplication>

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

class TstStoredFns : public QObject
{
    Q_OBJECT
private slots:
    void initTestCase();
    void cleanupTestCase();
    void init();

    void commitCell_acceptsCorrectVersion();
    void commitCell_rejectsStaleVersion();
    void commitCell_nullExpectedVersionIsLegacyNoOcc();
    void commitCellJson_acceptsCorrectVersion();
    void commitCellJson_rejectsStaleVersion();
    void focusCell_preservesOtherFocusRows();
    void focusCell_refocusUpdatesStartedAt();
    void cronJob_cellFocusCleanupExists();
    void bumpVersion_singleIncrementPerUpdate();
    void commitSessionLayout_returnsNewVersionOnSuccess();
    void commitSessionLayout_returnsNullOnStaleVersion();

private:
    PostgresConnection* m_conn = nullptr;
    qint64              m_dataRowId       = -1;
    qint64              m_sensoryId       = -1;

    int    rowVersion(const QString& table, qint64 id);
    double drawPressure(qint64 id);
};

int TstStoredFns::rowVersion(const QString& table, qint64 id) {
    QSqlQuery q(m_conn->queryDb());
    if (!q.exec(QString("SELECT version FROM %1 WHERE id=%2")
                .arg(table).arg(id))) return -1;
    if (!q.next()) return -1;
    return q.value(0).toInt();
}

double TstStoredFns::drawPressure(qint64 id) {
    QSqlQuery q(m_conn->queryDb());
    q.exec(QString("SELECT draw_pressure FROM data_rows WHERE id=%1").arg(id));
    if (!q.next()) return -1.0;
    return q.value(0).toDouble();
}

void TstStoredFns::initTestCase()
{
    if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
        QSKIP("DVE_TEST_PG_CONN not set; skipping");
    QCoreApplication::setOrganizationName("DataViewerTest");
    QCoreApplication::setApplicationName("tst_storedfns");

    m_conn = new PostgresConnection(this);
    QVERIFY(m_conn->open(pgConfig()));

    // Verify the migration is applied: dve_commit_cell now has 6 args.
    QSqlQuery probe(m_conn->queryDb());
    QVERIFY2(probe.exec(
        "SELECT pronargs FROM pg_proc WHERE proname='dve_commit_cell'"),
        qPrintable(probe.lastError().text()));
    QVERIFY(probe.next());
    QCOMPARE(probe.value(0).toInt(), 6);

    QSqlQuery q(m_conn->queryDb());
    q.exec("DELETE FROM files WHERE file_path='/tmp/tst_storedfns.xlsx'");
    q.exec("DELETE FROM sensory_sessions WHERE session_name='tst_storedfns'");
    q.exec("DELETE FROM cell_focus");
}

void TstStoredFns::cleanupTestCase()
{
    if (!m_conn || !m_conn->isOpen()) return;
    QSqlQuery q(m_conn->queryDb());
    q.exec("DELETE FROM files WHERE file_path='/tmp/tst_storedfns.xlsx'");
    q.exec("DELETE FROM sensory_sessions WHERE session_name='tst_storedfns'");
    q.exec("DELETE FROM cell_focus");
}

void TstStoredFns::init()
{
    // Fresh fixture rows per test so cross-test version state can't bleed.
    QSqlQuery q(m_conn->queryDb());
    q.exec("DELETE FROM files WHERE file_path='/tmp/tst_storedfns.xlsx'");
    q.exec("DELETE FROM sensory_sessions WHERE session_name='tst_storedfns'");
    q.exec("DELETE FROM cell_focus");

    QVERIFY(q.exec("INSERT INTO files(file_path, file_name, loaded_at) "
                   "VALUES('/tmp/tst_storedfns.xlsx', 't.xlsx', '2026-05-17') "
                   "RETURNING id"));
    QVERIFY(q.next());
    const qint64 fileId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString("INSERT INTO tests(file_id, sheet_name) "
                           "VALUES(%1, 'S1') RETURNING id").arg(fileId)));
    QVERIFY(q.next());
    const qint64 testId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString("INSERT INTO samples(test_id, sample_name) "
                           "VALUES(%1, 'A') RETURNING id").arg(testId)));
    QVERIFY(q.next());
    const qint64 sampleId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString(
        "INSERT INTO data_rows(sample_id, sort_order, draw_pressure) "
        "VALUES(%1, 0, 1.0) RETURNING id").arg(sampleId)));
    QVERIFY(q.next());
    m_dataRowId = q.value(0).toLongLong();

    QVERIFY(q.exec(
        "INSERT INTO sensory_sessions(session_name, json_data, layout_json) "
        "VALUES('tst_storedfns', "
        "'{\"samples\":[{\"name\":\"X\"}]}'::jsonb, "
        "'{\"v\":1}'::jsonb) RETURNING id"));
    QVERIFY(q.next());
    m_sensoryId = q.value(0).toLongLong();
}

void TstStoredFns::commitCell_acceptsCorrectVersion()
{
    const int v0 = rowVersion("data_rows", m_dataRowId);
    QVERIFY(v0 > 0);

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_commit_cell(?, ?, ?, ?, ?, ?)");
    q.addBindValue("data_rows");
    q.addBindValue(m_dataRowId);
    q.addBindValue("draw_pressure");
    q.addBindValue("3.14");
    q.addBindValue(QString("00000000-0000-0000-0000-000000000001"));
    q.addBindValue(v0);
    QVERIFY2(q.exec(), qPrintable(q.lastError().text()));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toBool(), true);

    QCOMPARE(drawPressure(m_dataRowId), 3.14);
    QCOMPARE(rowVersion("data_rows", m_dataRowId), v0 + 1);
}

void TstStoredFns::commitCell_rejectsStaleVersion()
{
    const int v0 = rowVersion("data_rows", m_dataRowId);
    const double before = drawPressure(m_dataRowId);

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_commit_cell(?, ?, ?, ?, ?, ?)");
    q.addBindValue("data_rows");
    q.addBindValue(m_dataRowId);
    q.addBindValue("draw_pressure");
    q.addBindValue("9.99");
    q.addBindValue(QString("00000000-0000-0000-0000-000000000001"));
    q.addBindValue(v0 - 1);  // stale
    QVERIFY(q.exec());
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toBool(), false);

    QCOMPARE(drawPressure(m_dataRowId), before);
    QCOMPARE(rowVersion("data_rows", m_dataRowId), v0);
}

void TstStoredFns::commitCell_nullExpectedVersionIsLegacyNoOcc()
{
    const int v0 = rowVersion("data_rows", m_dataRowId);

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_commit_cell(?, ?, ?, ?, ?, NULL)");
    q.addBindValue("data_rows");
    q.addBindValue(m_dataRowId);
    q.addBindValue("draw_pressure");
    q.addBindValue("2.71");
    q.addBindValue(QString("00000000-0000-0000-0000-000000000001"));
    QVERIFY(q.exec());
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toBool(), true);

    QCOMPARE(drawPressure(m_dataRowId), 2.71);
    QCOMPARE(rowVersion("data_rows", m_dataRowId), v0 + 1);
}

void TstStoredFns::commitCellJson_acceptsCorrectVersion()
{
    const int v0 = rowVersion("sensory_sessions", m_sensoryId);

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_commit_cell_json(?, ?, ?, ?::text[], ?, ?, ?)");
    q.addBindValue("sensory_sessions");
    q.addBindValue(m_sensoryId);
    q.addBindValue("samples[0].name");
    q.addBindValue(QString("{samples,0,name}"));
    q.addBindValue("Renamed");
    q.addBindValue(QString("00000000-0000-0000-0000-000000000001"));
    q.addBindValue(v0);
    QVERIFY2(q.exec(), qPrintable(q.lastError().text()));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toBool(), true);
    QCOMPARE(rowVersion("sensory_sessions", m_sensoryId), v0 + 1);
}

void TstStoredFns::commitCellJson_rejectsStaleVersion()
{
    const int v0 = rowVersion("sensory_sessions", m_sensoryId);

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_commit_cell_json(?, ?, ?, ?::text[], ?, ?, ?)");
    q.addBindValue("sensory_sessions");
    q.addBindValue(m_sensoryId);
    q.addBindValue("samples[0].name");
    q.addBindValue(QString("{samples,0,name}"));
    q.addBindValue("WontApply");
    q.addBindValue(QString("00000000-0000-0000-0000-000000000001"));
    q.addBindValue(v0 - 1);
    QVERIFY(q.exec());
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toBool(), false);
    QCOMPARE(rowVersion("sensory_sessions", m_sensoryId), v0);
}

void TstStoredFns::focusCell_preservesOtherFocusRows()
{
    const QString uuidA = "00000000-0000-0000-0000-aaaaaaaaaaaa";

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_focus_cell(?, ?, ?, ?, ?, ?)");
    q.addBindValue(uuidA);
    q.addBindValue("data_rows");
    q.addBindValue(m_dataRowId);
    q.addBindValue("draw_pressure");
    q.addBindValue("UserA");
    q.addBindValue("#aabbcc");
    QVERIFY(q.exec());

    // Second focus on the SAME row but DIFFERENT column. Under the M4
    // scoped DELETE the first focus is removed (because it's NOT the
    // new cell) and only the new focus survives.
    q.prepare("SELECT dve_focus_cell(?, ?, ?, ?, ?, ?)");
    q.addBindValue(uuidA);
    q.addBindValue("data_rows");
    q.addBindValue(m_dataRowId);
    q.addBindValue("smell");
    q.addBindValue("UserA");
    q.addBindValue("#aabbcc");
    QVERIFY(q.exec());

    QVERIFY(q.exec(QString(
        "SELECT count(*) FROM cell_focus WHERE user_uuid='%1'::uuid")
        .arg(uuidA)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toInt(), 1);

    QVERIFY(q.exec(QString(
        "SELECT column_name FROM cell_focus WHERE user_uuid='%1'::uuid")
        .arg(uuidA)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toString(), QStringLiteral("smell"));
}

void TstStoredFns::focusCell_refocusUpdatesStartedAt()
{
    const QString uuid = "00000000-0000-0000-0000-bbbbbbbbbbbb";
    QSqlQuery q(m_conn->queryDb());

    q.prepare("SELECT dve_focus_cell(?, ?, ?, ?, ?, ?)");
    q.addBindValue(uuid);
    q.addBindValue("data_rows");
    q.addBindValue(m_dataRowId);
    q.addBindValue("notes");
    q.addBindValue("U");
    q.addBindValue("#abcdef");
    QVERIFY(q.exec());

    QVERIFY(q.exec(QString(
        "SELECT started_at FROM cell_focus WHERE user_uuid='%1'::uuid")
        .arg(uuid)));
    QVERIFY(q.next());
    const QDateTime t1 = q.value(0).toDateTime();

    QTest::qWait(1100);

    q.prepare("SELECT dve_focus_cell(?, ?, ?, ?, ?, ?)");
    q.addBindValue(uuid);
    q.addBindValue("data_rows");
    q.addBindValue(m_dataRowId);
    q.addBindValue("notes");
    q.addBindValue("U");
    q.addBindValue("#abcdef");
    QVERIFY(q.exec());

    QVERIFY(q.exec(QString(
        "SELECT started_at FROM cell_focus WHERE user_uuid='%1'::uuid")
        .arg(uuid)));
    QVERIFY(q.next());
    const QDateTime t2 = q.value(0).toDateTime();
    QVERIFY(t2 > t1);
}

void TstStoredFns::cronJob_cellFocusCleanupExists()
{
    // pg_cron is not loaded in the ephemeral test container (init.sql's
    // CREATE EXTENSION tail is skipped by start-test-postgres.ps1). The
    // production NAS DOES have pg_cron, so skip cleanly here.
    QSqlQuery probe(m_conn->queryDb());
    if (!probe.exec("SELECT 1 FROM pg_extension WHERE extname='pg_cron'")
        || !probe.next()) {
        QSKIP("pg_cron extension not loaded in this test container");
    }

    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(
        "SELECT schedule FROM cron.job WHERE jobname='dve_cell_focus_cleanup'"));
    QVERIFY2(q.next(),
        "dve_cell_focus_cleanup not scheduled - migration not applied");
    QCOMPARE(q.value(0).toString(), QStringLiteral("* * * * *"));
}

void TstStoredFns::bumpVersion_singleIncrementPerUpdate()
{
    const int v0 = rowVersion("data_rows", m_dataRowId);

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_commit_cell(?, ?, ?, ?, ?, ?)");
    q.addBindValue("data_rows");
    q.addBindValue(m_dataRowId);
    q.addBindValue("draw_pressure");
    q.addBindValue("4.0");
    q.addBindValue(QString("00000000-0000-0000-0000-000000000001"));
    q.addBindValue(v0);
    QVERIFY(q.exec());
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toBool(), true);
    // Exactly +1 - no double bump from a redundant SET version = version + 1.
    QCOMPARE(rowVersion("data_rows", m_dataRowId), v0 + 1);
}

void TstStoredFns::commitSessionLayout_returnsNewVersionOnSuccess()
{
    const int v0 = rowVersion("sensory_sessions", m_sensoryId);

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_commit_session_layout(?, ?, ?::jsonb, ?, ?)");
    q.addBindValue("sensory_sessions");
    q.addBindValue(m_sensoryId);
    q.addBindValue(QString("{\"v\":2}"));
    q.addBindValue(QString("00000000-0000-0000-0000-000000000001"));
    q.addBindValue(v0);
    QVERIFY2(q.exec(), qPrintable(q.lastError().text()));
    QVERIFY(q.next());
    QVERIFY(!q.value(0).isNull());
    QCOMPARE(q.value(0).toInt(), v0 + 1);
}

void TstStoredFns::commitSessionLayout_returnsNullOnStaleVersion()
{
    const int v0 = rowVersion("sensory_sessions", m_sensoryId);

    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT dve_commit_session_layout(?, ?, ?::jsonb, ?, ?)");
    q.addBindValue("sensory_sessions");
    q.addBindValue(m_sensoryId);
    q.addBindValue(QString("{\"v\":99}"));
    q.addBindValue(QString("00000000-0000-0000-0000-000000000001"));
    q.addBindValue(v0 - 1);
    QVERIFY(q.exec());
    QVERIFY(q.next());
    QVERIFY(q.value(0).isNull());
    QCOMPARE(rowVersion("sensory_sessions", m_sensoryId), v0);
}

QTEST_MAIN(TstStoredFns)
#include "tst_storedfns.moc"

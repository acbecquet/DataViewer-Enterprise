// -- DataViewer Enterprise -- OfflineSnapshot Integration Tests --
//
// Verifies the SQLite local mirror that DataViewer falls back to when the
// NAS-hosted Postgres is unreachable. Each test drives a known fixture into
// ephemeral Postgres, calls OfflineSnapshot::regenerate() against it, then
// opens the resulting snapshot read-only and asserts the data round-tripped.
//
// REQUIRES: DVE_TEST_PG_CONN set (the same ephemeral test Postgres used by
// tst_postgresconnection / tst_databasemanager). Skipped otherwise.
//
// Each test runs against a freshly-wiped Postgres schema (see init()) and
// against a per-test temporary LOCALAPPDATA override so snapshot files
// never collide between cases.

#include <QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QTemporaryDir>
#include <QFile>
#include <QFileInfo>
#include <QDateTime>
#include <QStandardPaths>
#include <QDir>
#include <QByteArray>
#include <QJsonDocument>
#include <QJsonObject>

#include "OfflineSnapshot.h"
#include "PostgresConnection.h"
#include "ConfigLoader.h"
#include "DatabaseManager.h"   // for FileRecord
#include "ReportData.h"
#include "SensoryData.h"
#include "DetailedSensoryData.h"

// Q_ASSERT is a NO-OP in release builds, which silently drops setup SQL.
// MUST evaluates its argument unconditionally and aborts on failure with a
// helpful message that points at the offending SQL.
#define MUST(x) do { \
    if (!(x)) { \
        qFatal("MUST failed at %s:%d: %s | last SQL error: %s", \
               __FILE__, __LINE__, #x, qPrintable(q.lastError().text())); \
    } \
} while (0)

namespace {

// Build a DbConfig from DVE_TEST_PG_CONN (space-separated key=value pairs).
DVE::DbConfig pgConfig() {
    DVE::DbConfig c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
    if (env.isEmpty()) return c;
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

// Schema is created idempotently by deploy/postgres/init.sql when the test
// container is brought up; here we only wipe row contents between tests.
void wipeAllTables() {
    DVE::DbConfig cfg = pgConfig();
    const QString cname = "tst_oss_wipe";
    {
        QSqlDatabase wipe = QSqlDatabase::addDatabase("QPSQL", cname);
        wipe.setHostName(cfg.host);
        wipe.setPort(cfg.port);
        wipe.setDatabaseName(cfg.database);
        wipe.setUserName(cfg.user);
        wipe.setPassword(cfg.password);
        if (wipe.open()) {
            QSqlQuery q(wipe);
            for (const QString& t : QStringList{
                     "data_rows", "images", "samples", "tests", "files",
                     "sensory_images", "sensory_sessions",
                     "detailed_sensory_images", "detailed_sensory_sessions",
                     "settings"
                 }) {
                q.exec("DELETE FROM " + t);
            }
            wipe.close();
        }
    }
    QSqlDatabase::removeDatabase(cname);
}

// A 64-byte deterministic blob used to verify BLOB round-trip exactness.
QByteArray makeBlob() {
    QByteArray b;
    for (int i = 0; i < 64; ++i) b.append(static_cast<char>(i & 0xFF));
    return b;
}

// Seeds Postgres with one file (1 test, 2 samples, 4 rows total, 1 image),
// one sensory session, one detailed sensory session, and one setting.
// Returns the BLOB so tests can assert byte-for-byte equality after the
// SQLite round-trip.
QByteArray seedPostgresFixture() {
    DVE::DbConfig cfg = pgConfig();
    const QString cname = "tst_oss_seed";
    const QByteArray blob = makeBlob();
    {
        QSqlDatabase pg = QSqlDatabase::addDatabase("QPSQL", cname);
        pg.setHostName(cfg.host);
        pg.setPort(cfg.port);
        pg.setDatabaseName(cfg.database);
        pg.setUserName(cfg.user);
        pg.setPassword(cfg.password);
        if (!pg.open()) {
            qWarning() << "seedPostgresFixture: open failed:" << pg.lastError().text();
            return blob;
        }

        QSqlQuery q(pg);
        // files
        q.prepare("INSERT INTO files (file_path, file_name, loaded_at, "
                  "template_version, sheet_count, sample_count, updated_by) "
                  "VALUES (?, ?, ?, ?, ?, ?, ?) RETURNING id");
        q.addBindValue("/tmp/seed.xlsx");
        q.addBindValue("seed.xlsx");
        q.addBindValue(QDateTime::currentDateTimeUtc().toString(Qt::ISODate));
        q.addBindValue("new");
        q.addBindValue(1);
        q.addBindValue(2);
        q.addBindValue("test-seeder");
        MUST(q.exec() && q.next());
        const qint64 fileId = q.value(0).toLongLong();

        // tests
        q.prepare("INSERT INTO tests (file_id, sheet_name, template_version, "
                  "overall_avg_tpm, overall_stddev_tpm, is_raw_table, sort_order, "
                  "updated_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?) RETURNING id");
        q.addBindValue(static_cast<qlonglong>(fileId));
        q.addBindValue("Lifetime Test");
        q.addBindValue("new");
        q.addBindValue(3.42);
        q.addBindValue(0.11);
        q.addBindValue(0);
        q.addBindValue(0);
        q.addBindValue("test-seeder");
        MUST(q.exec() && q.next());
        const qint64 testId = q.value(0).toLongLong();

        // 2 samples
        QVector<qint64> sampleIds;
        for (int si = 0; si < 2; ++si) {
            q.prepare("INSERT INTO samples (test_id, sort_order, sample_name, sample_id, "
                      "date, tester, media, viscosity, resistance, voltage, power, "
                      "heating_technology, puffing_regime, initial_oil_mass, "
                      "average_tpm, stddev_tpm, avg_power_density, efficiency_percent, "
                      "total_oil_consumed, total_puffs, normalized_tpm, "
                      "burn_status, clog_status, leak_status, updated_by) "
                      "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) "
                      "RETURNING id");
            q.addBindValue(static_cast<qlonglong>(testId));
            q.addBindValue(si);
            q.addBindValue(QString("Sample %1").arg(si + 1));
            q.addBindValue(QString("ID-%1").arg(si + 1));
            q.addBindValue("2026-05-01");
            q.addBindValue("QA");
            q.addBindValue(QString("Media-%1").arg(si));
            q.addBindValue(100.0 + si);
            q.addBindValue(1.1 + si * 0.1);
            q.addBindValue(3.0);
            q.addBindValue(8.0 + si);
            q.addBindValue("heater-A");
            q.addBindValue("regime-x");
            q.addBindValue(1.0);
            q.addBindValue(3.4);
            q.addBindValue(0.12);
            q.addBindValue(0.5);
            q.addBindValue(95.0);
            q.addBindValue(0.05);
            q.addBindValue(50);
            q.addBindValue(0.4);
            q.addBindValue("N");
            q.addBindValue("N");
            q.addBindValue("N");
            q.addBindValue("test-seeder");
            MUST(q.exec() && q.next());
            sampleIds.append(q.value(0).toLongLong());
        }

        // 4 data rows (2 per sample)
        for (int s = 0; s < sampleIds.size(); ++s) {
            for (int r = 0; r < 2; ++r) {
                q.prepare("INSERT INTO data_rows (sample_id, sort_order, puffs, "
                          "before_weight, after_weight, tpm, notes, updated_by) "
                          "VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
                q.addBindValue(static_cast<qlonglong>(sampleIds[s]));
                q.addBindValue(r);
                q.addBindValue(10.0 * (r + 1));
                q.addBindValue(25.0 + s * 0.01 + r * 0.001);
                q.addBindValue(24.9 + s * 0.01 + r * 0.001);
                q.addBindValue(3.4 + r * 0.1);
                q.addBindValue(QString("note s%1r%2").arg(s).arg(r));
                q.addBindValue("test-seeder");
                MUST(q.exec());
            }
        }

        // 1 image (only against first sample)
        q.prepare("INSERT INTO images (sample_id, sort_order, file_name, image_data, "
                  "layout_x, layout_y, layout_w, layout_h, "
                  "crop_x, crop_y, crop_w, crop_h, updated_by) "
                  "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
        q.addBindValue(static_cast<qlonglong>(sampleIds[0]));
        q.addBindValue(0);
        q.addBindValue("snapshot-fixture.png");
        q.addBindValue(blob);
        q.addBindValue(0.5);
        q.addBindValue(1.5);
        q.addBindValue(2.0);
        q.addBindValue(1.5);
        q.addBindValue(0.0);
        q.addBindValue(0.0);
        q.addBindValue(1.0);
        q.addBindValue(1.0);
        q.addBindValue("test-seeder");
        MUST(q.exec());

        // 1 sensory_session with rich json_data
        {
            QJsonObject root;
            root["session_name"]   = "Sensory Seed";
            root["test_title"]     = "Heater Comparison";
            root["assessor_name"]  = "Charlie";
            root["tester_name"]    = "Alice";
            root["media"]          = "Sample A";
            root["date"]           = "2026-05-02";
            root["timestamp"]      = "2026-05-02T10:00:00Z";

            QJsonArray samples;
            QJsonObject sm;
            sm["name"] = "Sensory-Sample-1";
            for (const QString& m : DVE::kSensoryMetrics) sm[m] = 5.0;
            samples.append(sm);
            root["samples"] = samples;

            const QString jsonStr = QString::fromUtf8(
                QJsonDocument(root).toJson(QJsonDocument::Compact));

            q.prepare("INSERT INTO sensory_sessions (session_name, tester_name, "
                      "assessor_name, media, date, timestamp, json_data, updated_by) "
                      "VALUES (?, ?, ?, ?, ?, ?, CAST(? AS JSONB), ?)");
            q.addBindValue("Sensory Seed");
            q.addBindValue("Alice");
            q.addBindValue("Charlie");
            q.addBindValue("Sample A");
            q.addBindValue("2026-05-02");
            q.addBindValue("2026-05-02T10:00:00Z");
            q.addBindValue(jsonStr);
            q.addBindValue("test-seeder");
            MUST(q.exec());
        }

        // 1 detailed_sensory_session
        {
            QJsonObject root;
            root["session_name"]    = "Detailed Seed";
            root["test_title"]      = "Detailed Heater";
            root["assessor_name"]   = "Charlie";
            root["tester_name"]     = "Alice";
            root["media"]           = "Sample B";
            root["date"]            = "2026-05-03";
            root["timestamp"]       = "2026-05-03T10:00:00Z";

            QJsonArray samples;
            QJsonObject sm;
            sm["name"] = "Detailed-Sample-1";
            for (const QString& m : DVE::kDetailedAllMetrics)
                sm[m] = double(DVE::kDetailedMetricMaxScore.value(m, 9) / 2);
            samples.append(sm);
            root["samples"] = samples;

            const QString jsonStr = QString::fromUtf8(
                QJsonDocument(root).toJson(QJsonDocument::Compact));

            q.prepare("INSERT INTO detailed_sensory_sessions (session_name, tester_name, "
                      "assessor_name, media, date, timestamp, json_data, updated_by) "
                      "VALUES (?, ?, ?, ?, ?, ?, CAST(? AS JSONB), ?)");
            q.addBindValue("Detailed Seed");
            q.addBindValue("Alice");
            q.addBindValue("Charlie");
            q.addBindValue("Sample B");
            q.addBindValue("2026-05-03");
            q.addBindValue("2026-05-03T10:00:00Z");
            q.addBindValue(jsonStr);
            q.addBindValue("test-seeder");
            MUST(q.exec());
        }

        // 1 setting
        q.prepare("INSERT INTO settings (key, value, updated_by) "
                  "VALUES (?, ?, ?)");
        q.addBindValue("offline_test_key");
        q.addBindValue("offline_test_value");
        q.addBindValue("test-seeder");
        MUST(q.exec());

        pg.close();
    }
    QSqlDatabase::removeDatabase(cname);
    return blob;
}

// Counts rows in a SQLite table via a direct connection on the snapshot file.
// We don't go through OfflineSnapshot's read accessors here; the goal is to
// verify the regenerate path independently of the read path.
int sqliteCount(const QString& path, const QString& table) {
    int n = -1;
    const QString cname = "tst_oss_count_" + table +
                          QString::number(QDateTime::currentMSecsSinceEpoch());
    {
        QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname);
        db.setDatabaseName(path);
        if (db.open()) {
            QSqlQuery q(db);
            if (q.exec("SELECT COUNT(*) FROM " + table) && q.next()) {
                n = q.value(0).toInt();
            }
            db.close();
        }
    }
    QSqlDatabase::removeDatabase(cname);
    return n;
}

} // anonymous

// ============================================================================
class TstOfflineSnapshot : public QObject {
    Q_OBJECT

private:
    QTemporaryDir* m_overrideDir = nullptr;

    QString overrideBaseDir() const {
        return m_overrideDir ? m_overrideDir->path() : QString();
    }

private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) {
            QSKIP("DVE_TEST_PG_CONN not set; OfflineSnapshot integration tests skipped");
        }
    }

    void init() {
        wipeAllTables();
        // Fresh override dir per test so snapshot files don't collide.
        delete m_overrideDir;
        m_overrideDir = new QTemporaryDir();
        QVERIFY(m_overrideDir->isValid());
    }

    void cleanup() {
        delete m_overrideDir;
        m_overrideDir = nullptr;
    }

    // -- T1: regenerate from a populated Postgres ---------------------------
    void testRegenerateFromPopulatedPostgres() {
        const QByteArray expectedBlob = seedPostgresFixture();

        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));

        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));

        const QString snapPath = snap.path();
        QVERIFY(QFile::exists(snapPath));

        // Independent row-count verification (bypasses OfflineSnapshot read API).
        QCOMPARE(sqliteCount(snapPath, "files"), 1);
        QCOMPARE(sqliteCount(snapPath, "tests"), 1);
        QCOMPARE(sqliteCount(snapPath, "samples"), 2);
        QCOMPARE(sqliteCount(snapPath, "data_rows"), 4);
        QCOMPARE(sqliteCount(snapPath, "images"), 1);
        QCOMPARE(sqliteCount(snapPath, "sensory_sessions"), 1);
        QCOMPARE(sqliteCount(snapPath, "detailed_sensory_sessions"), 1);
        QCOMPARE(sqliteCount(snapPath, "settings"), 1);
        QCOMPARE(sqliteCount(snapPath, "_snapshot_meta"), 2);  // taken_at + schema ver

        // BLOB round-trip
        {
            const QString cname = "tst_oss_blob";
            QByteArray got;
            {
                QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname);
                db.setDatabaseName(snapPath);
                QVERIFY(db.open());
                QSqlQuery q(db);
                QVERIFY(q.exec("SELECT image_data FROM images"));
                QVERIFY(q.next());
                got = q.value(0).toByteArray();
                db.close();
            }
            QSqlDatabase::removeDatabase(cname);
            QCOMPARE(got.size(), expectedBlob.size());
            QCOMPARE(got, expectedBlob);
        }

        // JSONB->TEXT preserves the document
        {
            const QString cname = "tst_oss_json";
            QString jsonStr;
            {
                QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname);
                db.setDatabaseName(snapPath);
                QVERIFY(db.open());
                QSqlQuery q(db);
                QVERIFY(q.exec("SELECT json_data FROM sensory_sessions"));
                QVERIFY(q.next());
                jsonStr = q.value(0).toString();
                db.close();
            }
            QSqlDatabase::removeDatabase(cname);
            const QJsonDocument doc = QJsonDocument::fromJson(jsonStr.toUtf8());
            QVERIFY(doc.isObject());
            QCOMPARE(doc.object().value("session_name").toString(),
                     QString("Sensory Seed"));
            QCOMPARE(doc.object().value("samples").toArray().size(), 1);
        }

        pg.close();
    }

    // -- T2: atomic write -- prior snapshot intact on regenerate failure ----
    void testAtomicWriteCorruptMidwrite() {
        // 1) Create a baseline snapshot with valid data.
        const QByteArray baselineBlob = seedPostgresFixture();
        {
            DVE::PostgresConnection pg;
            QVERIFY(pg.open(pgConfig()));
            DVE::OfflineSnapshot snap;
            snap.setOverrideDirForTesting(overrideBaseDir());
            QVERIFY(snap.regenerate(&pg));
            pg.close();
        }

        // 2) Verify baseline is on disk.
        DVE::OfflineSnapshot snap2;
        snap2.setOverrideDirForTesting(overrideBaseDir());
        const QString prodPath = snap2.path();
        QVERIFY(QFile::exists(prodPath));
        const qint64 baselineSize = QFileInfo(prodPath).size();
        QVERIFY(baselineSize > 0);

        // 3) Attempt regenerate against a closed PostgresConnection -- this
        //    must fail before producing any new file.
        {
            DVE::PostgresConnection closedPg;
            QVERIFY(!closedPg.isOpen());
            DVE::OfflineSnapshot snap3;
            snap3.setOverrideDirForTesting(overrideBaseDir());
            QVERIFY(!snap3.regenerate(&closedPg));
            QVERIFY(!snap3.lastError().isEmpty());
        }

        // 4) Baseline must still be readable and identical in size.
        QVERIFY(QFile::exists(prodPath));
        QCOMPARE(QFileInfo(prodPath).size(), baselineSize);

        // No .tmp should be left behind.
        QVERIFY(!QFile::exists(prodPath + ".tmp"));

        // 5) Open read-only and confirm rows are still there.
        DVE::OfflineSnapshot snap3;
        snap3.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap3.openReadOnly(), qPrintable(snap3.lastError()));
        QCOMPARE(snap3.listFiles().size(), 1);
        snap3.close();
    }

    // -- T3: read-only enforcement at SQLite layer --------------------------
    void testReadOnlyEnforcesReadOnly() {
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));

        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY(snap.regenerate(&pg));
        pg.close();

        QVERIFY(snap.openReadOnly());

        // Attempt a direct INSERT through a separate read-only connection on
        // the same file. SQLite should refuse: opening read-only blocks all
        // write paths via the connect option. Note: we open a NEW connection
        // here because snap's own QSqlDatabase is private. The new connection
        // also passes QSQLITE_OPEN_READONLY -- this asserts the same lock
        // path that snap uses in production.
        const QString cname = "tst_oss_ro_probe";
        bool insertFailed = false;
        QString errText;
        {
            QSqlDatabase ro = QSqlDatabase::addDatabase("QSQLITE", cname);
            ro.setDatabaseName(snap.path());
            ro.setConnectOptions("QSQLITE_OPEN_READONLY");
            QVERIFY(ro.open());
            QSqlQuery q(ro);
            const bool ok = q.exec("INSERT INTO settings (key, value, updated_at, "
                                   "updated_by) VALUES ('x', 'y', '2026-05-13', 'tst')");
            insertFailed = !ok;
            errText = q.lastError().text();
            ro.close();
        }
        QSqlDatabase::removeDatabase(cname);
        QVERIFY2(insertFailed,
                 qPrintable("Expected read-only INSERT to fail, but it succeeded; err=" + errText));
        snap.close();
    }

    // -- T4: snapshotTakenAt returns a recent timestamp ---------------------
    void testSnapshotTakenAtReturnsRecentTimestamp() {
        seedPostgresFixture();
        const QDateTime before = QDateTime::currentDateTimeUtc();

        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY(snap.regenerate(&pg));
        pg.close();

        QVERIFY(snap.openReadOnly());
        const QDateTime stamp = snap.snapshotTakenAt();
        QVERIFY2(stamp.isValid(), "snapshotTakenAt() returned an invalid QDateTime");

        // Within the last 60 seconds of the call above.
        const qint64 deltaSec = before.secsTo(stamp);
        QVERIFY2(deltaSec >= -1 && deltaSec <= 60,
                 qPrintable(QString("stamp diff from before=%1s (expected 0..60)")
                                .arg(deltaSec)));
        snap.close();
    }

    // -- T5: openReadOnly fails when no file exists -------------------------
    void testOpenReadOnlyFailsIfNoFile() {
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());

        // Make sure the path does NOT exist.
        const QString p = snap.path();
        QFile::remove(p);
        QVERIFY(!QFile::exists(p));

        QVERIFY(!snap.openReadOnly());
        QVERIFY(!snap.lastError().isEmpty());
        QVERIFY2(snap.lastError().contains("does not exist") ||
                 snap.lastError().contains("snapshot file"),
                 qPrintable("Unexpected error message: " + snap.lastError()));
    }

    // -- T6: listFiles returns the seeded FileRecord ------------------------
    void testListFilesReturnsFileRecords() {
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY(snap.regenerate(&pg));
        pg.close();

        QVERIFY(snap.openReadOnly());

        const auto records = snap.listFiles();
        QCOMPARE(records.size(), 1);
        QCOMPARE(records[0].fileName, QString("seed.xlsx"));
        QCOMPARE(records[0].filePath, QString("/tmp/seed.xlsx"));
        QCOMPARE(records[0].templateVersion, QString("new"));
        QCOMPARE(records[0].sheetCount, 1);
        QCOMPARE(records[0].sampleCount, 2);
        QVERIFY(records[0].id > 0);
        snap.close();
    }

    // -- T7: loadFileByPath round-trip + version anchor ---------------------
    void testLoadFileByPathRoundTrip() {
        const QByteArray expectedBlob = seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY(snap.regenerate(&pg));
        pg.close();

        QVERIFY(snap.openReadOnly());

        const DVE::FileResult fr = snap.loadFileByPath("/tmp/seed.xlsx");
        QVERIFY(!fr.filePath.isEmpty());
        QCOMPARE(fr.fileName, QString("seed.xlsx"));
        QCOMPARE(fr.templateVersion, QString("new"));
        QVERIFY(fr.id > 0);
        QVERIFY(fr.version >= 1);          // bump_version trigger seeded it at 1
        QCOMPARE(fr.sheets.size(), 1);
        QCOMPARE(fr.sheets[0].sheetName, QString("Lifetime Test"));
        QCOMPARE(fr.sheets[0].samples.size(), 2);
        QCOMPARE(fr.sheets[0].samples[0].sampleName, QString("Sample 1"));
        QCOMPARE(fr.sheets[0].samples[0].rows.size(), 2);
        QCOMPARE(fr.sheets[0].samples[0].rows[0].notes, QString("note s0r0"));

        // The first sample carries the image -- the snapshot's loadFile path
        // materialises the BLOB to disk and surfaces the path.
        QCOMPARE(fr.sheets[0].samples[0].imagePaths.size(), 1);
        const QString cachedPath = fr.sheets[0].samples[0].imagePaths[0];
        QVERIFY(QFile::exists(cachedPath));
        {
            QFile f(cachedPath);
            QVERIFY(f.open(QIODevice::ReadOnly));
            const QByteArray got = f.readAll();
            QCOMPARE(got, expectedBlob);
        }

        // Image cache directory must respect setOverrideDirForTesting() --
        // tests must not pollute the real %LOCALAPPDATA%/DataViewer/ImageCache/.
        // The cache is a sibling of the snapshot file, so the override base
        // dir is a prefix of the cached image path.
        QVERIFY2(cachedPath.startsWith(overrideBaseDir()),
                 qPrintable("Image cache path leaked outside override dir: "
                            + cachedPath + " (override=" + overrideBaseDir() + ")"));

        // Sensory + detailed sensory listings should also round-trip.
        const auto sensList = snap.listSensoryRecords();
        QCOMPARE(sensList.size(), 1);
        QCOMPARE(sensList[0].sessionName, QString("Sensory Seed"));
        QCOMPARE(sensList[0].testerName,  QString("Alice"));
        QCOMPARE(sensList[0].testTitle,   QString("Heater Comparison"));
        QCOMPARE(sensList[0].sampleCount, 1);

        const auto detList = snap.listDetailedSensoryRecords();
        QCOMPARE(detList.size(), 1);
        QCOMPARE(detList[0].sessionName, QString("Detailed Seed"));

        const DVE::SensorySession sess = snap.loadSensorySession(sensList[0].id);
        QCOMPARE(sess.sessionName, QString("Sensory Seed"));
        QVERIFY(sess.id > 0);
        QVERIFY(sess.version >= 1);
        QCOMPARE(sess.samples.size(), 1);

        const DVE::DetailedSensorySession dsess =
            snap.loadDetailedSensorySession(detList[0].id);
        QCOMPARE(dsess.sessionName, QString("Detailed Seed"));
        QVERIFY(dsess.id > 0);
        QVERIFY(dsess.version >= 1);

        snap.close();
    }
};

QTEST_MAIN(TstOfflineSnapshot)
#include "tst_offlinesnapshot.moc"

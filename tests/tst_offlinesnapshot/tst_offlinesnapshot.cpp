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
#include "SeedRows.h"
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QTemporaryDir>
#include <QFile>
#include <QFileInfo>
#include <QDateTime>
#include <QTimeZone>
#include <QStandardPaths>
#include <QDir>
#include <QByteArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QMetaType>
#include <QRegularExpression>
#include <atomic>

#include "OfflineSnapshot.h"
#include "PostgresConnection.h"
#include "ConfigLoader.h"
#include "DatabaseManager.h"   // for FileRecord
#include "ReportData.h"
#include "SensoryData.h"
#include "DetailedSensoryData.h"
#include "RawGridJson.h"
#include "MipFallback.h"       // R5: MIP-fallback unit + store-level fallback
#include "SnapshotRegenWorker.h"  // SP4.5 Stage 2a
#include <QThread>
#include <QEventLoop>
#include <QTimer>

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

// SP4.5 Stage 2b: deterministic content signature of a snapshot table (row count
// + per-row id/updated_at, plus blob length for image tables) so two snapshot
// files can be compared without depending on blob byte-equality in SQL. Scan is
// rowid (== id) order, identical for an incremental vs a full rebuild.
QString tableSig(const QString& dbPath, const QString& table) {
    const QString conn = "tst_oss_sig";
    QString out;
    {
        QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", conn);
        db.setDatabaseName(dbPath);
        db.setConnectOptions("QSQLITE_OPEN_READONLY");
        if (db.open()) {
            const bool hasBlob = (table == "images" || table == "sensory_images"
                                  || table == "detailed_sensory_images");
            const QString expr = hasBlob
                ? "count(*)||'|'||coalesce(group_concat(id||':'||updated_at||':'||length(coalesce(image_data,x''))),'')"
                : "count(*)||'|'||coalesce(group_concat(id||':'||updated_at),'')";
            QSqlQuery q(db);
            if (q.exec("SELECT " + expr + " FROM " + table) && q.next())
                out = q.value(0).toString();
            db.close();
        }
    }
    QSqlDatabase::removeDatabase(conn);
    return out;
}

// v3 Phase 3c / HAZARD H24: the (kind|key) vocabulary pairs the open-metric
// slots below invent. `metric_defs` is a SHARED vocabulary table that is
// deliberately never wiped wholesale -- tst_v3longformat's migration RAISEs
// without the 35 seeded keys, and
// tst_databasemanager::v3LongFormat_ensureSchemaMatchesRegistry asserts an
// EXACT row count against the compiled registry. A key leaked by this suite
// therefore turns a DIFFERENT suite red for a reason that has nothing to do
// with it. Every key seedOpenMetricsFixture() invents must appear here.
QStringList openMetricKeys() {
    return QStringList{ "metric|snap_coil_temp", "metric|snap_rig_note",
                        "metric|snap_flag",      "metric|snap_pv",
                        "header|snap_rig_id",    "header|snap_coil_count" };
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
                     // v3 Phase 3c: the long tables go FIRST. samples' ON DELETE
                     // CASCADE would take them anyway; the explicit delete keeps
                     // the intent visible and survives a tightened cascade.
                     "measurements", "sample_headers",
                     "data_rows", "images", "samples", "tests", "files",
                     "sensory_images", "sensory_sessions",
                     "detailed_sensory_images", "detailed_sensory_sessions",
                     "settings"
                 }) {
                q.exec("DELETE FROM " + t);
            }
            // H24: exactly the invented vocabulary keys, after the long tables
            // so no FK is left dangling. See openMetricKeys().
            for (const QString& k : openMetricKeys()) {
                q.prepare("DELETE FROM metric_defs WHERE kind = ? AND key = ?");
                q.addBindValue(k.section('|', 0, 0));
                q.addBindValue(k.section('|', 1, 1));
                q.exec();
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

        // 2 samples. v3 Phase 3d (plan Task 4): via the dual-mode helper so
        // the fixture stays legal on both sides of the cutover.
        QVector<qint64> sampleIds;
        for (int si = 0; si < 2; ++si) {
            const qint64 sid = DVE::TestSeed::seedSample(pg, testId, {
                { QStringLiteral("sample_name"),        QString("Sample %1").arg(si + 1) },
                { QStringLiteral("sample_id"),          QString("ID-%1").arg(si + 1) },
                { QStringLiteral("date"),               QStringLiteral("2026-05-01") },
                { QStringLiteral("tester"),             QStringLiteral("QA") },
                { QStringLiteral("media"),              QString("Media-%1").arg(si) },
                { QStringLiteral("viscosity"),          100.0 + si },
                { QStringLiteral("resistance"),         1.1 + si * 0.1 },
                { QStringLiteral("voltage"),            3.0 },
                { QStringLiteral("power"),              8.0 + si },
                { QStringLiteral("heating_technology"), QStringLiteral("heater-A") },
                { QStringLiteral("puffing_regime"),     QStringLiteral("regime-x") },
                { QStringLiteral("initial_oil_mass"),   1.0 },
                { QStringLiteral("average_tpm"),        3.4 },
                { QStringLiteral("stddev_tpm"),         0.12 },
                { QStringLiteral("avg_power_density"),  0.5 },
                { QStringLiteral("efficiency_percent"), 95.0 },
                { QStringLiteral("total_oil_consumed"), 0.05 },
                { QStringLiteral("total_puffs"),        50 },
                { QStringLiteral("normalized_tpm"),     0.4 },
                { QStringLiteral("burn_status"),        QStringLiteral("N") },
                { QStringLiteral("clog_status"),        QStringLiteral("N") },
                { QStringLiteral("leak_status"),        QStringLiteral("N") },
            }, si, QStringLiteral("test-seeder"));
            MUST(sid > 0);
            sampleIds.append(sid);
        }

        // 4 data rows (2 per sample)
        for (int s = 0; s < sampleIds.size(); ++s) {
            for (int r = 0; r < 2; ++r) {
                MUST(DVE::TestSeed::seedDataRow(pg, sampleIds[s], r, {
                    { QStringLiteral("puffs"),         10.0 * (r + 1) },
                    { QStringLiteral("before_weight"), 25.0 + s * 0.01 + r * 0.001 },
                    { QStringLiteral("after_weight"),  24.9 + s * 0.01 + r * 0.001 },
                    { QStringLiteral("tpm"),           3.4 + r * 0.1 },
                    { QStringLiteral("notes"),         QString("note s%1r%2").arg(s).arg(r) },
                }, QStringLiteral("test-seeder")) >= 0);
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

// v3 Phase 3c: seeds the long-format vocabulary and one sample's worth of open
// metrics (DataRow::extra / SampleResult::extra) for the file that
// seedPostgresFixture() just created. Row extras land on the FIRST sample's two
// data rows (sort_order 0 and 1); header extras land on that same sample. The
// SECOND sample is left bare on purpose so the sparse rule (index D2 -- a value
// with nothing to say is not stored at all) is exercised in the same slot.
//
// The values deliberately span all four storage shapes the
// value_type <-> (value_num, value_text) contract routes on -- number, text,
// bool, numberlist -- so a regen plus an offline read is proved to hand them
// back TYPED rather than stringly. Exactly one of value_num / value_text is
// populated per row, which is why the two INSERTs are separate statements
// rather than one with NULL binds.
void seedOpenMetricsFixture() {
    DVE::DbConfig cfg = pgConfig();
    const QString cname = "tst_oss_openmetrics_seed";
    {
        QSqlDatabase pg = QSqlDatabase::addDatabase("QPSQL", cname);
        pg.setHostName(cfg.host);
        pg.setPort(cfg.port);
        pg.setDatabaseName(cfg.database);
        pg.setUserName(cfg.user);
        pg.setPassword(cfg.password);
        if (!pg.open()) {
            qWarning() << "seedOpenMetricsFixture: open failed:"
                       << pg.lastError().text();
            return;
        }

        QSqlQuery q(pg);

        QVector<qint64> sampleIds;
        MUST(q.exec("SELECT id FROM samples ORDER BY sort_order"));
        while (q.next()) sampleIds.append(q.value(0).toLongLong());
        MUST(sampleIds.size() >= 2);
        const qlonglong sid = static_cast<qlonglong>(sampleIds[0]);

        // A vocabulary row; the id is what measurements / sample_headers key on.
        auto def = [&](const char* kind, const char* key,
                       const char* valueType) -> qlonglong {
            q.prepare("INSERT INTO metric_defs (kind, key, display_name, "
                      "value_type, role, updated_by) "
                      "VALUES (?, ?, ?, ?, 'measured', 'test-seeder') RETURNING id");
            q.addBindValue(QLatin1String(kind));
            q.addBindValue(QLatin1String(key));
            q.addBindValue(QLatin1String(key));      // display_name IS the key
            q.addBindValue(QLatin1String(valueType));
            MUST(q.exec() && q.next());
            return q.value(0).toLongLong();
        };
        auto measNum = [&](qlonglong metricId, int sortOrder, double v) {
            q.prepare("INSERT INTO measurements (sample_id, metric_id, sort_order, "
                      "value_num, updated_by) VALUES (?, ?, ?, ?, 'test-seeder')");
            q.addBindValue(sid);
            q.addBindValue(metricId);
            q.addBindValue(sortOrder);
            q.addBindValue(v);
            MUST(q.exec());
        };
        auto measText = [&](qlonglong metricId, int sortOrder, const QString& v) {
            q.prepare("INSERT INTO measurements (sample_id, metric_id, sort_order, "
                      "value_text, updated_by) VALUES (?, ?, ?, ?, 'test-seeder')");
            q.addBindValue(sid);
            q.addBindValue(metricId);
            q.addBindValue(sortOrder);
            q.addBindValue(v);
            MUST(q.exec());
        };

        const qlonglong mCoil = def("metric", "snap_coil_temp", "number");
        const qlonglong mNote = def("metric", "snap_rig_note",  "text");
        const qlonglong mFlag = def("metric", "snap_flag",      "bool");
        const qlonglong mPv   = def("metric", "snap_pv",        "numberlist");
        const qlonglong hRig  = def("header", "snap_rig_id",    "text");
        const qlonglong hCnt  = def("header", "snap_coil_count","number");

        // Row 0 carries one of every shape; row 1 carries a single number, so
        // the per-row (sample_id, sort_order) mapping has to be right or the
        // two rows' values swap.
        measNum (mCoil, 0, 42.5);
        measText(mNote, 0, QStringLiteral("rig A"));
        measNum (mFlag, 0, 1.0);                       // bool -> value_num 0/1
        measText(mPv,   0, QStringLiteral("[1,2.5,3]"));  // numberlist -> JSON
        measNum (mCoil, 1, 43.5);

        q.prepare("INSERT INTO sample_headers (sample_id, field_id, value_text, "
                  "updated_by) VALUES (?, ?, ?, 'test-seeder')");
        q.addBindValue(sid);
        q.addBindValue(hRig);
        q.addBindValue(QStringLiteral("RIG-7"));
        MUST(q.exec());

        q.prepare("INSERT INTO sample_headers (sample_id, field_id, value_num, "
                  "updated_by) VALUES (?, ?, ?, 'test-seeder')");
        q.addBindValue(sid);
        q.addBindValue(hCnt);
        q.addBindValue(3.0);
        MUST(q.exec());

        pg.close();
    }
    QSqlDatabase::removeDatabase(cname);
}

// Seeds Postgres with one files row, one sensory_sessions row and one
// detailed_sensory_sessions row, each carrying an EXPLICIT non-NULL
// app_version (e.g. "DataViewer/2.4.1"). The server-side stamp trigger only
// fills app_version when it is NULL on write, so binding it directly here is
// honoured. Used by snapshot_carriesAppVersionColumn() to prove R7b copies
// the column through regenerate(). Returns the app_version string seeded.
QString seedAppVersionFixture() {
    DVE::DbConfig cfg = pgConfig();
    const QString cname = "tst_oss_appver_seed";
    const QString appVersion = QStringLiteral("DataViewer/2.4.1");
    {
        QSqlDatabase pg = QSqlDatabase::addDatabase("QPSQL", cname);
        pg.setHostName(cfg.host);
        pg.setPort(cfg.port);
        pg.setDatabaseName(cfg.database);
        pg.setUserName(cfg.user);
        pg.setPassword(cfg.password);
        if (!pg.open()) {
            qWarning() << "seedAppVersionFixture: open failed:"
                       << pg.lastError().text();
            return appVersion;
        }

        QSqlQuery q(pg);
        // files row with explicit app_version
        q.prepare("INSERT INTO files (file_path, file_name, loaded_at, "
                  "template_version, sheet_count, sample_count, updated_by, "
                  "app_version) VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
        q.addBindValue("/tmp/appver.xlsx");
        q.addBindValue("appver.xlsx");
        q.addBindValue(QDateTime::currentDateTimeUtc().toString(Qt::ISODate));
        q.addBindValue("new");
        q.addBindValue(0);
        q.addBindValue(0);
        q.addBindValue("test-seeder");
        q.addBindValue(appVersion);
        MUST(q.exec());

        // sensory_sessions row with explicit app_version
        q.prepare("INSERT INTO sensory_sessions (session_name, tester_name, "
                  "assessor_name, media, date, timestamp, json_data, "
                  "updated_by, app_version) "
                  "VALUES (?, ?, ?, ?, ?, ?, CAST(? AS JSONB), ?, ?)");
        q.addBindValue("AppVer Sensory");
        q.addBindValue("Alice");
        q.addBindValue("Charlie");
        q.addBindValue("Sample A");
        q.addBindValue("2026-05-02");
        q.addBindValue("2026-05-02T10:00:00Z");
        q.addBindValue("{}");
        q.addBindValue("test-seeder");
        q.addBindValue(appVersion);
        MUST(q.exec());

        // detailed_sensory_sessions row with explicit app_version
        q.prepare("INSERT INTO detailed_sensory_sessions (session_name, "
                  "tester_name, assessor_name, media, date, timestamp, "
                  "json_data, updated_by, app_version) "
                  "VALUES (?, ?, ?, ?, ?, ?, CAST(? AS JSONB), ?, ?)");
        q.addBindValue("AppVer Detailed");
        q.addBindValue("Alice");
        q.addBindValue("Charlie");
        q.addBindValue("Sample B");
        q.addBindValue("2026-05-03");
        q.addBindValue("2026-05-03T10:00:00Z");
        q.addBindValue("{}");
        q.addBindValue("test-seeder");
        q.addBindValue(appVersion);
        MUST(q.exec());

        pg.close();
    }
    QSqlDatabase::removeDatabase(cname);
    return appVersion;
}

// Reads the live Postgres server clock (now() in UTC) via the supplied
// connection's query db. Used by snapshot_stampIsServerClock to bracket
// regeneration in server time rather than client time.
QDateTime serverNowUtc(DVE::PostgresConnection& pg) {
    QSqlQuery q(pg.queryDb());
    if (!q.exec("SELECT now() AT TIME ZONE 'UTC'") || !q.next())
        return QDateTime();
    QDateTime dt = q.value(0).toDateTime();
    dt.setTimeZone(QTimeZone::UTC);
    return dt;
}

// Reads a single TEXT column from the snapshot SQLite file via a direct
// connection (bypassing OfflineSnapshot's read accessors, mirroring
// sqliteCount). Returns a null QString on any error / no row.
QString sqliteScalar(const QString& path, const QString& sql) {
    QString val;
    const QString cname = "tst_oss_scalar_" +
                          QString::number(QDateTime::currentMSecsSinceEpoch());
    {
        QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname);
        db.setDatabaseName(path);
        if (db.open()) {
            QSqlQuery q(db);
            if (q.exec(sql) && q.next()) val = q.value(0).toString();
            db.close();
        }
    }
    QSqlDatabase::removeDatabase(cname);
    return val;
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

    // R5 minor: PG-dependent tests call this at entry so they skip cleanly when
    // DVE_TEST_PG_CONN is unset. The class is NO LONGER skipped wholesale in
    // initTestCase(): the MIP-fallback tests (mipFallback_*) have no PG
    // dependency -- they craft files directly -- and must run even without a
    // test Postgres, matching their own "No PG dependency" comments. Per-test
    // gating is the mechanism that lets the two coexist.
    static bool pgAvailable() { return !qgetenv("DVE_TEST_PG_CONN").isEmpty(); }
#define REQUIRE_PG() do { \
        if (!pgAvailable()) \
            QSKIP("DVE_TEST_PG_CONN not set; PG-backed test skipped"); \
    } while (0)

private slots:
    void initTestCase() {
        // Intentionally NOT a class-wide QSKIP: see pgAvailable()/REQUIRE_PG().
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

    // v3 Phase 3c / HAZARD H24: init() clears the invented vocabulary keys
    // BEFORE every slot, which protects this run. This clears what the LAST
    // slot left, so the shared container is handed back exactly as it was
    // found and no other suite inherits a metric_defs row from here.
    void cleanupTestCase() {
        if (pgAvailable()) wipeAllTables();
    }

    // -- T1: regenerate from a populated Postgres ---------------------------
    void testRegenerateFromPopulatedPostgres() {
        REQUIRE_PG();
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
        QCOMPARE(sqliteCount(snapPath, "_snapshot_meta"), 3);  // taken_at + schema ver + content_fingerprint (SP4.5)

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
        REQUIRE_PG();
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

    // ===================================================================
    // H14 (Phase 3a): the column-arity guard
    // ===================================================================

    // Every copied table in regenToPath() is a hand-maintained triple: a SELECT
    // column list, an INSERT with N placeholders, and a bind loop with a bound.
    // When those drift the snapshot silently mis-copies data, and the offline
    // read path is the one place a user sees data with no server to cross-check
    // against. The guard used to be a Q_ASSERT_X, i.e. it vanished from exactly
    // the release build that ships. This test runs in the release test binary
    // and pins the runtime contract: a mismatch is logged at CRITICAL severity,
    // names the table and all three counts, and is reported to the caller.
    //
    // No PG dependency -- the guard only needs a QSqlQuery whose result record
    // has a known column count, which in-memory SQLite provides.
    void columnArityMismatchIsReportedAtRuntime() {
        const QString cname = "tst_oss_arity";
        {
            QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname);
            db.setDatabaseName(":memory:");
            QVERIFY(db.open());
            QSqlQuery q(db);
            QVERIFY2(q.exec("SELECT 1 AS a, 2 AS b, 3 AS c"),
                     qPrintable(q.lastError().text()));   // 3-column result

            // Agreement (3 == 3 == 3): passes, leaves the error string alone.
            QString err = QStringLiteral("untouched");
            QVERIFY(DVE::OfflineSnapshot::checkColumnArity(q, 3, 3, "widgets", &err));
            QCOMPARE(err, QStringLiteral("untouched"));

            // The INSERT grew a placeholder the SELECT does not have.
            err.clear();
            QTest::ignoreMessage(QtCriticalMsg, QRegularExpression("widgets"));
            QVERIFY(!DVE::OfflineSnapshot::checkColumnArity(q, 4, 3, "widgets", &err));
            QVERIFY2(err.contains(QStringLiteral("widgets")), qPrintable(err));
            QVERIFY2(err.contains(QStringLiteral("SELECT columns=3")), qPrintable(err));
            QVERIFY2(err.contains(QStringLiteral("INSERT placeholders=4")), qPrintable(err));
            QVERIFY2(err.contains(QStringLiteral("bind-loop bound=3")), qPrintable(err));

            // The bind loop drifted away from the (agreeing) SELECT/INSERT pair.
            err.clear();
            QTest::ignoreMessage(QtCriticalMsg, QRegularExpression("gadgets"));
            QVERIFY(!DVE::OfflineSnapshot::checkColumnArity(q, 3, 2, "gadgets", &err));
            QVERIFY2(err.contains(QStringLiteral("bind-loop bound=2")), qPrintable(err));

            // A null out-param is tolerated (still logs, still returns false).
            QTest::ignoreMessage(QtCriticalMsg, QRegularExpression("gizmos"));
            QVERIFY(!DVE::OfflineSnapshot::checkColumnArity(q, 9, 9, "gizmos", nullptr));

            db.close();
        }
        QSqlDatabase::removeDatabase(cname);
    }

    // The rationale for aborting a regen on an arity mismatch is that a
    // stale-but-correct previous snapshot beats a silently mis-copied fresh
    // one -- which only holds if an early return really does leave the old
    // snapshot alone. regenToPath writes to a per-call unique .tmp and promotes
    // it with a single MoveFileExW at the very end, so every early return is
    // safe; this pins that empirically using the cancel seam, which bails out
    // AFTER files/tests/samples have already been copied into the tmp (the same
    // teardown the arity guard now takes).
    void midRegenAbortLeavesPriorSnapshotIntact() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));

        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        const QString prodPath = snap.path();

        QByteArray before;
        { QFile f(prodPath); QVERIFY(f.open(QIODevice::ReadOnly)); before = f.readAll(); }
        QVERIFY(!before.isEmpty());

        std::atomic<bool> cancel(true);
        QString err;
        QVERIFY(!DVE::OfflineSnapshot::regenToPath(&pg, prodPath, &cancel,
                                                   nullptr, nullptr, &err));
        QVERIFY(!err.isEmpty());

        QByteArray after;
        { QFile f(prodPath); QVERIFY(f.open(QIODevice::ReadOnly)); after = f.readAll(); }
        QCOMPARE(after, before);          // never promoted -> byte-identical

        // ...and no tmp litter survives the abort.
        const QFileInfo pi(prodPath);
        QCOMPARE(pi.absoluteDir().entryList(
                     QStringList{ pi.fileName() + QStringLiteral(".*.tmp") },
                     QDir::Files).size(), 0);
        pg.close();
    }

    // ===================================================================
    // SP4.5 Stage 2b: incremental snapshot + progress
    // ===================================================================

    // The progress callback fires and lands at done==total==kRegenPhases.
    void testRegenProgressInvokedToCompletion() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());

        int calls = 0, lastDone = -1, lastTotal = -1;
        bool monotonic = true;
        auto cb = [&](int d, int t, const QString&) {
            if (d < lastDone) monotonic = false;
            lastDone = d; lastTotal = t; ++calls;
        };
        QVERIFY2(snap.regenerate(&pg, cb), qPrintable(snap.lastError()));
        QVERIFY(calls > 0);
        QVERIFY(monotonic);
        // kRegenPhases: prepare + 13 tables + meta + finalize. v3 Phase 3c took
        // it from 13 to 16 by mirroring the three long-format tables.
        QCOMPARE(lastTotal, 16);
        QCOMPARE(lastDone, lastTotal);      // ends at 100%
        pg.close();
    }

    // Pure unit test (no PG): per-table fingerprint segment compare.
    //
    // v3 Phase 3c: the fingerprint grew from 9 tables to 12 (metric_defs,
    // measurements, sample_headers appended LAST so no existing segment index
    // moved). The literals here are built from kFingerprintTables rather than
    // hand-typed so the segment COUNT can never drift out of step with the
    // implementation again -- that drift across five coupled sites is what
    // hazard H10 was about.
    void testFingerprintSegmentChanged() {
        const int n = DVE::OfflineSnapshot::kFingerprintTableCount;
        QCOMPARE(n, 12);
        QStringList segs;
        for (int i = 0; i < n; ++i)
            segs << QStringLiteral("%1/%2").arg(i + 1).arg((i + 1) * 100);
        const QString a = segs.join(';');

        QStringList mutated = segs;
        mutated[4] = QStringLiteral("5/999");                 // images = idx 4
        const QString b = mutated.join(';');
        QVERIFY(!DVE::OfflineSnapshot::segmentChanged(a, a, "images"));
        QVERIFY( DVE::OfflineSnapshot::segmentChanged(a, b, "images"));
        QVERIFY(!DVE::OfflineSnapshot::segmentChanged(a, b, "data_rows")); // idx 3 unchanged
        QVERIFY( DVE::OfflineSnapshot::segmentChanged("garbage", b, "images")); // unparseable
        QVERIFY( DVE::OfflineSnapshot::segmentChanged(QString(), b, "images")); // empty
        QVERIFY( DVE::OfflineSnapshot::segmentChanged(a, b, "no_such_table"));  // unknown

        // A 9-segment fingerprint is what every snapshot written before schema
        // v4 stored. It must be rejected outright, not silently index-shifted.
        QVERIFY(DVE::OfflineSnapshot::fingerprintSegments(
                    QStringList(segs.mid(0, 9)).join(';')).isEmpty());
        QCOMPARE(DVE::OfflineSnapshot::fingerprintSegments(a).size(), n);

        // The three tables 3c added are addressable by name, not just present.
        QStringList longMut = segs;
        longMut[n - 1] = QStringLiteral("12/1");              // sample_headers
        QVERIFY( DVE::OfflineSnapshot::segmentChanged(a, longMut.join(';'),
                                                      "sample_headers"));
        QVERIFY(!DVE::OfflineSnapshot::segmentChanged(a, longMut.join(';'),
                                                      "measurements"));
        QVERIFY(!DVE::OfflineSnapshot::segmentChanged(a, longMut.join(';'),
                                                      "metric_defs"));
    }

    // Incremental regen produces a snapshot identical to a full rebuild.
    void testIncrementalMatchesFullRebuild() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));

        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));   // full build #1

        // Mutate a data_row so the prior snapshot is stale (data-only change).
        {
            QSqlQuery up(pg.queryDb());
            QVERIFY2(up.exec("UPDATE data_rows SET notes='bumped', "
                             "updated_at = now() + interval '10 seconds' "
                             "WHERE id = (SELECT min(id) FROM data_rows)"),
                     qPrintable(up.lastError().text()));
        }

        // Incremental regen over the (now stale) prior snapshot.
        DVE::OfflineSnapshot::RegenStats st;
        QString fp, err; QDateTime srv;
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, snap.path(), nullptr,
                     &fp, &srv, &err, {}, &st), qPrintable(err));
        QVERIFY(st.wasIncremental);

        // A fresh full rebuild of the SAME pg state to a different path.
        const QString full2 = overrideBaseDir() + "/full2.sqlite";
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, full2, nullptr,
                     &fp, &srv, &err, {}, nullptr), qPrintable(err));

        for (const char* t : DVE::OfflineSnapshot::kFingerprintTables)
            QCOMPARE(tableSig(snap.path(), t), tableSig(full2, t));
        QCOMPARE(tableSig(snap.path(), "settings"), tableSig(full2, "settings"));
        pg.close();
    }

    // A schema-version mismatch in the prior snapshot forces a full rebuild.
    void testIncrementalFallsBackOnSchemaMismatch() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));

        // Corrupt the stored schema version so incremental is rejected.
        {
            const QString conn = "tst_oss_corrupt";
            {
                QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", conn);
                db.setDatabaseName(snap.path());
                QVERIFY(db.open());
                QSqlQuery q(db);
                QVERIFY(q.exec("UPDATE _snapshot_meta SET value='0' "
                               "WHERE key='source_schema_version'"));
                db.close();
            }
            QSqlDatabase::removeDatabase(conn);
        }

        DVE::OfflineSnapshot::RegenStats st;
        QString fp, err; QDateTime srv;
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, snap.path(), nullptr,
                     &fp, &srv, &err, {}, &st), qPrintable(err));
        QVERIFY(!st.wasIncremental);
        pg.close();
    }

    // THE WIN: a data-only change re-pulls ZERO image blobs from PG, and the
    // snapshot's images are byte-identical to before.
    void testDataOnlyChangePullsZeroImageBlobs() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        const QString before = tableSig(snap.path(), "images");

        {
            QSqlQuery up(pg.queryDb());
            QVERIFY2(up.exec("UPDATE data_rows SET notes='x', "
                             "updated_at = now() + interval '5 seconds' "
                             "WHERE id = (SELECT min(id) FROM data_rows)"),
                     qPrintable(up.lastError().text()));
        }
        DVE::OfflineSnapshot::RegenStats st;
        QString fp, err; QDateTime srv;
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, snap.path(), nullptr,
                     &fp, &srv, &err, {}, &st), qPrintable(err));
        QVERIFY(st.wasIncremental);
        QCOMPARE(st.imageRowsPulledFromPg, 0);                 // no blobs re-read
        QCOMPARE(tableSig(snap.path(), "images"), before);     // images untouched
        pg.close();
    }

    // An image-table change pulls ONLY the new/changed blob(s), not every image,
    // and the result matches a full rebuild. The change is an INSERT (a count
    // change), which the fingerprint detects regardless of its whole-second
    // updated_at granularity -- a same-second UPDATE can be missed, which is the
    // accepted pre-existing limitation of the snapshot's freshness fingerprint.
    void testImageChangePullsOnlyChanged() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        // Add a 2nd image so "only the new one" is meaningful (seed has 1).
        {
            QSqlQuery ins(pg.queryDb());
            QVERIFY2(ins.exec("INSERT INTO images (sample_id, sort_order, file_name, "
                              "image_data, updated_at, updated_by, version) "
                              "SELECT min(id), 1, 'img2.png', '\\x0102'::bytea, "
                              "now(), 'tester', 1 FROM samples"),
                     qPrintable(ins.lastError().text()));
        }
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));   // full, 2 images

        // Add a 3rd image; the 2 already-snapshotted images are unchanged.
        {
            QSqlQuery ins(pg.queryDb());
            QVERIFY2(ins.exec("INSERT INTO images (sample_id, sort_order, file_name, "
                              "image_data, updated_at, updated_by, version) "
                              "SELECT min(id), 2, 'img3.png', '\\x03040506'::bytea, "
                              "now(), 'tester', 1 FROM samples"),
                     qPrintable(ins.lastError().text()));
        }
        DVE::OfflineSnapshot::RegenStats st;
        QString fp, err; QDateTime srv;
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, snap.path(), nullptr,
                     &fp, &srv, &err, {}, &st), qPrintable(err));
        QVERIFY(st.wasIncremental);
        QCOMPARE(st.imageRowsPulledFromPg, 1);                 // only the new one (not all 3)

        const QString full2 = overrideBaseDir() + "/full2img.sqlite";
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, full2, nullptr,
                     &fp, &srv, &err, {}, nullptr), qPrintable(err));
        QCOMPARE(tableSig(snap.path(), "images"), tableSig(full2, "images"));
        pg.close();
    }

    // ===================================================================
    // v3 Phase 3c: open metrics survive offline (snapshot schema v4)
    // ===================================================================

    // THE POINT OF TASK 4. DataRow::extra / SampleResult::extra now ride the
    // save transaction into measurements / sample_headers and come back out of
    // DatabaseManager::loadFile -- but the offline snapshot mirrored neither
    // table, so going offline silently erased every custom column. Mirror the
    // three long-format tables, copy them on regen, and decode them on the
    // offline read.
    //
    // RED before the fix: the snapshot has no `measurements` table at all
    // (sqliteCount returns -1) and every `extra` map comes back empty.
    void v3OpenMetrics_surviveRegenAndOfflineRead() {
        REQUIRE_PG();
        seedPostgresFixture();
        seedOpenMetricsFixture();

        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        const QString snapPath = snap.path();
        pg.close();

        // 1) The regen copied the rows (independent of the read path).
        QCOMPARE(sqliteCount(snapPath, "measurements"),   5);
        QCOMPARE(sqliteCount(snapPath, "sample_headers"), 2);
        // metric_defs carries the whole shared vocabulary, not just our six.
        QVERIFY(sqliteCount(snapPath, "metric_defs") >= 6);

        // 2) The offline read hands them back on the right rows, TYPED.
        QVERIFY2(snap.openReadOnly(), qPrintable(snap.lastError()));
        const DVE::FileResult fr = snap.loadFileByPath("/tmp/seed.xlsx");
        QCOMPARE(fr.sheets.size(), 1);
        QCOMPARE(fr.sheets[0].samples.size(), 2);

        const DVE::SampleResult& s0 = fr.sheets[0].samples[0];
        QCOMPARE(s0.rows.size(), 2);

        const QMap<QString, QVariant>& e0 = s0.rows[0].extra;
        QCOMPARE(e0.size(), 4);
        QCOMPARE(e0.value("snap_coil_temp").typeId(), int(QMetaType::Double));
        QCOMPARE(e0.value("snap_coil_temp").toDouble(), 42.5);
        QCOMPARE(e0.value("snap_rig_note").typeId(), int(QMetaType::QString));
        QCOMPARE(e0.value("snap_rig_note").toString(), QString("rig A"));
        QCOMPARE(e0.value("snap_flag").typeId(), int(QMetaType::Bool));
        QCOMPARE(e0.value("snap_flag").toBool(), true);
        QCOMPARE(e0.value("snap_pv").typeId(), int(QMetaType::QVariantList));
        const QVariantList pv = e0.value("snap_pv").toList();
        QCOMPARE(pv.size(), 3);
        QCOMPARE(pv[0].toDouble(), 1.0);
        QCOMPARE(pv[1].toDouble(), 2.5);
        QCOMPARE(pv[2].toDouble(), 3.0);

        // Row 1 must get ITS value, not row 0's -- (sample_id, sort_order) is
        // the only join measurements has to a data row (hazard H22).
        QCOMPARE(s0.rows[1].extra.size(), 1);
        QCOMPARE(s0.rows[1].extra.value("snap_coil_temp").toDouble(), 43.5);

        // Sample headers, likewise typed.
        QCOMPARE(s0.extra.size(), 2);
        QCOMPARE(s0.extra.value("snap_rig_id").typeId(), int(QMetaType::QString));
        QCOMPARE(s0.extra.value("snap_rig_id").toString(), QString("RIG-7"));
        QCOMPARE(s0.extra.value("snap_coil_count").typeId(), int(QMetaType::Double));
        QCOMPARE(s0.extra.value("snap_coil_count").toDouble(), 3.0);

        // 3) The sparse rule reads the way it writes: a sample with nothing
        //    stored keeps the empty maps it was constructed with.
        const DVE::SampleResult& s1 = fr.sheets[0].samples[1];
        QVERIFY(s1.extra.isEmpty());
        for (const DVE::DataRow& dr : s1.rows) QVERIFY(dr.extra.isEmpty());

        // 4) The standard wide columns are untouched by any of this.
        QCOMPARE(s0.rows[0].notes, QString("note s0r0"));
        QCOMPARE(s0.sampleName, QString("Sample 1"));
        snap.close();
    }

    // The version bump is what stops a v3 snapshot -- which has no
    // measurements / sample_headers table at all -- from being reused under the
    // v4 shape. Both gates that read kSnapshotSchemaVersion must hold: the
    // read-side reject in openReadOnly(), and the incremental gate in
    // regenToPath() (which copies the prior file WHOLESALE, schema included,
    // and so would otherwise inherit the v3 layout).
    //
    // RED before the fix: kSnapshotSchemaVersion is still 3, so the stamp is
    // "3" and a v3-stamped snapshot opens and is reused incrementally.
    void v3OpenMetrics_schemaVersionFourInvalidatesV3Snapshots() {
        REQUIRE_PG();
        seedPostgresFixture();
        seedOpenMetricsFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        const QString snapPath = snap.path();

        QCOMPARE(sqliteScalar(snapPath,
                     "SELECT value FROM _snapshot_meta "
                     "WHERE key='source_schema_version'"),
                 QString("4"));

        // Restamp as v3 -- exactly what every snapshot on a user's machine
        // looks like right now.
        {
            const QString cname = QStringLiteral("tst_oss_v3stamp");
            {
                QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname);
                db.setDatabaseName(snapPath);
                QVERIFY(db.open());
                QSqlQuery q(db);
                QVERIFY2(q.exec("UPDATE _snapshot_meta SET value='3' "
                                "WHERE key='source_schema_version'"),
                         qPrintable(q.lastError().text()));
                db.close();
            }
            QSqlDatabase::removeDatabase(cname);
        }

        // (a) Read side: refuse to open it.
        DVE::OfflineSnapshot stale;
        stale.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(!stale.openReadOnly(),
                 "a v3-stamped snapshot must not open under schema v4");
        QVERIFY2(stale.lastError().contains("schema"),
                 qPrintable("expected a schema-mismatch error; got: "
                            + stale.lastError()));

        // (b) Regen side: a FULL rebuild, never an incremental reuse of the
        //     v3 shape -- and the rebuilt file has the v4 tables populated.
        DVE::OfflineSnapshot::RegenStats st;
        QString fp, err; QDateTime srv;
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, snapPath, nullptr,
                     &fp, &srv, &err, {}, &st), qPrintable(err));
        QVERIFY2(!st.wasIncremental,
                 "a v3 snapshot must force a full rebuild, not be reused");
        QCOMPARE(sqliteCount(snapPath, "measurements"), 5);
        pg.close();
    }

    // The incremental path DELETEs and re-copies the long tables like every
    // other non-blob table, so a change to an open metric has to show up in the
    // mirror -- and the result has to match a full rebuild exactly. This is the
    // same equivalence testIncrementalMatchesFullRebuild pins for the nine
    // pre-3c tables, narrowed onto the three new ones.
    void v3OpenMetrics_incrementalRefreshesLongTables() {
        REQUIRE_PG();
        seedPostgresFixture();
        seedOpenMetricsFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));

        // Drop one measurement and change another: the mirror must follow.
        {
            QSqlQuery up(pg.queryDb());
            QVERIFY2(up.exec("DELETE FROM measurements WHERE sort_order = 1"),
                     qPrintable(up.lastError().text()));
            QVERIFY2(up.exec("UPDATE measurements SET value_num = 99.5, "
                             "updated_at = now() + interval '10 seconds' "
                             "WHERE value_num = 42.5"),
                     qPrintable(up.lastError().text()));
        }

        DVE::OfflineSnapshot::RegenStats st;
        QString fp, err; QDateTime srv;
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, snap.path(), nullptr,
                     &fp, &srv, &err, {}, &st), qPrintable(err));
        QVERIFY(st.wasIncremental);
        QCOMPARE(sqliteCount(snap.path(), "measurements"), 4);

        const QString full2 = overrideBaseDir() + "/full2long.sqlite";
        QVERIFY2(DVE::OfflineSnapshot::regenToPath(&pg, full2, nullptr,
                     &fp, &srv, &err, {}, nullptr), qPrintable(err));
        for (const char* t : { "metric_defs", "measurements", "sample_headers" })
            QCOMPARE(tableSig(snap.path(), t), tableSig(full2, t));

        QVERIFY2(snap.openReadOnly(), qPrintable(snap.lastError()));
        const DVE::FileResult fr = snap.loadFileByPath("/tmp/seed.xlsx");
        QCOMPARE(fr.sheets[0].samples[0].rows[0].extra.value("snap_coil_temp")
                     .toDouble(), 99.5);
        QVERIFY(fr.sheets[0].samples[0].rows[1].extra.isEmpty());
        snap.close();
        pg.close();
    }

    // -- T3: read-only enforcement at SQLite layer --------------------------
    void testReadOnlyEnforcesReadOnly() {
        REQUIRE_PG();
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
        REQUIRE_PG();
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

    // -- SP4.5: isCurrentVsLive drives the skip-unchanged-regen optimization --
    // The close-time regen is skipped when the live DB is unchanged since the
    // last snapshot (cheap COUNT + MAX(updated_at) fingerprint). Verify: current
    // right after regenerate(); stale after a live INSERT (count changes); current
    // again after a re-regenerate.
    void isCurrentVsLive_tracksLiveChanges() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        QVERIFY(snap.openReadOnly());     // so storedContentFingerprint() can read m_db

        // 1) Nothing changed since regen → snapshot current → safe to skip.
        QVERIFY2(snap.isCurrentVsLive(&pg),
                 "expected snapshot current immediately after regenerate()");

        // 2) Mutate the live DB (count changes) → fingerprint differs → stale.
        {
            QSqlQuery q(pg.queryDb());
            QVERIFY2(q.exec("INSERT INTO sensory_sessions "
                            "(session_name, tester_name, date, json_data, updated_by, version) "
                            "VALUES ('fp_change','t','2026-06-18','{\"samples\":[]}'::jsonb,'t',1)"),
                     qPrintable(q.lastError().text()));
        }
        QVERIFY2(!snap.isCurrentVsLive(&pg),
                 "expected snapshot stale after a live INSERT");

        // 3) Regenerate → current again (regenerate() closes m_db, so reopen).
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        QVERIFY(snap.openReadOnly());
        QVERIFY2(snap.isCurrentVsLive(&pg),
                 "expected snapshot current after re-regenerate()");

        snap.close();
        pg.close();
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
        REQUIRE_PG();
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

    // -- R7b: snapshot copies the app_version era column --------------------
    // SP1 added a nullable app_version column to files / sensory_sessions /
    // detailed_sensory_sessions in Postgres. SP3-T1 (R7b) extends the snapshot
    // SELECT/INSERT lists so the column round-trips into SQLite, feeding SP4's
    // offline era display. RED before the fix: the SQLite schema lacks the
    // column, so the SELECT below errors / returns null.
    void snapshot_carriesAppVersionColumn() {
        REQUIRE_PG();
        const QString expected = seedAppVersionFixture();

        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        pg.close();

        const QString snapPath = snap.path();
        QVERIFY(QFile::exists(snapPath));

        // Verify app_version copied for all three tables (direct SQLite read,
        // independent of the OfflineSnapshot read accessors).
        QCOMPARE(sqliteScalar(snapPath,
                              "SELECT app_version FROM files "
                              "WHERE file_path = '/tmp/appver.xlsx'"),
                 expected);
        QCOMPARE(sqliteScalar(snapPath,
                              "SELECT app_version FROM sensory_sessions "
                              "WHERE session_name = 'AppVer Sensory'"),
                 expected);
        QCOMPARE(sqliteScalar(snapPath,
                              "SELECT app_version FROM detailed_sensory_sessions "
                              "WHERE session_name = 'AppVer Detailed'"),
                 expected);
    }

    // -- T8: per-cell pending-edit queue round-trip (v2.0.1 Task 9) ---------
    // Verifies enqueueCellEdit + drainPendingEdits without requiring a
    // PG connection (the queue file is independent of the read-only
    // snapshot file).
    void cellEditsRoundTrip() {
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());

        QVERIFY(snap.enqueueCellEdit(QStringLiteral("data_rows"), 7,
                                     QStringLiteral("draw_pressure"),
                                     QVariant(1.5)));
        QVERIFY(snap.enqueueCellEdit(
            QStringLiteral("sensory_sessions"), 42,
            QStringLiteral("json_path:samples[0].voltage"), QVariant(3.7)));

        QCOMPARE(snap.pendingEditCount(), 2);

        QVector<QStringList> applied;
        const int n = snap.drainPendingEdits(
            [&](const QString& t, qint64 r, const QString& c, const QVariant& v,
                int, int) {
                applied << QStringList{t, QString::number(r), c, v.toString()};
                return true;
            });
        QCOMPARE(n, 2);
        QCOMPARE(applied.size(), 2);
        QCOMPARE(applied[0][0], QStringLiteral("data_rows"));
        QCOMPARE(applied[0][1], QStringLiteral("7"));
        QCOMPARE(applied[0][2], QStringLiteral("draw_pressure"));
        QCOMPARE(applied[1][0], QStringLiteral("sensory_sessions"));
        QCOMPARE(applied[1][2],
                 QStringLiteral("json_path:samples[0].voltage"));

        // Second drain must be empty (applied rows were DELETEd).
        const int n2 = snap.drainPendingEdits(
            [](const QString&, qint64, const QString&, const QVariant&, int, int) {
                return true;
            });
        QCOMPARE(n2, 0);
        QCOMPARE(snap.pendingEditCount(), 0);
    }

    // -- v3 Phase 3d (hazard H9): the schema_version=2 measurement flavor ----
    // A v2 row carries the natural (sample id, metric key, ordinal) identity;
    // a v1 row drains with the sentinel defaults (schemaVersion 1, sortOrder
    // -1) so LiveSync::flushPending can route the generations.
    void measurementEditsCarrySchemaVersionAndSortOrder() {
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());

        QVERIFY(snap.enqueueMeasurementEdit(41, QStringLiteral("smell"), 3,
                                            QVariant(QStringLiteral("2"))));
        QVERIFY(snap.enqueueCellEdit(QStringLiteral("sensory_sessions"), 9,
                                     QStringLiteral("assessor_name"),
                                     QVariant(QStringLiteral("legacy"))));
        QCOMPARE(snap.pendingEditCount(), 2);

        struct Seen { QString t; qint64 r; QString c; QString v; int sv; int so; };
        QVector<Seen> seen;
        const int n = snap.drainPendingEdits(
            [&](const QString& t, qint64 r, const QString& c, const QVariant& v,
                int sv, int so) {
                seen.append({t, r, c, v.toString(), sv, so});
                return true;
            });
        QCOMPARE(n, 2);
        QCOMPARE(seen.size(), 2);

        QCOMPARE(seen[0].t,  QStringLiteral("measurements"));
        QCOMPARE(seen[0].r,  qint64(41));      // the SAMPLE id
        QCOMPARE(seen[0].c,  QStringLiteral("smell"));
        QCOMPARE(seen[0].v,  QStringLiteral("2"));
        QCOMPARE(seen[0].sv, 2);
        QCOMPARE(seen[0].so, 3);

        QCOMPARE(seen[1].t,  QStringLiteral("sensory_sessions"));
        QCOMPARE(seen[1].sv, 1);               // legacy default
        QCOMPARE(seen[1].so, -1);
        QCOMPARE(snap.pendingEditCount(), 0);
    }

    // -- T9: partial-failure replay — failing callbacks leave rows queued ---
    void cellEditsPartialFailureKeepsRows() {
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());

        QVERIFY(snap.enqueueCellEdit(QStringLiteral("data_rows"), 1,
                                     QStringLiteral("notes"),
                                     QVariant(QStringLiteral("first"))));
        QVERIFY(snap.enqueueCellEdit(QStringLiteral("data_rows"), 2,
                                     QStringLiteral("notes"),
                                     QVariant(QStringLiteral("second"))));
        QCOMPARE(snap.pendingEditCount(), 2);

        // Apply succeeds only for rowId=1; rowId=2 stays queued.
        const int n = snap.drainPendingEdits(
            [](const QString&, qint64 r, const QString&, const QVariant&, int, int) {
                return r == 1;
            });
        QCOMPARE(n, 1);
        QCOMPARE(snap.pendingEditCount(), 1);

        // Second drain — the leftover row should apply this time.
        QVector<qint64> seen;
        const int n2 = snap.drainPendingEdits(
            [&](const QString&, qint64 r, const QString&, const QVariant&, int, int) {
                seen.append(r);
                return true;
            });
        QCOMPARE(n2, 1);
        QCOMPARE(seen.size(), 1);
        QCOMPARE(seen[0], qint64(2));
        QCOMPARE(snap.pendingEditCount(), 0);
    }

    // -- C5: drainPendingEdits must skip rows already flagged as replayed.
    //        Before the fix, the SELECT had no replayed_at filter, so a row
    //        marked as replayed (e.g., by a previous drain whose bulk DELETE
    //        failed and fell back to UPDATE) would be re-applied on every
    //        subsequent drain — silently double-writing the cell to Postgres.
    //        After C5, drainPendingEdits' SELECT filters AND replayed_at IS
    //        NULL and ensureQueueOpen adds the column via additive ALTER.
    void drainPendingEdits_respectsReplayedAtSentinel() {
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());

        QVERIFY(snap.enqueueCellEdit(QStringLiteral("data_rows"), 1,
                                     QStringLiteral("notes"),
                                     QVariant(QStringLiteral("first"))));
        QVERIFY(snap.enqueueCellEdit(QStringLiteral("data_rows"), 2,
                                     QStringLiteral("notes"),
                                     QVariant(QStringLiteral("second"))));
        QVERIFY(snap.enqueueCellEdit(QStringLiteral("data_rows"), 3,
                                     QStringLiteral("notes"),
                                     QVariant(QStringLiteral("third"))));
        QCOMPARE(snap.pendingEditCount(), 3);

        // Inject replayed_at on row_id=2 via a sidecar SQLite connection.
        // Before C5, this column doesn't exist → q.exec returns false →
        // test RED at the QVERIFY2 below.
        {
            const QString qpath = snap.queuePath();
            const QString cname = QStringLiteral("tst_c5_inject");
            QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname);
            db.setDatabaseName(qpath);
            QVERIFY(db.open());
            QSqlQuery q(db);
            QVERIFY2(q.exec("UPDATE pending_edits SET replayed_at = "
                            "'2026-01-01T00:00:00Z' WHERE row_id = 2"),
                     qPrintable(q.lastError().text()));
            db.close();
            QSqlDatabase::removeDatabase(cname);
        }

        QVector<qint64> seen;
        const int n = snap.drainPendingEdits(
            [&](const QString&, qint64 r, const QString&,
                const QVariant&, int, int) {
                seen.append(r);
                return true;
            });

        // Only rows 1 and 3 should be drained; row 2 is sentinel-skipped.
        QCOMPARE(n, 2);
        QCOMPARE(seen.size(), 2);
        QVERIFY( seen.contains(qint64(1)));
        QVERIFY(!seen.contains(qint64(2)));
        QVERIFY( seen.contains(qint64(3)));
    }

    // -- T10: raw_grid (SOP sheet) round-trip through OfflineSnapshot -------
    // Verifies that OfflineSnapshot::regenerate copies the raw_grid TEXT from
    // Postgres and that OfflineSnapshot::loadFile reconstructs rawHeaders /
    // rawRows via rawGridFromJson. This mirrors the DatabaseManager round-trip
    // added in Task 4 (feat(db): persist + reconstruct SOP raw_grid).
    void testRawGridRoundTrip() {
        REQUIRE_PG();
        // 1. Seed Postgres: one file with one raw SOP sheet carrying a
        //    known raw_grid JSON.
        const QStringList expectedHeaders = {"Product", "Lot", "Pass/Fail", "Notes"};
        const QVector<QStringList> expectedRows = {
            {"Alpha-1",  "LOT-001", "Pass", "OK"},
            {"Beta-2",   "LOT-002", "Fail", "Burnt"},
            {"Gamma-3",  "LOT-003", "Pass", ""},
        };
        const QString rawGridJson =
            DVE::rawGridToJson(expectedHeaders, expectedRows);
        QVERIFY(!rawGridJson.isEmpty());

        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_oss_raw_seed";
        qint64 fileId = -1;
        {
            QSqlDatabase pg = QSqlDatabase::addDatabase("QPSQL", cname);
            pg.setHostName(cfg.host);
            pg.setPort(cfg.port);
            pg.setDatabaseName(cfg.database);
            pg.setUserName(cfg.user);
            pg.setPassword(cfg.password);
            QVERIFY2(pg.open(), qPrintable(pg.lastError().text()));

            QSqlQuery q(pg);
            // file row
            q.prepare("INSERT INTO files (file_path, file_name, loaded_at, "
                      "template_version, sheet_count, sample_count, updated_by) "
                      "VALUES (?, ?, ?, ?, ?, ?, ?) RETURNING id");
            q.addBindValue("/tmp/sop_raw.xlsx");
            q.addBindValue("sop_raw.xlsx");
            q.addBindValue(QDateTime::currentDateTimeUtc().toString(Qt::ISODate));
            q.addBindValue("new");
            q.addBindValue(1);
            q.addBindValue(0);
            q.addBindValue("test-raw-seeder");
            MUST(q.exec() && q.next());
            fileId = q.value(0).toLongLong();

            // test row with is_raw_table=1 and the raw_grid JSON
            q.prepare("INSERT INTO tests (file_id, sheet_name, template_version, "
                      "overall_avg_tpm, overall_stddev_tpm, is_raw_table, sort_order, "
                      "raw_grid, updated_by) "
                      "VALUES (?, ?, ?, ?, ?, ?, ?, CAST(? AS JSONB), ?) RETURNING id");
            q.addBindValue(static_cast<qlonglong>(fileId));
            q.addBindValue("Test SOP's");
            q.addBindValue("new");
            q.addBindValue(0.0);
            q.addBindValue(0.0);
            q.addBindValue(1);    // is_raw_table
            q.addBindValue(0);    // sort_order
            q.addBindValue(rawGridJson);
            q.addBindValue("test-raw-seeder");
            MUST(q.exec() && q.next());

            pg.close();
        }
        QSqlDatabase::removeDatabase(cname);
        QVERIFY(fileId > 0);

        // 2. Regenerate snapshot from live Postgres.
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        pg.close();

        // 3. Verify the raw_grid TEXT column was copied into SQLite.
        {
            const QString cname2 = "tst_oss_raw_verify";
            QString gotJson;
            {
                QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname2);
                db.setDatabaseName(snap.path());
                QVERIFY(db.open());
                QSqlQuery q(db);
                QVERIFY(q.exec("SELECT raw_grid FROM tests WHERE is_raw_table = 1"));
                QVERIFY2(q.next(), "Expected one raw test row in SQLite snapshot");
                gotJson = q.value(0).toString();
                db.close();
            }
            QSqlDatabase::removeDatabase(cname2);
            QVERIFY2(!gotJson.isEmpty(), "raw_grid should be non-empty for raw sheet");
        }

        // 4. Open read-only and assert rawHeaders / rawRows round-trip.
        QVERIFY2(snap.openReadOnly(), qPrintable(snap.lastError()));
        const DVE::FileResult fr = snap.loadFileByPath("/tmp/sop_raw.xlsx");
        QVERIFY(!fr.filePath.isEmpty());
        QCOMPARE(fr.sheets.size(), 1);
        const DVE::SheetResult& sheet = fr.sheets[0];
        QVERIFY(sheet.isRawTable);
        QCOMPARE(sheet.sheetName, QString("Test SOP's"));

        // Headers must survive exactly.
        QCOMPARE(sheet.rawHeaders, expectedHeaders);

        // Rows must survive exactly.
        QCOMPARE(sheet.rawRows.size(), expectedRows.size());
        for (int r = 0; r < expectedRows.size(); ++r) {
            QCOMPARE(sheet.rawRows[r], expectedRows[r]);
        }
        snap.close();
    }

    // -- R7 (a): openReadOnly rejects a stale snapshot schema version -------
    // SP3-T1 bumped kSnapshotSchemaVersion to 3 and writes it into
    // _snapshot_meta. SP3-T2 (R7) makes openReadOnly() validate that field:
    // a snapshot stamped with a different (older) source_schema_version must
    // be treated as no usable snapshot so the app regenerates on the next
    // clean online close rather than trusting a layout it can't interpret.
    // RED before the fix: openReadOnly() ignores source_schema_version and
    // opens a v2-stamped file anyway.
    void snapshot_rejectsStaleSchemaVersion() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        pg.close();

        const QString snapPath = snap.path();
        QVERIFY(QFile::exists(snapPath));

        // Directly downgrade the recorded schema version to simulate a
        // snapshot left behind by an older client (e.g. a v2.4.1 install that
        // wrote schema v2). The on-disk layout is irrelevant to the test; the
        // version field alone must trigger the reject.
        {
            const QString cname = QStringLiteral("tst_oss_downgrade");
            {
                QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", cname);
                db.setDatabaseName(snapPath);
                QVERIFY(db.open());
                QSqlQuery q(db);
                QVERIFY2(q.exec("UPDATE _snapshot_meta SET value='2' "
                                "WHERE key='source_schema_version'"),
                         qPrintable(q.lastError().text()));
                db.close();
            }
            QSqlDatabase::removeDatabase(cname);
        }

        // A fresh OfflineSnapshot pointed at the same (now-stale) file must
        // refuse to open it.
        DVE::OfflineSnapshot stale;
        stale.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(!stale.openReadOnly(),
                 "openReadOnly() must reject a snapshot whose "
                 "source_schema_version does not match the current build");
        QVERIFY(!stale.isOpen());
        QVERIFY(!stale.lastError().isEmpty());
        QVERIFY2(stale.lastError().contains("schema"),
                 qPrintable("Expected a schema-mismatch error; got: "
                            + stale.lastError()));
    }

    // -- R7 (b): the freshness stamp comes from the PG server clock ---------
    // Before SP3-T2 the snapshot stamped QDateTime::currentDateTimeUtc()
    // (the local client clock, which may be NTP-skewed). After the fix the
    // stamp is captured from the SOURCE Postgres connection (SELECT now()),
    // so the offline-era display reflects the authoritative server time.
    //
    // Deterministic discriminator (not dependent on incidental client/server
    // skew): regenerate() records the exact server timestamp it captured,
    // exposed for tests via lastRegenServerTimeUtc(). The persisted stamp read
    // back by snapshotTakenAt() must equal that server value to the
    // millisecond. RED today: lastRegenServerTimeUtc() is unset (invalid) and
    // the stamp is the client clock, so the equality fails.
    void snapshot_stampIsServerClock() {
        REQUIRE_PG();
        seedPostgresFixture();
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));

        // Bracket regeneration with two direct server-clock reads on the SAME
        // live connection so the window is server-time, not client-time.
        const QDateTime serverBefore = serverNowUtc(pg);
        QVERIFY2(serverBefore.isValid(), "could not read server now() (before)");

        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));

        const QDateTime serverAfter = serverNowUtc(pg);
        QVERIFY2(serverAfter.isValid(), "could not read server now() (after)");
        pg.close();

        // The server timestamp regenerate captured must be exposed and valid.
        const QDateTime captured = snap.lastRegenServerTimeUtc();
        QVERIFY2(captured.isValid(),
                 "regenerate() did not capture a server timestamp "
                 "(stamp is still sourced from the client clock)");

        QVERIFY(snap.openReadOnly());
        const QDateTime stamp = snap.snapshotTakenAt();
        QVERIFY2(stamp.isValid(), "snapshotTakenAt() returned an invalid QDateTime");

        // The persisted stamp must be EXACTLY the server time regenerate read,
        // not the client clock.
        QCOMPARE(stamp.toUTC(), captured.toUTC());

        // Independent cross-check: that server timestamp must fall inside the
        // bracketing server-clock window (proves it really is server time).
        QVERIFY2(captured.toUTC() >= serverBefore.toUTC().addSecs(-1) &&
                 captured.toUTC() <= serverAfter.toUTC().addSecs(1),
                 qPrintable(QString("captured server stamp %1 outside "
                                    "server window [%2, %3]")
                                .arg(captured.toString(Qt::ISODateWithMs),
                                     serverBefore.toString(Qt::ISODateWithMs),
                                     serverAfter.toString(Qt::ISODateWithMs))));
        snap.close();
    }

    // -- T7: loadFileByPath round-trip + version anchor ---------------------
    void testLoadFileByPathRoundTrip() {
        REQUIRE_PG();
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

    // ── R5 (v2.4.4): MIP-resilient snapshot + pending-edits queue ───────────
    //
    // These cases craft files directly (no Postgres needed) so they pin the
    // detection + fallback-branch + LOUD-failure wiring even on a machine where
    // true MIP decryption can't run.
    //
    // SCOPE NOTE: true MIP decryption is only verifiable on the work machine
    // (the bundled python must be MIP-allowlisted AND the file must carry a real
    // sensitivity label). Here we write the marker "%TSD-Header-###%" followed
    // by junk -- the fallback python copies the bytes but they are not a valid
    // SQLite db -- so we assert the store TAKES the fallback branch and SURFACES
    // a loud decode failure (openReadOnly()==false + lastOpenWasDecodeFailure(),
    // enqueueCellEdit()==false) rather than silently behaving as "no snapshot".

    void mipFallback_openReadOnly_encryptedSnapshotSurfacesDecodeFailure() {
        // No PG dependency: write a fake encrypted snapshot at the snapshot path.
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        const QString p = snap.path();   // mkpaths the dir
        {
            // Explicit-length QByteArray: a "\x00..." C-string write would
            // truncate at the NUL. We want genuine binary junk after the marker.
            QByteArray bytes("%TSD-Header-###%");        // MIP marker
            const char junk[] = { '\x10', '\x20', '\x30', '\x00', ' ', 'n', 'o',
                                  't', ' ', 's', 'q', 'l', 'i', 't', 'e' };
            bytes.append(junk, int(sizeof(junk)));
            QFile f(p);
            QVERIFY(f.open(QIODevice::WriteOnly));
            QVERIFY(f.write(bytes) == bytes.size());
            f.close();
        }
        QVERIFY(DVE::looksEncrypted(p));

        // openReadOnly must NOT succeed and must NOT look like mere absence:
        // it took the decrypt branch, which produced non-SQLite bytes -> false
        // + a LOUD decode-failure flag the MainWindow caller warns on.
        QVERIFY2(!snap.openReadOnly(), "encrypted/garbage snapshot must not open");
        QVERIFY2(snap.lastOpenWasDecodeFailure(),
                 "an encrypted-but-undecodable snapshot must surface a loud "
                 "DECODE FAILURE (the signal the MainWindow caller warns on), "
                 "not behave like 'no snapshot present'");
        QVERIFY(!snap.lastError().isEmpty());
        snap.close();
    }

    void mipFallback_openReadOnly_absentSnapshotIsNotDecodeFailure() {
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        // Do NOT create the file -> genuine absence.
        QVERIFY(!snap.openReadOnly());
        QVERIFY2(!snap.lastOpenWasDecodeFailure(),
                 "a genuinely-absent snapshot is first-run absence, NOT a decode "
                 "failure (the caller must stay silent in that case)");
        snap.close();
    }

    void mipFallback_enqueueCellEdit_encryptedQueueFailsLoudNotSilent() {
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        // Write a fake encrypted pending_edits queue at its path.
        const QString qp = snap.queuePath();
        QDir().mkpath(QFileInfo(qp).absolutePath());
        {
            // Explicit-length QByteArray so an embedded NUL doesn't truncate
            // the junk written after the marker.
            QByteArray bytes("%TSD-Header-###%");
            const char junk[] = { '\x01', '\x02', '\x00', ' ', 'n', 'o', 't',
                                  ' ', 's', 'q', 'l', 'i', 't', 'e' };
            bytes.append(junk, int(sizeof(junk)));
            QFile f(qp);
            QVERIFY(f.open(QIODevice::WriteOnly));
            QVERIFY(f.write(bytes) == bytes.size());
            f.close();
        }
        QVERIFY(DVE::looksEncrypted(qp));

        // enqueueCellEdit -> ensureQueueOpen -> decrypt branch -> garbage db ->
        // open fails -> enqueueCellEdit returns FALSE (the bool the LiveSync
        // call sites now act on instead of discarding).
        const bool ok = snap.enqueueCellEdit("files", 1,
                                             "file_name", QVariant("x.xlsx"));
        QVERIFY2(!ok, "an undecodable encrypted queue must make enqueueCellEdit "
                      "return false (the failure is surfaced, not swallowed)");
        QVERIFY(!snap.lastError().isEmpty());

        // R5 IMPORTANT 3: pendingEditCount() reports 0 on an undecodable queue
        // (it can't open it to count), but queueDegraded() must be TRUE so the
        // caller does NOT mistake the false 0 for "queue empty / fully drained".
        QCOMPARE(snap.pendingEditCount(), 0);
        QVERIFY2(snap.queueDegraded(),
                 "a present-but-undecodable queue must report queueDegraded(), "
                 "not look empty");
        snap.close();
    }

    // R5 IMPORTANT 3 complement: a genuinely-absent queue (first run, no offline
    // edits yet) is NOT degraded -- pendingEditCount()==0 here really does mean
    // "empty". A clean enqueue then clears any degraded state.
    void mipFallback_absentQueueIsNotDegraded() {
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        // No queue file written -> SQLite creates a fresh one on first open.
        QCOMPARE(snap.pendingEditCount(), 0);
        QVERIFY2(!snap.queueDegraded(),
                 "an absent queue is genuinely empty, not degraded");
        QVERIFY(snap.enqueueCellEdit("data_rows", 5, "notes", QVariant("ok")));
        QCOMPARE(snap.pendingEditCount(), 1);
        QVERIFY(!snap.queueDegraded());
        snap.close();
    }

    // ====================================================================
    //  SP4.5 Stage 2a -- SnapshotRegenWorker (background regen)
    // ====================================================================

    // The worker runs regenToPath on its own thread + own PostgresConnection and
    // produces a valid, openable snapshot file.
    void tst_regenWorker_backgroundRegen() {
        DVE::DbConfig cfg = pgConfig();
        if (cfg.host.isEmpty()) QSKIP("DVE_TEST_PG_CONN not set");

        const QString snap = QDir::tempPath()
            + QStringLiteral("/dve_regen_%1.sqlite").arg(QDateTime::currentMSecsSinceEpoch());

        QThread thread;
        DVE::SnapshotRegenWorker worker(cfg, snap);
        worker.moveToThread(&thread);
        thread.start();
        QMetaObject::invokeMethod(&worker, "start", Qt::QueuedConnection);

        bool ok = false;
        QString err = QStringLiteral("no signal");
        QEventLoop loop;
        QObject::connect(&worker, &DVE::SnapshotRegenWorker::regenFinished, &loop,
            [&](bool o, QString e) { ok = o; err = e; loop.quit(); }, Qt::QueuedConnection);

        QMetaObject::invokeMethod(&worker, "requestRegen", Qt::QueuedConnection);
        QTimer::singleShot(30000, &loop, &QEventLoop::quit);
        loop.exec();

        QVERIFY2(ok, qPrintable("regenToPath failed: " + err));
        QVERIFY2(QFile::exists(snap), "snapshot file not produced");

        // The produced SQLite must open and expose a files table.
        {
            const QString cn = QStringLiteral("dve_regen_verify_%1")
                                   .arg(QDateTime::currentMSecsSinceEpoch());
            {
                QSqlDatabase v = QSqlDatabase::addDatabase("QSQLITE", cn);
                v.setDatabaseName(snap);
                QVERIFY(v.open());
                QSqlQuery q(v);
                QVERIFY(q.exec("SELECT count(*) FROM files"));
                QVERIFY(q.next());
                v.close();
            }
            QSqlDatabase::removeDatabase(cn);
        }

        QMetaObject::invokeMethod(&worker, "stop", Qt::BlockingQueuedConnection);
        thread.quit();
        thread.wait(3000);
        QFile::remove(snap);
    }

    // Posting a second requestRegen from within the first regenFinished handler
    // (worker now idle) runs a fresh second regen -- the worker serialises into
    // exactly two finishes (first + one follow-up), never a spurious third.
    void tst_regenWorker_pendingCoalesces() {
        DVE::DbConfig cfg = pgConfig();
        if (cfg.host.isEmpty()) QSKIP("DVE_TEST_PG_CONN not set");

        const QString snap = QDir::tempPath()
            + QStringLiteral("/dve_coalesce_%1.sqlite").arg(QDateTime::currentMSecsSinceEpoch());

        QThread thread;
        DVE::SnapshotRegenWorker worker(cfg, snap);
        worker.moveToThread(&thread);
        thread.start();
        QMetaObject::invokeMethod(&worker, "start", Qt::QueuedConnection);

        int finishes = 0;
        bool postedSecond = false;
        QEventLoop loop;
        QObject::connect(&worker, &DVE::SnapshotRegenWorker::regenFinished, &loop,
            [&](bool, QString) {
                ++finishes;
                if (!postedSecond) {
                    postedSecond = true;
                    QMetaObject::invokeMethod(&worker, "requestRegen", Qt::QueuedConnection);
                }
                if (finishes >= 2) loop.quit();
            }, Qt::QueuedConnection);

        QMetaObject::invokeMethod(&worker, "requestRegen", Qt::QueuedConnection);
        QTimer::singleShot(70000, &loop, &QEventLoop::quit);
        loop.exec();

        QCOMPARE(finishes, 2);

        QMetaObject::invokeMethod(&worker, "stop", Qt::BlockingQueuedConnection);
        thread.quit();
        thread.wait(3000);
        QFile::remove(snap);
    }
};

QTEST_MAIN(TstOfflineSnapshot)
#include "tst_offlinesnapshot.moc"

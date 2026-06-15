#include "MigrationTool.h"

#include <QFile>
#include <QSqlError>
#include <QSqlQuery>
#include <QUuid>
#include <QCryptographicHash>
#include <QDateTime>
#include <QElapsedTimer>
#include <QCoreApplication>
#include <QThread>

namespace DVE {

MigrationTool::MigrationTool(QObject* parent) : QObject(parent) {
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    m_pgConnName = "dve_mig_" + tag;
}

MigrationTool::~MigrationTool() {
    const QString srcName = "dve_mig_src_" + m_pgConnName;
    if (m_sqlite.isOpen()) {
        m_sqlite.close();
        m_sqlite = QSqlDatabase();
    }
    if (m_pg.isOpen()) {
        m_pg.close();
        m_pg = QSqlDatabase();
    }
    // Only try to remove if still registered
    if (QSqlDatabase::contains(srcName))     QSqlDatabase::removeDatabase(srcName);
    if (QSqlDatabase::contains(m_pgConnName)) QSqlDatabase::removeDatabase(m_pgConnName);
}

bool MigrationTool::open(const QString& sqlitePath, const DbConfig& pg) {
    m_sqlitePath = sqlitePath;
    m_report.setSourcePath(sqlitePath);

    if (!QFile::exists(sqlitePath)) {
        m_lastError = "Source SQLite not found at " + sqlitePath;
        return false;
    }

    const QString srcName = "dve_mig_src_" + m_pgConnName;
    m_sqlite = QSqlDatabase::addDatabase("QSQLITE", srcName);
    m_sqlite.setDatabaseName(sqlitePath);
    m_sqlite.setConnectOptions("QSQLITE_OPEN_READONLY");
    if (!m_sqlite.open()) {
        m_lastError = "SQLite open failed: " + m_sqlite.lastError().text();
        return false;
    }

    QFile f(sqlitePath);
    if (f.open(QIODevice::ReadOnly)) {
        QCryptographicHash h(QCryptographicHash::Sha256);
        if (h.addData(&f)) m_sourceSha256 = h.result().toHex();
    }

    m_pg = QSqlDatabase::addDatabase("QPSQL", m_pgConnName);
    m_pg.setHostName(pg.host);
    m_pg.setPort(pg.port);
    m_pg.setDatabaseName(pg.database);
    m_pg.setUserName(pg.user);
    m_pg.setPassword(pg.password);
    m_pg.setConnectOptions(QStringLiteral("connect_timeout=5;") + pgSharedConnectOptions());
    if (!m_pg.open()) {
        m_lastError = "Postgres open failed: " + m_pg.lastError().text();
        return false;
    }
    applyPgSessionSettings(m_pg);     // statement_timeout (best-effort)
    return true;
}

bool MigrationTool::checkSchemaMetaEmpty() {
    QSqlQuery q(m_pg);
    if (!q.exec("SELECT value FROM schema_meta WHERE key = 'migrated_at'")) {
        m_lastError = "Could not read schema_meta: " + q.lastError().text();
        return false;
    }
    return !q.next();
}

namespace {
static const QStringList kColsFiles = {
    "id", "file_path", "file_name", "loaded_at",
    "template_version", "sheet_count", "sample_count"
};
static const QStringList kColsTests = {
    "id", "file_id", "sheet_name", "template_version",
    "overall_avg_tpm", "overall_stddev_tpm", "is_raw_table", "sort_order"
};
static const QStringList kColsSamples = {
    "id", "test_id", "sort_order", "sample_name", "sample_id", "date", "tester",
    "media", "viscosity", "resistance", "voltage", "power", "heating_technology",
    "puffing_regime", "initial_oil_mass", "average_tpm", "stddev_tpm",
    "avg_power_density", "efficiency_percent", "total_oil_consumed",
    "total_puffs", "normalized_tpm", "burn_status", "clog_status", "leak_status"
};
static const QStringList kColsDataRows = {
    "id", "sample_id", "sort_order", "puffs", "before_weight", "after_weight",
    "draw_pressure", "resistance", "smell", "clog", "notes", "tpm",
    "tpm_power_density", "variation_tpm", "oil_consumed"
    // NOTE: puffing_regime is intentionally NOT here. This tool migrates a
    // legacy (pre-v2.2.1) sidecar SQLite that has no such column; the dest
    // Postgres column defaults to NULL, which is the correct "old-template"
    // value for legacy per-row data.
};
static const QStringList kColsImages = {
    "id", "sample_id", "sort_order", "file_name", "image_data",
    "layout_x", "layout_y", "layout_w", "layout_h",
    "crop_x", "crop_y", "crop_w", "crop_h"
};
static const QStringList kColsSensorySessions = {
    "id", "session_name", "tester_name", "assessor_name", "media", "puff_length",
    "date", "timestamp", "json_data", "layout_json"
};
static const QStringList kColsSensoryImages = {
    "id", "session_id", "sort_order", "file_name", "image_data",
    "layout_x", "layout_y", "layout_w", "layout_h",
    "crop_x", "crop_y", "crop_w", "crop_h"
};
static const QStringList kColsDetailedSensorySessions = {
    "id", "session_name", "tester_name", "assessor_name", "media",
    "date", "timestamp", "json_data"
};
static const QStringList kColsDetailedSensoryImages = {
    "id", "session_id", "sort_order", "file_name", "image_data",
    "layout_x", "layout_y", "layout_w", "layout_h",
    "crop_x", "crop_y", "crop_w", "crop_h"
};
static const QStringList kColsSettings = { "key", "value" };

static QStringList columnsFor(const QString& name) {
    if (name == "files")                      return kColsFiles;
    if (name == "tests")                      return kColsTests;
    if (name == "samples")                    return kColsSamples;
    if (name == "data_rows")                  return kColsDataRows;
    if (name == "images")                     return kColsImages;
    if (name == "sensory_sessions")           return kColsSensorySessions;
    if (name == "sensory_images")             return kColsSensoryImages;
    if (name == "detailed_sensory_sessions")  return kColsDetailedSensorySessions;
    if (name == "detailed_sensory_images")    return kColsDetailedSensoryImages;
    if (name == "settings")                   return kColsSettings;
    return {};
}
} // namespace

bool MigrationTool::migrateTable(const QString& name) {
    QStringList cols = columnsFor(name);
    if (cols.isEmpty()) {
        m_lastError = "Unknown table for migrateTable: " + name;
        return false;
    }

    const bool hasId = (cols.first() == "id");
    const QString overriding = hasId ? "OVERRIDING SYSTEM VALUE " : "";

    QSqlQuery src(m_sqlite);
    if (!src.exec(QString("SELECT %1 FROM %2").arg(cols.join(", "), name))) {
        m_lastError = "Source read failed for " + name + ": " + src.lastError().text();
        return false;
    }

    QStringList placeholders;
    for (int i = 0; i < cols.size(); ++i) placeholders << QString(":v%1").arg(i);
    QSqlQuery dst(m_pg);
    dst.prepare(QString("INSERT INTO %1 (%2) %3VALUES (%4)")
                .arg(name, cols.join(", "), overriding, placeholders.join(", ")));

    int n = 0;
    while (src.next()) {
        for (int i = 0; i < cols.size(); ++i) {
            dst.bindValue(QString(":v%1").arg(i), src.value(i));
        }
        if (!dst.exec()) {
            m_lastError = QString("Insert row %1 into %2 failed: %3")
                            .arg(n).arg(name, dst.lastError().text());
            return false;
        }
        ++n;
    }
    return true;
}

int MigrationTool::sqliteRowCount(const QString& table) {
    QSqlQuery q(m_sqlite);
    if (!q.exec(QString("SELECT COUNT(*) FROM %1").arg(table))) return -1;
    return q.next() ? q.value(0).toInt() : -1;
}

int MigrationTool::postgresRowCount(const QString& table) {
    QSqlQuery q(m_pg);
    if (!q.exec(QString("SELECT COUNT(*) FROM %1").arg(table))) return -1;
    return q.next() ? q.value(0).toInt() : -1;
}

bool MigrationTool::bumpSequence(const QString& table) {
    if (table == "settings") return true;  // no id sequence on settings
    QSqlQuery q(m_pg);
    if (!q.exec(QString("SELECT setval(pg_get_serial_sequence('%1','id'), "
                        "  COALESCE((SELECT MAX(id) FROM %1), 1), true)")
                .arg(table))) {
        m_lastError = "Sequence bump failed for " + table + ": "
                      + q.lastError().text();
        return false;
    }
    return true;
}

bool MigrationTool::writeSchemaMeta() {
    QSqlQuery q(m_pg);
    auto upsert = [&q](const QString& k, const QString& v) -> bool {
        q.prepare("INSERT INTO schema_meta(key, value) VALUES (?, ?) "
                  "ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value");
        q.addBindValue(k);
        q.addBindValue(v);
        return q.exec();
    };
    if (!upsert("schema_version", "2"))                                          return false;
    if (!upsert("migrated_from",  m_sqlitePath))                                 return false;
    if (!upsert("migrated_at",
                QDateTime::currentDateTimeUtc().toString(Qt::ISODate)))          return false;
    if (!upsert("source_sha256",  m_sourceSha256))                               return false;
    return true;
}

bool MigrationTool::wipePostgresData() {
    const QStringList order = {
        "data_rows", "images", "samples", "tests", "files",
        "sensory_images", "sensory_sessions",
        "detailed_sensory_images", "detailed_sensory_sessions",
        "settings", "schema_meta"
    };
    QSqlQuery q(m_pg);
    for (const QString& t : order) {
        if (!q.exec("DELETE FROM " + t)) {
            m_lastError = "Wipe " + t + " failed: " + q.lastError().text();
            return false;
        }
    }
    return true;
}

bool MigrationTool::run(bool force) {
    QElapsedTimer timer;
    timer.start();

    if (!force && !checkSchemaMetaEmpty()) {
        m_lastError = "Postgres already has migration metadata. Use --force "
                      "only after rolling back via the pre-migration SQLite.";
        m_report.setStatus("aborted");
        m_report.addError(m_lastError);
        m_report.setDuration(timer.elapsed());
        return false;
    }

    if (force) {
        if (!wipePostgresData()) {
            m_report.setStatus("aborted");
            m_report.addError(m_lastError);
            m_report.setDuration(timer.elapsed());
            return false;
        }
    }

    const QStringList order = {
        "files", "tests", "samples", "data_rows", "images",
        "sensory_sessions", "sensory_images",
        "detailed_sensory_sessions", "detailed_sensory_images",
        "settings"
    };

    if (!m_pg.transaction()) {
        m_lastError = "BEGIN failed: " + m_pg.lastError().text();
        m_report.setStatus("aborted");
        m_report.addError(m_lastError);
        m_report.setDuration(timer.elapsed());
        return false;
    }

    for (const QString& t : order) {
        const int sqliteN = sqliteRowCount(t);
        if (!migrateTable(t)) {
            m_pg.rollback();
            m_report.setStatus("rolled_back");
            m_report.addError(m_lastError);
            m_report.setDuration(timer.elapsed());
            return false;
        }
        const int pgN = postgresRowCount(t);
        m_report.addTable(t, sqliteN, pgN);
        if (sqliteN != pgN) {
            m_lastError = QString("Row count mismatch on %1: sqlite=%2 pg=%3")
                            .arg(t).arg(sqliteN).arg(pgN);
            m_pg.rollback();
            m_report.setStatus("rolled_back");
            m_report.addError(m_lastError);
            m_report.setDuration(timer.elapsed());
            return false;
        }
    }

    for (const QString& t : order) {
        if (!bumpSequence(t)) {
            m_pg.rollback();
            m_report.setStatus("rolled_back");
            m_report.addError(m_lastError);
            m_report.setDuration(timer.elapsed());
            return false;
        }
    }

    if (!writeSchemaMeta()) {
        m_pg.rollback();
        m_report.setStatus("rolled_back");
        m_report.addError(m_lastError);
        m_report.setDuration(timer.elapsed());
        return false;
    }

    if (!m_pg.commit()) {
        m_lastError = "COMMIT failed: " + m_pg.lastError().text();
        m_report.setStatus("rolled_back");
        m_report.addError(m_lastError);
        m_report.setDuration(timer.elapsed());
        return false;
    }

    if (!finalizeSource()) {
        m_report.addError("Migration committed; source rename failed: " + m_lastError);
    }

    m_report.setStatus("success");
    m_report.setDuration(timer.elapsed());
    return true;
}

bool MigrationTool::finalizeSource() {
    const QString target = m_sqlitePath + ".pre-migration.sqlite";

    // On Windows, Qt's SQLite driver keeps the file handle open until the
    // connection is formally removed via QSqlDatabase::removeDatabase, NOT
    // just closed. Drop the in-flight reference first (so removeDatabase
    // doesn't warn about it being "in use"), then remove the connection,
    // then rename.
    if (m_sqlite.isOpen()) m_sqlite.close();
    m_sqlite = QSqlDatabase();
    const QString srcName = "dve_mig_src_" + m_pgConnName;
    if (QSqlDatabase::contains(srcName)) {
        QSqlDatabase::removeDatabase(srcName);
    }

    // Brief retry handles any residual OS-level handle release lag.
    for (int attempt = 0; attempt < 5; ++attempt) {
        if (QFile::rename(m_sqlitePath, target)) {
            return true;
        }
        QThread::msleep(10 * (1 + attempt));
    }

    m_lastError = "Could not rename source to " + target;
    return false;
}

} // namespace DVE

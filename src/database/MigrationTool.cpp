#include "MigrationTool.h"

#include <QFile>
#include <QSqlError>
#include <QSqlQuery>
#include <QUuid>
#include <QCryptographicHash>

namespace DVE {

MigrationTool::MigrationTool(QObject* parent) : QObject(parent) {
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    m_pgConnName = "dve_mig_" + tag;
}

MigrationTool::~MigrationTool() {
    if (m_sqlite.isOpen()) m_sqlite.close();
    if (m_pg.isOpen())     m_pg.close();
    m_sqlite = QSqlDatabase();
    m_pg     = QSqlDatabase();
    const QString srcName = "dve_mig_src_" + m_pgConnName;
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
    m_pg.setConnectOptions("connect_timeout=5");
    if (!m_pg.open()) {
        m_lastError = "Postgres open failed: " + m_pg.lastError().text();
        return false;
    }
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

// Stubs (implemented in tasks 19-20):
bool MigrationTool::wipePostgresData()                 { return false; }
bool MigrationTool::bumpSequence(const QString&)       { return false; }
bool MigrationTool::writeSchemaMeta()                  { return false; }
bool MigrationTool::finalizeSource()                   { return false; }
bool MigrationTool::run(bool)                          { return false; }

} // namespace DVE

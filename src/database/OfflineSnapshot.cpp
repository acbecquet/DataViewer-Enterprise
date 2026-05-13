#include "OfflineSnapshot.h"

#include "PostgresConnection.h"

#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QStandardPaths>
#include <QSqlDatabase>
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlRecord>
#include <QUuid>
#include <QVariant>
#include <QDebug>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QRectF>
#include <QStringList>

namespace DVE {

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
static const char* const kCreateStatements[] = {
    R"(CREATE TABLE files (
        id               INTEGER PRIMARY KEY,
        file_path        TEXT NOT NULL,
        file_name        TEXT NOT NULL,
        loaded_at        TEXT NOT NULL,
        template_version TEXT,
        sheet_count      INTEGER DEFAULT 0,
        sample_count     INTEGER DEFAULT 0,
        updated_at       TEXT NOT NULL,
        updated_by       TEXT NOT NULL,
        version          INTEGER NOT NULL DEFAULT 1
    ))",
    R"(CREATE UNIQUE INDEX idx_files_path ON files(file_path))",

    R"(CREATE TABLE tests (
        id                 INTEGER PRIMARY KEY,
        file_id            INTEGER NOT NULL,
        sheet_name         TEXT NOT NULL,
        template_version   TEXT,
        overall_avg_tpm    REAL DEFAULT 0.0,
        overall_stddev_tpm REAL DEFAULT 0.0,
        is_raw_table       INTEGER DEFAULT 0,
        sort_order         INTEGER DEFAULT 0,
        updated_at         TEXT NOT NULL,
        updated_by         TEXT NOT NULL,
        version            INTEGER NOT NULL DEFAULT 1
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
        version       INTEGER NOT NULL DEFAULT 1
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
        version       INTEGER NOT NULL DEFAULT 1
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
                           QStringList* outPaths,
                           QVector<QRectF>* outLayouts,
                           QVector<QRectF>* outCrops);

bool deserializeSensoryJsonLocal(const QByteArray& bytes, SensorySession& sess);
bool deserializeDetailedSensoryJsonLocal(const QByteArray& bytes,
                                         DetailedSensorySession& sess);
} // anonymous

// ----------------------------------------------------------------------------
// Construction / destruction
// ----------------------------------------------------------------------------
OfflineSnapshot::OfflineSnapshot(QObject* parent) : QObject(parent) {
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    m_connName = QStringLiteral("dve_snapshot_ro_") + tag;
}

OfflineSnapshot::~OfflineSnapshot() {
    close();
}

void OfflineSnapshot::setOverrideDirForTesting(const QString& dir) {
    m_overrideDir = dir;
    m_path.clear();
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
// ----------------------------------------------------------------------------
bool OfflineSnapshot::regenerate(PostgresConnection* live) {
    m_lastError.clear();
    if (!live || !live->isOpen()) {
        m_lastError = QStringLiteral("regenerate: PostgresConnection is not open");
        return false;
    }
    if (m_open) close();

    const QString prodPath = path();
    const QString tmpPath  = prodPath + ".tmp";
    const QString tmpConn  = QStringLiteral("dve_snapshot_tmp_") +
                             QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);

    QFile::remove(tmpPath);   // wipe any stale .tmp from a previous failure

    bool success = false;
    auto cleanup = [&]() {
        if (QSqlDatabase::contains(tmpConn)) {
            QSqlDatabase::removeDatabase(tmpConn);
        }
        if (!success) QFile::remove(tmpPath);
    };

    {
        QSqlDatabase tmpDb = QSqlDatabase::addDatabase("QSQLITE", tmpConn);
        tmpDb.setDatabaseName(tmpPath);
        if (!tmpDb.open()) {
            m_lastError = QStringLiteral("regenerate: failed to open tmp SQLite: ")
                          + tmpDb.lastError().text();
            cleanup();
            return false;
        }

        // SQLite write tuning: WAL + NORMAL sync keeps regenerate fast for
        // large databases while still being crash-safe enough for an
        // application-managed atomic-rename promotion.
        {
            QSqlQuery pragma(tmpDb);
            pragma.exec("PRAGMA journal_mode=WAL");
            pragma.exec("PRAGMA synchronous=NORMAL");
        }

        for (const char* stmt : kCreateStatements) {
            QSqlQuery q(tmpDb);
            if (!q.exec(stmt)) {
                m_lastError = QStringLiteral("regenerate(CREATE): ")
                              + q.lastError().text();
                tmpDb.close();
                cleanup();
                return false;
            }
        }

        if (!tmpDb.transaction()) {
            m_lastError = QStringLiteral("regenerate(begin): ") + tmpDb.lastError().text();
            tmpDb.close();
            cleanup();
            return false;
        }

        QSqlDatabase& pg = live->queryDb();

        // ---- files ---------------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, file_path, file_name, loaded_at, template_version, "
                          "sheet_count, sample_count, updated_at, updated_by, version "
                          "FROM files ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT files): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO files (id, file_path, file_name, loaded_at, "
                        "template_version, sheet_count, sample_count, "
                        "updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 10; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT files): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- tests ---------------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, file_id, sheet_name, template_version, "
                          "overall_avg_tpm, overall_stddev_tpm, is_raw_table, "
                          "sort_order, updated_at, updated_by, version "
                          "FROM tests ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT tests): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO tests (id, file_id, sheet_name, template_version, "
                        "overall_avg_tpm, overall_stddev_tpm, is_raw_table, sort_order, "
                        "updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 11; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT tests): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
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
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
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
                for (int c = 0; c < 28; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT samples): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- data_rows -----------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, sample_id, sort_order, puffs, before_weight, "
                          "after_weight, draw_pressure, resistance, smell, clog, notes, "
                          "tpm, tpm_power_density, variation_tpm, oil_consumed, "
                          "updated_at, updated_by, version "
                          "FROM data_rows ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT data_rows): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO data_rows (id, sample_id, sort_order, puffs, "
                        "before_weight, after_weight, draw_pressure, resistance, smell, "
                        "clog, notes, tpm, tpm_power_density, variation_tpm, oil_consumed, "
                        "updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 18; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT data_rows): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- images (BYTEA -> BLOB) ----------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, sample_id, sort_order, file_name, image_data, "
                          "layout_x, layout_y, layout_w, layout_h, "
                          "crop_x, crop_y, crop_w, crop_h, "
                          "updated_at, updated_by, version "
                          "FROM images ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT images): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO images (id, sample_id, sort_order, file_name, "
                        "image_data, layout_x, layout_y, layout_w, layout_h, "
                        "crop_x, crop_y, crop_w, crop_h, "
                        "updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 16; ++c) {
                    if (c == 4) dst.bindValue(c, src.value(c).toByteArray());
                    else        dst.bindValue(c, src.value(c));
                }
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT images): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- sensory_sessions (JSONB -> TEXT) ------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, session_name, tester_name, assessor_name, media, "
                          "puff_length, date, timestamp, "
                          "json_data::text, layout_json::text, "
                          "updated_at, updated_by, version "
                          "FROM sensory_sessions ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT sensory_sessions): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO sensory_sessions (id, session_name, tester_name, "
                        "assessor_name, media, puff_length, date, timestamp, "
                        "json_data, layout_json, updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 13; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT sensory_sessions): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- sensory_images ------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, session_id, sort_order, file_name, image_data, "
                          "layout_x, layout_y, layout_w, layout_h, "
                          "crop_x, crop_y, crop_w, crop_h, "
                          "updated_at, updated_by, version "
                          "FROM sensory_images ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT sensory_images): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO sensory_images (id, session_id, sort_order, "
                        "file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                        "crop_x, crop_y, crop_w, crop_h, "
                        "updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 16; ++c) {
                    if (c == 4) dst.bindValue(c, src.value(c).toByteArray());
                    else        dst.bindValue(c, src.value(c));
                }
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT sensory_images): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- detailed_sensory_sessions -------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, session_name, tester_name, assessor_name, media, "
                          "date, timestamp, json_data::text, "
                          "updated_at, updated_by, version "
                          "FROM detailed_sensory_sessions ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT detailed_sensory_sessions): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO detailed_sensory_sessions (id, session_name, "
                        "tester_name, assessor_name, media, date, timestamp, json_data, "
                        "updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 11; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT detailed_sensory_sessions): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- detailed_sensory_images ---------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT id, session_id, sort_order, file_name, image_data, "
                          "layout_x, layout_y, layout_w, layout_h, "
                          "crop_x, crop_y, crop_w, crop_h, "
                          "updated_at, updated_by, version "
                          "FROM detailed_sensory_images ORDER BY id")) {
                m_lastError = QStringLiteral("regenerate(SELECT detailed_sensory_images): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO detailed_sensory_images (id, session_id, sort_order, "
                        "file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                        "crop_x, crop_y, crop_w, crop_h, "
                        "updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 16; ++c) {
                    if (c == 4) dst.bindValue(c, src.value(c).toByteArray());
                    else        dst.bindValue(c, src.value(c));
                }
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT detailed_sensory_images): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- settings ------------------------------------------------------
        {
            QSqlQuery src(pg);
            if (!src.exec("SELECT key, value, updated_at, updated_by, version "
                          "FROM settings ORDER BY key")) {
                m_lastError = QStringLiteral("regenerate(SELECT settings): ")
                              + src.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO settings (key, value, updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 5; ++c) dst.bindValue(c, src.value(c));
                if (!dst.exec()) {
                    m_lastError = QStringLiteral("regenerate(INSERT settings): ")
                                  + dst.lastError().text();
                    tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
                }
            }
        }

        // ---- _snapshot_meta ------------------------------------------------
        {
            const QString stamp =
                QDateTime::currentDateTimeUtc().toString(Qt::ISODateWithMs);
            QSqlQuery dst(tmpDb);
            dst.prepare("INSERT INTO _snapshot_meta (key, value) VALUES (?, ?)");
            dst.bindValue(0, "snapshot_taken_at");
            dst.bindValue(1, stamp);
            if (!dst.exec()) {
                m_lastError = QStringLiteral("regenerate(INSERT meta): ")
                              + dst.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
            dst.bindValue(0, "source_schema_version");
            dst.bindValue(1, "2");
            if (!dst.exec()) {
                m_lastError = QStringLiteral("regenerate(INSERT meta v): ")
                              + dst.lastError().text();
                tmpDb.rollback(); tmpDb.close(); cleanup(); return false;
            }
        }

        if (!tmpDb.commit()) {
            m_lastError = QStringLiteral("regenerate(commit): ") + tmpDb.lastError().text();
            tmpDb.rollback();
            tmpDb.close();
            cleanup();
            return false;
        }
        tmpDb.close();
    }
    // tmpDb goes out of scope; remove the named connection BEFORE the file-
    // system swap so Windows lets us delete/rename the .tmp file.
    QSqlDatabase::removeDatabase(tmpConn);

    // ---- Atomic-rename promotion -------------------------------------------
    // SQLite on Windows holds a file handle until removeDatabase() completes,
    // which is why the temp connection is removed above. After that, the
    // swap is: delete production (if any) then rename tmp over it.
    //
    // QFile::rename on Windows does not overwrite an existing target, so the
    // explicit remove-first is required.
    if (QFile::exists(prodPath)) {
        if (!QFile::remove(prodPath)) {
            m_lastError = QStringLiteral("regenerate: cannot delete previous snapshot ")
                          + prodPath;
            cleanup();
            return false;
        }
    }
    QFile::remove(prodPath + "-wal");
    QFile::remove(prodPath + "-shm");

    if (!QFile::rename(tmpPath, prodPath)) {
        m_lastError = QStringLiteral("regenerate: cannot rename tmp to ") + prodPath;
        cleanup();
        return false;
    }

    success = true;
    cleanup();
    return true;
}

// ----------------------------------------------------------------------------
// openReadOnly
// ----------------------------------------------------------------------------
bool OfflineSnapshot::openReadOnly() {
    m_lastError.clear();
    if (m_open) return true;

    const QString p = path();
    if (!QFile::exists(p)) {
        m_lastError = QStringLiteral("openReadOnly: snapshot file does not exist: ") + p;
        return false;
    }

    if (QSqlDatabase::contains(m_connName)) {
        QSqlDatabase::removeDatabase(m_connName);
    }

    m_db = QSqlDatabase::addDatabase("QSQLITE", m_connName);
    m_db.setDatabaseName(p);
    // Qt 6's QSQLITE driver expects flag-style options without "=1"; the
    // "QSQLITE_OPEN_READONLY=1" form is rejected as "Unsupported option".
    m_db.setConnectOptions("QSQLITE_OPEN_READONLY");
    if (!m_db.open()) {
        m_lastError = QStringLiteral("openReadOnly: ") + m_db.lastError().text();
        m_db = QSqlDatabase();
        QSqlDatabase::removeDatabase(m_connName);
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

    m_open = true;
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
    // We wrote the stamp via QDateTime::currentDateTimeUtc().toString(),
    // so the ISO string already encodes the UTC offset; Qt parses the spec
    // from the string. No explicit setTimeSpec() needed -- and Qt 6.10's
    // setTimeSpec(Qt::UTC) is deprecated in favour of setTimeZone().
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
    if (!q.exec("SELECT id, file_path, file_name, loaded_at, template_version, "
                "sheet_count, sample_count FROM files ORDER BY loaded_at DESC")) {
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
    q.prepare("SELECT id FROM files WHERE file_path = ? LIMIT 1");
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
    };
    QVector<TestInfo> tests;
    {
        QSqlQuery q(m_db);
        q.prepare("SELECT id, sheet_name, template_version, overall_avg_tpm, "
                  "overall_stddev_tpm, is_raw_table FROM tests "
                  "WHERE file_id = ? ORDER BY sort_order");
        q.addBindValue(id);
        if (!q.exec()) {
            m_lastError = QStringLiteral("loadFile(SELECT tests): ") + q.lastError().text();
            return result;
        }
        while (q.next()) {
            tests.append({q.value(0).toInt(), q.value(1).toString(),
                          q.value(2).toString(), q.value(3).toDouble(),
                          q.value(4).toDouble(), q.value(5).toInt() != 0});
        }
    }

    for (const TestInfo& ti : tests) {
        SheetResult sheet;
        sheet.sheetName        = ti.sheetName;
        sheet.templateVersion  = ti.templateVersion;
        sheet.overallAvgTPM    = ti.avgTPM;
        sheet.overallStdDevTPM = ti.stddevTPM;
        sheet.isRawTable       = ti.isRaw;
        result.sheetNames.append(ti.sheetName);

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
            if (q.exec()) {
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
            } else {
                m_lastError = QStringLiteral("loadFile(SELECT samples): ") + q.lastError().text();
            }
        }

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
                          "variation_tpm, oil_consumed "
                          "FROM data_rows WHERE sample_id = ? ORDER BY sort_order");
                q.addBindValue(si.id);
                if (q.exec()) {
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
                        sr.rows.append(dr);
                    }
                } else {
                    m_lastError = QStringLiteral("loadFile(SELECT data_rows): ") + q.lastError().text();
                }
            }

            // images -- materialise BLOBs to disk so imagePaths works
            {
                QSqlDatabase nonConstDb = m_db;
                loadImagesForSnapshot(nonConstDb, "images", si.id, "sample_id",
                                      &sr.imagePaths, &sr.imageLayouts, &sr.imageCrops);
            }

            sheet.samples.append(sr);
        }

        for (const SampleResult& sr : sheet.samples) {
            for (const DataRow& dr : sr.rows) {
                if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
                sheet.tpmTrend.append(dr.tpm);
                sheet.puffCounts.append(dr.puffs);
            }
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
    if (!q.exec("SELECT id, session_name, assessor_name, media, date, json_data "
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
    loadImagesForSnapshot(nonConstDb, "sensory_images", id, "session_id",
                          &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
    return sess;
}

QVector<DetailedSensoryRecord> OfflineSnapshot::listDetailedSensoryRecords() const {
    QVector<DetailedSensoryRecord> result;
    if (!m_open) {
        m_lastError = QStringLiteral("listDetailedSensoryRecords: snapshot not open");
        return result;
    }
    QSqlQuery q(m_db);
    if (!q.exec("SELECT id, session_name, assessor_name, media, date, json_data "
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
    loadImagesForSnapshot(nonConstDb, "detailed_sensory_images", id, "session_id",
                          &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
    return sess;
}

// ----------------------------------------------------------------------------
// File-static helpers
// ----------------------------------------------------------------------------
namespace {

void loadImagesForSnapshot(QSqlDatabase& db,
                           const QString& tableName,
                           qint64 parentId,
                           const QString& parentCol,
                           QStringList* outPaths,
                           QVector<QRectF>* outLayouts,
                           QVector<QRectF>* outCrops)
{
    QSqlQuery qi(db);
    if (!qi.prepare(QString("SELECT file_name, image_data, layout_x, layout_y, "
                            "layout_w, layout_h, crop_x, crop_y, crop_w, crop_h "
                            "FROM %1 WHERE %2 = ? ORDER BY sort_order")
                        .arg(tableName, parentCol))) {
        return;
    }
    qi.addBindValue(static_cast<qlonglong>(parentId));
    if (!qi.exec()) return;

    const QString tempDir =
        QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation)
        + "/ImageCache";
    QDir().mkpath(tempDir);

    int ii = 0;
    while (qi.next()) {
        const QString  fileName = qi.value(0).toString();
        const QByteArray blob   = qi.value(1).toByteArray();
        const QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                            qi.value(4).toDouble(), qi.value(5).toDouble());
        const QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                          qi.value(8).toDouble(), qi.value(9).toDouble());

        const QString tempPath = tempDir + "/snap_" + tableName + "_" +
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

// Mirrors DatabaseManager's deserializeSensoryJson (kept local to keep the
// snapshot self-contained -- the DatabaseManager copy is in an anonymous
// namespace inside DatabaseManager.cpp and isn't reachable from here).
bool deserializeSensoryJsonLocal(const QByteArray& bytes, SensorySession& sess)
{
    const QJsonDocument doc = QJsonDocument::fromJson(bytes);
    if (doc.isNull() || !doc.isObject()) return false;

    const QJsonObject root = doc.object();
    sess.sessionName        = root["session_name"].toString();
    sess.testTitle          = root["test_title"].toString();
    sess.assessorName       = root["assessor_name"].toString();
    sess.testerName         = root["tester_name"].toString();
    sess.media              = root["media"].toString();
    sess.date               = root["date"].toString();
    sess.timestamp          = root["timestamp"].toString();
    sess.control            = root["control"].toString();
    sess.isBlind            = root["is_blind"].toBool(false);
    sess.primaryDifferences = root["primary_differences"].toString();
    sess.puffLength         = root["puff_length"].toString();
    sess.burnStatus         = root["burn_status"].toString();
    sess.clogStatus         = root["clog_status"].toString();
    sess.leakStatus         = root["leak_status"].toString();
    sess.resistance         = root["resistance"].toDouble();
    sess.voltage            = root["voltage"].toDouble();
    sess.power              = root["power"].toDouble();
    sess.heatingTechnology  = root["heating_technology"].toString();

    for (const QJsonValue& sv : root["samples"].toArray()) {
        const QJsonObject sObj = sv.toObject();
        SensorySample sample;
        sample.name              = sObj["name"].toString();
        sample.comments          = sObj["comments"].toString();
        sample.voltage           = sObj["voltage"].toDouble();
        sample.resistance        = sObj["resistance"].toDouble();
        sample.power             = sObj["power"].toDouble();
        sample.heatingTechnology = sObj["heating_technology"].toString();
        for (const QString& metric : kSensoryMetrics) {
            sample.scores[metric] = qBound(1.0, sObj[metric].toDouble(5.0), 9.0);
        }
        sess.samples.append(sample);
    }
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

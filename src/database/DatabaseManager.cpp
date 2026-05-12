#include "DatabaseManager.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "ConfigLoader.h"

#include <QDebug>
#include <QSqlQuery>
#include <QSqlError>
#include <QSqlRecord>
#include <QSqlDatabase>
#include <QVariant>
#include <QUuid>
#include <QDateTime>
#include <QFile>
#include <QFileInfo>
#include <QDir>
#include <QStandardPaths>
#include <QRectF>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>
#include <QRegularExpression>

namespace DVE {

// --- ctor / dtor / open / close ---------------------------------------------
DatabaseManager::DatabaseManager(QObject* parent)
    : QObject(parent), m_pg(new PostgresConnection(this)) {}

DatabaseManager::~DatabaseManager() { close(); }

bool DatabaseManager::open(const DbConfig& cfg, IdentityManager* identity) {
    m_identity = identity;
    if (!m_pg->open(cfg)) {
        m_lastError = m_pg->lastError();
        m_open = false;
        return false;
    }
    m_lastError.clear();
    m_open = true;
    return true;
}

void DatabaseManager::close() {
    if (m_pg) m_pg->close();
    m_open = false;
}

bool DatabaseManager::isOpen() const {
    return m_open && m_pg && m_pg->isOpen();
}

QString DatabaseManager::currentPath() const { return QString(); }

void DatabaseManager::logDebug(const QString& msg) const {
    qDebug().noquote() << "[DatabaseManager]" << msg;
}

// --- helper: who is making this change? -------------------------------------
// Postgres requires a non-null updated_by on every write. IdentityManager
// always supplies one once open() has been called with a real identity. If
// m_identity is somehow null (shouldn't happen post-3a) we fall back to a
// generic marker rather than crash.
static QString writerUuid(IdentityManager* id) {
    if (id) return id->uuid().toString(QUuid::WithoutBraces);
    return QStringLiteral("unknown");
}

// ============================================================================
//  Hierarchical file storage
// ============================================================================
bool DatabaseManager::saveFile(const FileResult& result) {
    if (!isOpen()) {
        m_lastError = QStringLiteral("database not open");
        return false;
    }

    int totalSamples = 0;
    for (const auto& sheet : result.sheets)
        totalSamples += sheet.samples.size();

    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("could not begin transaction: ")
                      + db.lastError().text();
        logDebug("saveFile " + m_lastError);
        return false;
    }

    // -- Upsert the file row by file_path and capture the resulting id ------
    // ON CONFLICT (file_path) DO UPDATE overwrites metadata; RETURNING id
    // chains children regardless of insert-vs-update.
    int fileId = -1;
    {
        QSqlQuery q(db);
        q.prepare(
            "INSERT INTO files (file_path, file_name, loaded_at, template_version, "
            "sheet_count, sample_count, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT (file_path) DO UPDATE SET "
            "  file_name = EXCLUDED.file_name, "
            "  loaded_at = EXCLUDED.loaded_at, "
            "  template_version = EXCLUDED.template_version, "
            "  sheet_count = EXCLUDED.sheet_count, "
            "  sample_count = EXCLUDED.sample_count, "
            "  updated_by = EXCLUDED.updated_by "
            "RETURNING id");
        q.addBindValue(result.filePath);
        q.addBindValue(result.fileName);
        q.addBindValue(QDateTime::currentDateTime().toString(Qt::ISODate));
        q.addBindValue(result.templateVersion);
        q.addBindValue(result.sheets.size());
        q.addBindValue(totalSamples);
        q.addBindValue(who);
        if (!q.exec() || !q.next()) {
            m_lastError = q.lastError().text();
            db.rollback();
            logDebug("saveFile UPSERT files failed: " + m_lastError);
            return false;
        }
        fileId = q.value(0).toInt();
    }

    // -- Wipe all children of this file and rebuild them. CASCADE on tests
    //    handles samples / data_rows / images automatically.
    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM tests WHERE file_id = ?");
        del.addBindValue(fileId);
        if (!del.exec()) {
            m_lastError = del.lastError().text();
            db.rollback();
            logDebug("saveFile DELETE tests failed: " + m_lastError);
            return false;
        }
    }

    // -- Insert tests -> samples -> data_rows + images ----------------------
    // Prepare the four insert statements once outside the loops, then rebind
    // per iteration. The old SQLite saveFile did this; the initial Postgres
    // port regressed it by constructing fresh QSqlQuery objects inside the
    // deepest loop, re-parsing identical SQL on every row.
    QSqlQuery insertTest(db);
    if (!insertTest.prepare(
            "INSERT INTO tests (file_id, sheet_name, template_version, "
            "overall_avg_tpm, overall_stddev_tpm, is_raw_table, sort_order, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?) RETURNING id")) {
        m_lastError = insertTest.lastError().text();
        db.rollback();
        logDebug("saveFile prepare INSERT test failed: " + m_lastError);
        return false;
    }

    QSqlQuery insertSample(db);
    if (!insertSample.prepare(
            "INSERT INTO samples (test_id, sort_order, sample_name, sample_id, date, tester, "
            "media, viscosity, resistance, voltage, power, heating_technology, puffing_regime, "
            "initial_oil_mass, average_tpm, stddev_tpm, avg_power_density, efficiency_percent, "
            "total_oil_consumed, total_puffs, normalized_tpm, burn_status, clog_status, leak_status, "
            "updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "RETURNING id")) {
        m_lastError = insertSample.lastError().text();
        db.rollback();
        logDebug("saveFile prepare INSERT sample failed: " + m_lastError);
        return false;
    }

    QSqlQuery insertRow(db);
    if (!insertRow.prepare(
            "INSERT INTO data_rows (sample_id, sort_order, puffs, before_weight, after_weight, "
            "draw_pressure, resistance, smell, clog, notes, tpm, tpm_power_density, "
            "variation_tpm, oil_consumed, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")) {
        m_lastError = insertRow.lastError().text();
        db.rollback();
        logDebug("saveFile prepare INSERT data_row failed: " + m_lastError);
        return false;
    }

    QSqlQuery insertImage(db);
    if (!insertImage.prepare(
            "INSERT INTO images (sample_id, sort_order, file_name, image_data, "
            "layout_x, layout_y, layout_w, layout_h, crop_x, crop_y, crop_w, crop_h, "
            "updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")) {
        m_lastError = insertImage.lastError().text();
        db.rollback();
        logDebug("saveFile prepare INSERT image failed: " + m_lastError);
        return false;
    }

    for (int si = 0; si < result.sheets.size(); ++si) {
        const SheetResult& sheet = result.sheets[si];

        int testId = -1;
        {
            insertTest.bindValue(0, fileId);
            insertTest.bindValue(1, sheet.sheetName);
            insertTest.bindValue(2, sheet.templateVersion);
            insertTest.bindValue(3, sheet.overallAvgTPM);
            insertTest.bindValue(4, sheet.overallStdDevTPM);
            insertTest.bindValue(5, sheet.isRawTable ? 1 : 0);
            insertTest.bindValue(6, si);
            insertTest.bindValue(7, who);
            if (!insertTest.exec() || !insertTest.next()) {
                m_lastError = insertTest.lastError().text();
                db.rollback();
                logDebug("saveFile INSERT test failed: " + m_lastError);
                return false;
            }
            testId = insertTest.value(0).toInt();
        }

        for (int sj = 0; sj < sheet.samples.size(); ++sj) {
            const SampleResult& sr = sheet.samples[sj];

            int sampleId = -1;
            {
                insertSample.bindValue(0, testId);
                insertSample.bindValue(1, sj);
                insertSample.bindValue(2, sr.sampleName);
                insertSample.bindValue(3, sr.sampleID);
                insertSample.bindValue(4, sr.date);
                insertSample.bindValue(5, sr.tester);
                insertSample.bindValue(6, sr.media);
                insertSample.bindValue(7, sr.viscosity);
                insertSample.bindValue(8, sr.resistance);
                insertSample.bindValue(9, sr.voltage);
                insertSample.bindValue(10, sr.power);
                insertSample.bindValue(11, sr.heatingTechnology);
                insertSample.bindValue(12, sr.puffingRegime);
                insertSample.bindValue(13, sr.initialOilMass);
                insertSample.bindValue(14, sr.averageTPM);
                insertSample.bindValue(15, sr.stdDevTPM);
                insertSample.bindValue(16, sr.averagePowerDensity);
                insertSample.bindValue(17, sr.efficiencyPercent);
                insertSample.bindValue(18, sr.totalOilConsumed);
                insertSample.bindValue(19, sr.totalPuffs);
                insertSample.bindValue(20, sr.normalizedTPM);
                insertSample.bindValue(21, sr.burnStatus);
                insertSample.bindValue(22, sr.clogStatus);
                insertSample.bindValue(23, sr.leakStatus);
                insertSample.bindValue(24, who);
                if (!insertSample.exec() || !insertSample.next()) {
                    m_lastError = insertSample.lastError().text();
                    db.rollback();
                    logDebug("saveFile INSERT sample failed: " + m_lastError);
                    return false;
                }
                sampleId = insertSample.value(0).toInt();
            }

            // -- data rows ------------------------------------------------
            for (int ri = 0; ri < sr.rows.size(); ++ri) {
                const DataRow& dr = sr.rows[ri];
                insertRow.bindValue(0, sampleId);
                insertRow.bindValue(1, ri);
                insertRow.bindValue(2, dr.puffs);
                insertRow.bindValue(3, dr.beforeWeight);
                insertRow.bindValue(4, dr.afterWeight);
                insertRow.bindValue(5, dr.drawPressure);
                insertRow.bindValue(6, dr.resistance);
                insertRow.bindValue(7, dr.smell);
                insertRow.bindValue(8, dr.clog);
                insertRow.bindValue(9, dr.notes);
                insertRow.bindValue(10, dr.tpm);
                insertRow.bindValue(11, dr.tpmPowerDensity);
                insertRow.bindValue(12, dr.variationTPM);
                insertRow.bindValue(13, dr.oilConsumed);
                insertRow.bindValue(14, who);
                if (!insertRow.exec()) {
                    m_lastError = insertRow.lastError().text();
                    db.rollback();
                    logDebug("saveFile INSERT data_row failed: " + m_lastError);
                    return false;
                }
            }

            // -- images (per-sample) --------------------------------------
            // Reads each on-disk image into a BYTEA blob and stores the
            // layout/crop rectangles. Skips files we can't open or that
            // exceed 100 MB - matching the old behaviour.
            for (int ii = 0; ii < sr.imagePaths.size(); ++ii) {
                QByteArray imgData;
                QFile imgFile(sr.imagePaths[ii]);
                if (imgFile.open(QIODevice::ReadOnly)) {
                    constexpr qint64 kMaxImageSize = 100 * 1024 * 1024;
                    if (imgFile.size() <= kMaxImageSize)
                        imgData = imgFile.readAll();
                    else
                        qWarning() << "Skipping oversized image:" << imgFile.fileName();
                }

                const QRectF layout = (ii < sr.imageLayouts.size()) ? sr.imageLayouts[ii] : QRectF();
                const QRectF crop   = (ii < sr.imageCrops.size())   ? sr.imageCrops[ii]   : QRectF(0,0,1,1);

                insertImage.bindValue(0, sampleId);
                insertImage.bindValue(1, ii);
                insertImage.bindValue(2, QFileInfo(sr.imagePaths[ii]).fileName());
                insertImage.bindValue(3, imgData);
                insertImage.bindValue(4, layout.x());
                insertImage.bindValue(5, layout.y());
                insertImage.bindValue(6, layout.width());
                insertImage.bindValue(7, layout.height());
                insertImage.bindValue(8, crop.x());
                insertImage.bindValue(9, crop.y());
                insertImage.bindValue(10, crop.width());
                insertImage.bindValue(11, crop.height());
                insertImage.bindValue(12, who);
                if (!insertImage.exec()) {
                    m_lastError = insertImage.lastError().text();
                    db.rollback();
                    logDebug("saveFile INSERT image failed: " + m_lastError);
                    return false;
                }
            }
        }
    }

    if (!db.commit()) {
        m_lastError = db.lastError().text();
        db.rollback();
        logDebug("saveFile commit failed: " + m_lastError);
        return false;
    }
    logDebug(QString("Saved file '%1' (fileId=%2, %3 sheets, %4 samples)")
                 .arg(result.fileName).arg(fileId)
                 .arg(result.sheets.size()).arg(totalSamples));
    return true;
}

// --- hasFile ----------------------------------------------------------------
bool DatabaseManager::hasFile(const QString& filePath) const {
    if (!isOpen()) return false;
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT 1 FROM files WHERE file_path = ? LIMIT 1");
    q.addBindValue(filePath);
    return q.exec() && q.next();
}

// --- loadFile ---------------------------------------------------------------
// Pure read path - no transaction. Walks files -> tests -> samples ->
// data_rows + images. Same hierarchical traversal as the old SQLite version:
// we collect child ids into vectors before issuing the next-level query to
// avoid nested-cursor surprises on some drivers.
FileResult DatabaseManager::loadFile(int id) const {
    FileResult result;
    if (!isOpen()) return result;

    QSqlDatabase& db = m_pg->queryDb();

    // Step 1: file metadata
    {
        QSqlQuery q(db);
        q.prepare("SELECT file_path, file_name, template_version FROM files WHERE id = ?");
        q.addBindValue(id);
        if (!q.exec() || !q.next()) return result;
        result.filePath        = q.value(0).toString();
        result.fileName        = q.value(1).toString();
        result.templateVersion = q.value(2).toString();
    }

    // Step 2: tests
    struct TestInfo {
        int id; QString sheetName; QString templateVersion;
        double avgTPM; double stddevTPM; bool isRaw;
    };
    QVector<TestInfo> tests;
    {
        QSqlQuery q(db);
        q.prepare("SELECT id, sheet_name, template_version, overall_avg_tpm, "
                  "overall_stddev_tpm, is_raw_table FROM tests "
                  "WHERE file_id = ? ORDER BY sort_order");
        q.addBindValue(id);
        if (!q.exec()) return result;
        while (q.next()) {
            tests.append({q.value(0).toInt(), q.value(1).toString(),
                          q.value(2).toString(), q.value(3).toDouble(),
                          q.value(4).toDouble(), q.value(5).toInt() != 0});
        }
    }

    // Step 3: per-test, samples + their children
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
            QSqlQuery q(db);
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
                QSqlQuery q(db);
                q.prepare("SELECT puffs, before_weight, after_weight, draw_pressure, resistance, "
                          "smell, clog, notes, tpm, tpm_power_density, variation_tpm, oil_consumed "
                          "FROM data_rows WHERE sample_id = ? ORDER BY sort_order");
                q.addBindValue(si.id);
                if (q.exec()) {
                    while (q.next()) {
                        DataRow dr;
                        dr.puffs           = q.value(0).toDouble();
                        dr.beforeWeight    = q.value(1).toDouble();
                        dr.afterWeight     = q.value(2).toDouble();
                        dr.drawPressure    = q.value(3).toDouble();
                        dr.resistance      = q.value(4).toDouble();
                        dr.smell           = q.value(5).toString();
                        dr.clog            = q.value(6).toString();
                        dr.notes           = q.value(7).toString();
                        dr.tpm             = q.value(8).toDouble();
                        dr.tpmPowerDensity = q.value(9).toDouble();
                        dr.variationTPM    = q.value(10).toDouble();
                        dr.oilConsumed     = q.value(11).toDouble();
                        sr.rows.append(dr);
                    }
                }
            }

            // images - materialise BLOBs to disk so imagePaths works
            {
                QSqlQuery q(db);
                q.prepare("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                          "crop_x, crop_y, crop_w, crop_h "
                          "FROM images WHERE sample_id = ? ORDER BY sort_order");
                q.addBindValue(si.id);
                if (q.exec()) {
                    const QString tempDir =
                        QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation)
                        + "/ImageCache";
                    QDir().mkpath(tempDir);
                    while (q.next()) {
                        const QString fileName = q.value(0).toString();
                        const QByteArray blob  = q.value(1).toByteArray();
                        const QRectF layout(q.value(2).toDouble(), q.value(3).toDouble(),
                                            q.value(4).toDouble(), q.value(5).toDouble());
                        const QRectF crop(q.value(6).toDouble(), q.value(7).toDouble(),
                                          q.value(8).toDouble(), q.value(9).toDouble());

                        const QString tempPath = tempDir + "/" + QString::number(si.id) + "_" + fileName;
                        QFile tmpFile(tempPath);
                        if (tmpFile.open(QIODevice::WriteOnly)) {
                            tmpFile.write(blob);
                            tmpFile.close();
                        }
                        sr.imagePaths.append(tempPath);
                        sr.imageLayouts.append(layout);
                        sr.imageCrops.append(crop);
                    }
                }
            }

            sheet.samples.append(sr);
        }

        // Rebuild plot series from row data (not stored separately).
        for (const SampleResult& sr : sheet.samples) {
            for (const DataRow& dr : sr.rows) {
                if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
                sheet.tpmTrend.append(dr.tpm);
                sheet.puffCounts.append(dr.puffs);
            }
        }

        result.sheets.append(sheet);
    }

    logDebug(QString("Loaded file id=%1 '%2' (%3 sheets)")
                 .arg(id).arg(result.fileName).arg(result.sheets.size()));
    return result;
}

FileResult DatabaseManager::loadFileByPath(const QString& filePath) const {
    FileResult result;
    if (!isOpen()) return result;
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT id FROM files WHERE file_path = ? LIMIT 1");
    q.addBindValue(filePath);
    if (q.exec() && q.next())
        return loadFile(q.value(0).toInt());
    return result;
}

// --- listFiles --------------------------------------------------------------
QVector<FileRecord> DatabaseManager::listFiles() const {
    QVector<FileRecord> records;
    if (!isOpen()) return records;

    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT id, file_path, file_name, loaded_at, template_version, "
              "sheet_count, sample_count FROM files ORDER BY loaded_at DESC");
    if (!q.exec()) return records;
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

// --- removeFile -------------------------------------------------------------
bool DatabaseManager::removeFile(int id) {
    if (!isOpen()) return false;
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM files WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }
    return true;
}

// --- deduplicateFiles -------------------------------------------------------
// 1. Delete every row whose template_version is the literal string "unknown"
//    (these come from earlier broken loads).
// 2. For each distinct file_name, keep the N most recently loaded rows and
//    delete the rest. CASCADE handles all the children.
int DatabaseManager::deduplicateFiles(int keepPerName) {
    if (!isOpen()) return 0;

    QSqlDatabase& db = m_pg->queryDb();
    int deleted = 0;

    // 1. unknown-template wipeout
    {
        QSqlQuery q(db);
        q.prepare("SELECT id FROM files WHERE template_version = 'unknown'");
        if (q.exec()) {
            QVector<int> ids;
            while (q.next()) ids.append(q.value(0).toInt());
            for (int id : ids) {
                if (removeFile(id)) ++deleted;
            }
        }
    }

    // 2. per-file_name retention
    QStringList names;
    {
        QSqlQuery q(db);
        if (q.exec("SELECT DISTINCT file_name FROM files")) {
            while (q.next()) names.append(q.value(0).toString());
        }
    }
    for (const QString& name : names) {
        QSqlQuery q(db);
        q.prepare("SELECT id FROM files WHERE file_name = ? ORDER BY loaded_at DESC");
        q.addBindValue(name);
        if (!q.exec()) continue;

        QVector<int> ids;
        while (q.next()) ids.append(q.value(0).toInt());

        for (int i = keepPerName; i < ids.size(); ++i) {
            if (removeFile(ids[i])) ++deleted;
        }
    }

    logDebug(QString("Deduplicated files: removed %1 entries").arg(deleted));
    return deleted;
}

// --- recentFilePaths --------------------------------------------------------
QStringList DatabaseManager::recentFilePaths() const {
    QStringList paths;
    if (!isOpen()) return paths;
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT file_path FROM files ORDER BY loaded_at DESC LIMIT 20");
    if (!q.exec()) return paths;
    while (q.next()) paths << q.value(0).toString();
    return paths;
}

// ============================================================================
//  Sensory sessions
// ============================================================================
//
// JSON-serialization contract: all SensorySession fields are packed into a
// single JSONB blob (json_data). A subset (session_name, tester_name, date,
// assessor_name, media, puff_length, timestamp) also goes into dedicated
// columns to support the natural-key UNIQUE index and SELECT-without-parse on
// the listing path. layout_json lives in its own column so the report-preview
// preserve it independently of saveSensorySession (saveSensoryLayout below
// UPDATEs only layout_json).
//
// The upsert key is (session_name, tester_name, date). ON CONFLICT ... DO
// UPDATE ... RETURNING id covers both insert and update branches in one
// round-trip, which is what the by-ref saveSensorySession(SensorySession&)
// overload needs to populate s.id.

namespace {

// Pack a SensorySession into the JSON blob the old SQLite code wrote. Kept
// byte-identical (in field set and key names) so pre-existing data loads.
QString serializeSensoryJson(const SensorySession& s)
{
    QJsonObject root;
    root["session_name"]         = s.sessionName;
    root["test_title"]           = s.testTitle;
    root["assessor_name"]        = s.assessorName;
    root["tester_name"]          = s.testerName;
    root["media"]                = s.media;
    root["date"]                 = s.date;
    root["timestamp"]            = s.timestamp;
    root["control"]              = s.control;
    root["is_blind"]             = s.isBlind;
    root["primary_differences"]  = s.primaryDifferences;
    root["puff_length"]          = s.puffLength;
    root["burn_status"]          = s.burnStatus;
    root["clog_status"]          = s.clogStatus;
    root["leak_status"]          = s.leakStatus;
    root["resistance"]           = s.resistance;
    root["voltage"]              = s.voltage;
    root["power"]                = s.power;
    root["heating_technology"]   = s.heatingTechnology;

    QJsonArray samplesArr;
    for (const SensorySample& sample : s.samples) {
        QJsonObject sObj;
        sObj["name"]     = sample.name;
        sObj["comments"] = sample.comments;
        for (const QString& metric : kSensoryMetrics) {
            sObj[metric] = sample.scores.value(metric, 5.0);
        }
        sObj["voltage"]            = sample.voltage;
        sObj["resistance"]         = sample.resistance;
        sObj["power"]              = sample.power;
        sObj["heating_technology"] = sample.heatingTechnology;
        samplesArr.append(sObj);
    }
    root["samples"] = samplesArr;

    return QString::fromUtf8(QJsonDocument(root).toJson(QJsonDocument::Compact));
}

// Inverse of serializeSensoryJson; fills the non-id, non-image fields. Returns
// false on malformed/empty JSON so the caller can short-circuit.
bool deserializeSensoryJson(const QByteArray& bytes, SensorySession& sess)
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

QString serializeDetailedSensoryJson(const DetailedSensorySession& s)
{
    QJsonObject root;
    root["session_name"]        = s.sessionName;
    root["test_title"]          = s.testTitle;
    root["assessor_name"]       = s.assessorName;
    root["tester_name"]         = s.testerName;
    root["facilitator_name"]    = s.facilitatorName;
    root["facilitator_comment"] = s.facilitatorComment;
    root["media"]               = s.media;
    root["date"]                = s.date;
    root["timestamp"]           = s.timestamp;
    root["oil_smell_liking"]    = s.oilSmellLiking;
    root["clog"]                = s.clog;
    root["clog_oil_level"]      = s.clogOilLevel;
    root["mouthpiece_notes"]    = s.mouthpieceNotes;
    root["device_return_date"]  = s.deviceReturnDate;
    root["viscosity"]           = s.viscosity;

    QJsonArray samplesArr;
    for (const DetailedSensorySample& sample : s.samples) {
        QJsonObject sObj;
        sObj["name"]     = sample.name;
        sObj["comments"] = sample.comments;
        for (const QString& metric : kDetailedAllMetrics) {
            sObj[metric] = sample.scores.value(metric, 0.0);
        }
        sObj["voltage"]            = sample.voltage;
        sObj["resistance"]         = sample.resistance;
        sObj["power"]              = sample.power;
        sObj["heating_technology"] = sample.heatingTechnology;
        samplesArr.append(sObj);
    }
    root["samples"] = samplesArr;

    return QString::fromUtf8(QJsonDocument(root).toJson(QJsonDocument::Compact));
}

bool deserializeDetailedSensoryJson(const QByteArray& bytes, DetailedSensorySession& sess)
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

// Walk imagePaths/imageLayouts/imageCrops and INSERT a row per image into the
// supplied images table. Uses a single prepared statement per call. Layout
// missing → default-constructed QRectF, crop missing → (0,0,1,1) per the old
// SQLite behaviour.
bool insertImagesFor(QSqlDatabase& db,
                     const QString& tableName,
                     qint64 sessionId,
                     const QStringList& imagePaths,
                     const QVector<QRectF>& imageLayouts,
                     const QVector<QRectF>& imageCrops,
                     const QString& who,
                     QString* outError)
{
    if (imagePaths.isEmpty()) return true;

    QSqlQuery imgQ(db);
    if (!imgQ.prepare(QString("INSERT INTO %1 "
                              "(session_id, sort_order, file_name, image_data,"
                              " layout_x, layout_y, layout_w, layout_h,"
                              " crop_x, crop_y, crop_w, crop_h, updated_by) "
                              "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").arg(tableName))) {
        if (outError) *outError = imgQ.lastError().text();
        return false;
    }

    for (int i = 0; i < imagePaths.size(); ++i) {
        QByteArray imgData;
        QFile imgFile(imagePaths[i]);
        if (imgFile.open(QIODevice::ReadOnly)) {
            constexpr qint64 kMaxImageSize = 100 * 1024 * 1024;
            if (imgFile.size() <= kMaxImageSize)
                imgData = imgFile.readAll();
            else
                qWarning() << "Skipping oversized image:" << imgFile.fileName();
        }

        const QRectF layout = (i < imageLayouts.size()) ? imageLayouts[i] : QRectF();
        const QRectF crop   = (i < imageCrops.size())   ? imageCrops[i]   : QRectF(0,0,1,1);

        imgQ.bindValue(0,  static_cast<qlonglong>(sessionId));
        imgQ.bindValue(1,  i);
        imgQ.bindValue(2,  QFileInfo(imagePaths[i]).fileName());
        imgQ.bindValue(3,  imgData);
        imgQ.bindValue(4,  layout.x());
        imgQ.bindValue(5,  layout.y());
        imgQ.bindValue(6,  layout.width());
        imgQ.bindValue(7,  layout.height());
        imgQ.bindValue(8,  crop.x());
        imgQ.bindValue(9,  crop.y());
        imgQ.bindValue(10, crop.width());
        imgQ.bindValue(11, crop.height());
        imgQ.bindValue(12, who);
        if (!imgQ.exec()) {
            if (outError) *outError = imgQ.lastError().text();
            return false;
        }
    }
    return true;
}

// Read all rows from one of the *_images tables into the supplied path/layout/
// crop vectors. BYTEA blobs are materialised under AppLocalDataLocation/
// ImageCache/ so callers can treat them as on-disk files (mirrors the
// file-hierarchy loadFile pattern from 3a/3b).
void loadImagesFor(QSqlDatabase& db,
                   const QString& tableName,
                   const QString& cachePrefix,
                   qint64 sessionId,
                   QStringList* outPaths,
                   QVector<QRectF>* outLayouts,
                   QVector<QRectF>* outCrops)
{
    QSqlQuery qi(db);
    if (!qi.prepare(QString("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                            "crop_x, crop_y, crop_w, crop_h "
                            "FROM %1 WHERE session_id = ? ORDER BY sort_order").arg(tableName))) {
        return;
    }
    qi.addBindValue(static_cast<qlonglong>(sessionId));
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

        const QString tempPath = tempDir + "/" + cachePrefix + "_" +
                                 QString::number(sessionId) + "_" +
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

} // namespace

bool DatabaseManager::saveSensorySession(const SensorySession& s)
{
    if (!isOpen()) {
        m_lastError = QStringLiteral("database not open");
        return false;
    }

    const QString jsonStr = serializeSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("could not begin transaction: ")
                      + db.lastError().text();
        logDebug("saveSensorySession " + m_lastError);
        return false;
    }

    logDebug(QString("saveSensorySession: name='%1' tester='%2' date='%3' samples=%4")
                 .arg(s.sessionName, s.testerName, s.date)
                 .arg(s.samples.size()));

    // Upsert by natural key. layout_json is preserved by COALESCE(EXCLUDED,
    // existing) — EXCLUDED is the NULL we bind here, so the existing column
    // value survives. Inserts get NULL, which is what we want for a brand-new
    // row. This is the Postgres-side equivalent of the old SQLite read-then-
    // re-bind dance.
    qint64 sessionId = -1;
    {
        QSqlQuery q(db);
        q.prepare(R"(
            INSERT INTO sensory_sessions
                (session_name, tester_name, assessor_name, media, puff_length,
                 date, timestamp, json_data, layout_json, updated_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, CAST(? AS JSONB), NULL, ?)
            ON CONFLICT (session_name, tester_name, date) DO UPDATE SET
                assessor_name = EXCLUDED.assessor_name,
                media         = EXCLUDED.media,
                puff_length   = EXCLUDED.puff_length,
                timestamp     = EXCLUDED.timestamp,
                json_data     = EXCLUDED.json_data,
                updated_by    = EXCLUDED.updated_by
            RETURNING id
        )");
        q.addBindValue(s.sessionName);
        q.addBindValue(s.testerName);
        q.addBindValue(s.assessorName);
        q.addBindValue(s.media);
        q.addBindValue(s.puffLength);
        q.addBindValue(s.date);
        q.addBindValue(s.timestamp);
        q.addBindValue(jsonStr);
        q.addBindValue(who);
        if (!q.exec() || !q.next()) {
            m_lastError = q.lastError().text();
            db.rollback();
            logDebug("saveSensorySession UPSERT failed: " + m_lastError);
            return false;
        }
        sessionId = q.value(0).toLongLong();
    }

    // Wipe and rebuild this session's images. CASCADE wouldn't fire here
    // because the parent row was UPDATEd (not deleted) on the conflict path.
    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM sensory_images WHERE session_id = ?");
        del.addBindValue(static_cast<qlonglong>(sessionId));
        if (!del.exec()) {
            m_lastError = del.lastError().text();
            db.rollback();
            logDebug("saveSensorySession DELETE images failed: " + m_lastError);
            return false;
        }
    }

    QString imgErr;
    if (!insertImagesFor(db, "sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops, who, &imgErr)) {
        m_lastError = imgErr;
        db.rollback();
        logDebug("saveSensorySession INSERT image failed: " + m_lastError);
        return false;
    }

    if (!db.commit()) {
        m_lastError = db.lastError().text();
        db.rollback();
        logDebug("saveSensorySession commit failed: " + m_lastError);
        return false;
    }
    return true;
}

bool DatabaseManager::saveSensorySession(SensorySession& s)
{
    // Same upsert as the const overload, but we capture the RETURNING id into
    // s.id before any commit. Doing this inline avoids a follow-up SELECT
    // (which the old SQLite version needed because INSERT OR REPLACE assigned
    // a new rowid every time and lastInsertId() didn't survive a transaction
    // re-entry).
    if (!isOpen()) {
        m_lastError = QStringLiteral("database not open");
        return false;
    }

    const QString jsonStr = serializeSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("could not begin transaction: ")
                      + db.lastError().text();
        logDebug("saveSensorySession(byRef) " + m_lastError);
        return false;
    }

    qint64 sessionId = -1;
    {
        QSqlQuery q(db);
        q.prepare(R"(
            INSERT INTO sensory_sessions
                (session_name, tester_name, assessor_name, media, puff_length,
                 date, timestamp, json_data, layout_json, updated_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, CAST(? AS JSONB), NULL, ?)
            ON CONFLICT (session_name, tester_name, date) DO UPDATE SET
                assessor_name = EXCLUDED.assessor_name,
                media         = EXCLUDED.media,
                puff_length   = EXCLUDED.puff_length,
                timestamp     = EXCLUDED.timestamp,
                json_data     = EXCLUDED.json_data,
                updated_by    = EXCLUDED.updated_by
            RETURNING id
        )");
        q.addBindValue(s.sessionName);
        q.addBindValue(s.testerName);
        q.addBindValue(s.assessorName);
        q.addBindValue(s.media);
        q.addBindValue(s.puffLength);
        q.addBindValue(s.date);
        q.addBindValue(s.timestamp);
        q.addBindValue(jsonStr);
        q.addBindValue(who);
        if (!q.exec() || !q.next()) {
            m_lastError = q.lastError().text();
            db.rollback();
            logDebug("saveSensorySession(byRef) UPSERT failed: " + m_lastError);
            return false;
        }
        sessionId = q.value(0).toLongLong();
    }

    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM sensory_images WHERE session_id = ?");
        del.addBindValue(static_cast<qlonglong>(sessionId));
        if (!del.exec()) {
            m_lastError = del.lastError().text();
            db.rollback();
            return false;
        }
    }

    QString imgErr;
    if (!insertImagesFor(db, "sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops, who, &imgErr)) {
        m_lastError = imgErr;
        db.rollback();
        return false;
    }

    if (!db.commit()) {
        m_lastError = db.lastError().text();
        db.rollback();
        return false;
    }

    s.id = static_cast<int>(sessionId);
    return true;
}

QVector<SensorySession> DatabaseManager::loadSensorySessions() const
{
    QVector<SensorySession> result;
    if (!isOpen()) return result;

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT id, json_data FROM sensory_sessions ORDER BY id DESC");
    if (!q.exec()) return result;

    // Step 1: read every row's (id, json) into memory before we issue the
    // per-session image queries — re-entering the cursor on the same QSqlQuery
    // while another QSqlQuery is in flight can confuse the QPSQL driver.
    struct Row { qint64 id; QByteArray json; };
    QVector<Row> rows;
    while (q.next()) {
        Row r;
        r.id   = q.value(0).toLongLong();
        r.json = q.value(1).toString().toUtf8();
        rows.append(r);
    }

    for (const Row& r : rows) {
        SensorySession sess;
        if (!deserializeSensoryJson(r.json, sess)) continue;
        sess.id = static_cast<int>(r.id);

        loadImagesFor(db, "sensory_images", "dve_sensimg", r.id,
                      &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
        result.append(sess);
    }
    return result;
}

SensorySession DatabaseManager::loadSensorySession(int id) const
{
    SensorySession sess;
    if (!isOpen()) return sess;

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT json_data FROM sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec() || !q.next()) return sess;

    if (!deserializeSensoryJson(q.value(0).toString().toUtf8(), sess)) return sess;
    sess.id = id;

    loadImagesFor(db, "sensory_images", "dve_sensimg", id,
                  &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
    return sess;
}

QVector<SensoryRecord> DatabaseManager::listSensoryRecords() const
{
    QVector<SensoryRecord> result;
    if (!isOpen()) return result;

    QSqlQuery q(m_pg->queryDb());
    // Pull from columns first; tap json_data only for the extras the listing
    // needs (test_title, tester_name, sample count). jsonb_array_length is
    // O(1) on a JSONB so it's cheap to compute server-side per row.
    q.prepare("SELECT id, session_name, assessor_name, media, date, "
              "       json_data->>'test_title'   AS test_title, "
              "       json_data->>'tester_name'  AS tester_name, "
              "       COALESCE(jsonb_array_length(json_data->'samples'), 0) AS sample_count "
              "FROM sensory_sessions ORDER BY id DESC");
    if (!q.exec()) return result;

    while (q.next()) {
        SensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();
        rec.testTitle    = q.value(5).toString();
        rec.testerName   = q.value(6).toString();
        rec.sampleCount  = q.value(7).toInt();
        result.append(rec);
    }
    return result;
}

bool DatabaseManager::removeSensorySession(int id)
{
    if (!isOpen()) return false;
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }
    return true;
}

QString DatabaseManager::nextDefaultTestName() const
{
    if (!isOpen()) return QStringLiteral("test_0001");

    // Scan existing test_NNNN titles; return one past the max. Gaps in the
    // numbering are preserved on purpose — sequential is what users expect.
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT json_data->>'test_title' FROM sensory_sessions "
              "WHERE json_data->>'test_title' ~ '^test_[0-9]+$'");
    int maxNum = 0;
    const QRegularExpression rx(QStringLiteral("^test_(\\d+)$"));
    if (q.exec()) {
        while (q.next()) {
            const QString title = q.value(0).toString();
            const auto match = rx.match(title);
            if (match.hasMatch()) {
                const int num = match.captured(1).toInt();
                if (num > maxNum) maxNum = num;
            }
        }
    }
    return QString("test_%1").arg(maxNum + 1, 4, 10, QLatin1Char('0'));
}

// ============================================================================
//  Detailed Sensory Sessions
// ============================================================================
//
// Mirrors the sensory path. Notably, detailed_sensory_sessions has NO
// layout_json column — the radar chart layout for detailed sensory mode is
// not persisted separately. Only the json_data blob round-trips.

bool DatabaseManager::saveDetailedSensorySession(const DetailedSensorySession& s)
{
    if (!isOpen()) {
        m_lastError = QStringLiteral("database not open");
        return false;
    }

    const QString jsonStr = serializeDetailedSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("could not begin transaction: ")
                      + db.lastError().text();
        logDebug("saveDetailedSensorySession " + m_lastError);
        return false;
    }

    qint64 sessionId = -1;
    {
        QSqlQuery q(db);
        q.prepare(R"(
            INSERT INTO detailed_sensory_sessions
                (session_name, tester_name, assessor_name, media, date, timestamp,
                 json_data, updated_by)
            VALUES (?, ?, ?, ?, ?, ?, CAST(? AS JSONB), ?)
            ON CONFLICT (session_name, tester_name, date) DO UPDATE SET
                assessor_name = EXCLUDED.assessor_name,
                media         = EXCLUDED.media,
                timestamp     = EXCLUDED.timestamp,
                json_data     = EXCLUDED.json_data,
                updated_by    = EXCLUDED.updated_by
            RETURNING id
        )");
        q.addBindValue(s.sessionName);
        q.addBindValue(s.testerName);
        q.addBindValue(s.assessorName);
        q.addBindValue(s.media);
        q.addBindValue(s.date);
        q.addBindValue(s.timestamp);
        q.addBindValue(jsonStr);
        q.addBindValue(who);
        if (!q.exec() || !q.next()) {
            m_lastError = q.lastError().text();
            db.rollback();
            logDebug("saveDetailedSensorySession UPSERT failed: " + m_lastError);
            return false;
        }
        sessionId = q.value(0).toLongLong();
    }

    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM detailed_sensory_images WHERE session_id = ?");
        del.addBindValue(static_cast<qlonglong>(sessionId));
        if (!del.exec()) {
            m_lastError = del.lastError().text();
            db.rollback();
            return false;
        }
    }

    QString imgErr;
    if (!insertImagesFor(db, "detailed_sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops, who, &imgErr)) {
        m_lastError = imgErr;
        db.rollback();
        logDebug("saveDetailedSensorySession INSERT image failed: " + m_lastError);
        return false;
    }

    if (!db.commit()) {
        m_lastError = db.lastError().text();
        db.rollback();
        logDebug("saveDetailedSensorySession commit failed: " + m_lastError);
        return false;
    }
    return true;
}

QVector<DetailedSensorySession> DatabaseManager::loadDetailedSensorySessions() const
{
    QVector<DetailedSensorySession> result;
    if (!isOpen()) return result;

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT id, json_data FROM detailed_sensory_sessions ORDER BY id DESC");
    if (!q.exec()) return result;

    struct Row { qint64 id; QByteArray json; };
    QVector<Row> rows;
    while (q.next()) {
        Row r;
        r.id   = q.value(0).toLongLong();
        r.json = q.value(1).toString().toUtf8();
        rows.append(r);
    }

    for (const Row& r : rows) {
        DetailedSensorySession sess;
        if (!deserializeDetailedSensoryJson(r.json, sess)) continue;
        loadImagesFor(db, "detailed_sensory_images", "dve_detsensimg", r.id,
                      &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
        result.append(sess);
    }
    return result;
}

DetailedSensorySession DatabaseManager::loadDetailedSensorySession(int id) const
{
    DetailedSensorySession sess;
    if (!isOpen()) return sess;

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT json_data FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec() || !q.next()) return sess;

    if (!deserializeDetailedSensoryJson(q.value(0).toString().toUtf8(), sess)) return sess;

    loadImagesFor(db, "detailed_sensory_images", "dve_detsensimg", id,
                  &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
    return sess;
}

QVector<DetailedSensoryRecord> DatabaseManager::listDetailedSensoryRecords() const
{
    QVector<DetailedSensoryRecord> result;
    if (!isOpen()) return result;

    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT id, session_name, assessor_name, media, date, "
              "       json_data->>'test_title'  AS test_title, "
              "       json_data->>'tester_name' AS tester_name, "
              "       COALESCE(jsonb_array_length(json_data->'samples'), 0) AS sample_count "
              "FROM detailed_sensory_sessions ORDER BY id DESC");
    if (!q.exec()) return result;

    while (q.next()) {
        DetailedSensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();
        rec.testTitle    = q.value(5).toString();
        rec.testerName   = q.value(6).toString();
        rec.sampleCount  = q.value(7).toInt();
        result.append(rec);
    }
    return result;
}

bool DatabaseManager::removeDetailedSensorySession(int id)
{
    if (!isOpen()) return false;
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }
    return true;
}

// ============================================================================
//  Layout JSON persistence (sensory report preview)
// ============================================================================

QString DatabaseManager::loadSensoryLayout(int sessionId) const
{
    if (!isOpen()) return {};
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT layout_json FROM sensory_sessions WHERE id = ?");
    q.addBindValue(sessionId);
    if (!q.exec() || !q.next()) return {};
    return q.value(0).toString();
}

bool DatabaseManager::saveSensoryLayout(int sessionId, const QString& layoutJson)
{
    if (!isOpen()) {
        m_lastError = QStringLiteral("database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("UPDATE sensory_sessions "
              "SET layout_json = CAST(? AS JSONB), updated_by = ? "
              "WHERE id = ?");
    // Empty layoutJson → SQL NULL (preserves the "no layout yet" state and
    // sidesteps Postgres rejecting '' as invalid JSONB).
    q.addBindValue(layoutJson.isEmpty() ? QVariant() : QVariant(layoutJson));
    q.addBindValue(writerUuid(m_identity));
    q.addBindValue(sessionId);
    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }
    return true;
}

QString DatabaseManager::loadCumulativeLayout() const
{
    return getSetting("sensory.cumulative_layout");
}

bool DatabaseManager::saveCumulativeLayout(const QString& layoutJson)
{
    return setSetting("sensory.cumulative_layout", layoutJson);
}

// ============================================================================
//  Settings key/value store
// ============================================================================

bool DatabaseManager::setSetting(const QString& key, const QString& value)
{
    if (!isOpen()) {
        m_lastError = QStringLiteral("database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("INSERT INTO settings (key, value, updated_by) VALUES (?, ?, ?) "
              "ON CONFLICT (key) DO UPDATE SET "
              "value = EXCLUDED.value, "
              "updated_by = EXCLUDED.updated_by");
    q.addBindValue(key);
    q.addBindValue(value);
    q.addBindValue(writerUuid(m_identity));
    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }
    return true;
}

QString DatabaseManager::getSetting(const QString& key, const QString& defaultVal) const
{
    if (!isOpen()) return defaultVal;
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT value FROM settings WHERE key = ?");
    q.addBindValue(key);
    if (q.exec() && q.next()) return q.value(0).toString();
    return defaultVal;
}

} // namespace DVE

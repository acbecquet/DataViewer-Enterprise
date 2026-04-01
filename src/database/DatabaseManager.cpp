#include "DatabaseManager.h"
#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>
#include <QDateTime>
#include <QFileInfo>
#include <QFile>
#include <QDir>
#include <QStandardPaths>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QRegularExpression>

namespace DVE {

DatabaseManager::DatabaseManager(QObject* parent)
    : QObject(parent)
{}

DatabaseManager::~DatabaseManager()
{
    close();
}

void DatabaseManager::logDebug(const QString& msg) const
{
    qDebug() << "[DatabaseManager]" << msg;
}

bool DatabaseManager::open(const QString& dbPath)
{
    if (m_open) close();

    m_db = QSqlDatabase::addDatabase("QSQLITE", "dve_main");
    m_db.setDatabaseName(dbPath);

    if (!m_db.open()) {
        m_lastError = m_db.lastError().text();
        logDebug("Failed to open DB: " + m_lastError);
        return false;
    }

    // Enable foreign key enforcement (SQLite requires this per-connection).
    QSqlQuery pragma(m_db);
    pragma.exec("PRAGMA foreign_keys = ON");
    // WAL mode is unsafe over SMB/network shares — use DELETE journal + busy timeout instead
    bool isNetwork = dbPath.startsWith("//") || dbPath.startsWith("\\\\");
    if (isNetwork) {
        pragma.exec("PRAGMA journal_mode = DELETE");
        pragma.exec("PRAGMA synchronous = FULL");  // flush completely on every write
        pragma.exec("PRAGMA busy_timeout = 5000");
    } else {
        pragma.exec("PRAGMA journal_mode = WAL");
        pragma.exec("PRAGMA synchronous = NORMAL");
    }

    m_open = true;
    m_currentPath = dbPath;
    logDebug("Database opened: " + dbPath);
    return initSchema();
}

void DatabaseManager::close()
{
    if (m_open) {
        m_db.close();
        m_open = false;
    }
}

bool DatabaseManager::isOpen() const { return m_open; }

// ============================================================================
// Schema
// ============================================================================
bool DatabaseManager::initSchema()
{
    QSqlQuery q(m_db);

    // ── files ────────────────────────────────────────────────────────────────
    bool ok = q.exec(
        "CREATE TABLE IF NOT EXISTS files ("
        "  id               INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  file_path        TEXT NOT NULL,"
        "  file_name        TEXT NOT NULL,"
        "  loaded_at        TEXT NOT NULL,"
        "  template_version TEXT,"
        "  sheet_count      INTEGER DEFAULT 0,"
        "  sample_count     INTEGER DEFAULT 0"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // Add UNIQUE index on file_path (safe if already exists).
    q.exec("CREATE UNIQUE INDEX IF NOT EXISTS idx_files_path ON files(file_path)");

    // ── tests (sheets) ───────────────────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS tests ("
        "  id                 INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  file_id            INTEGER NOT NULL REFERENCES files(id) ON DELETE CASCADE,"
        "  sheet_name         TEXT NOT NULL,"
        "  template_version   TEXT,"
        "  overall_avg_tpm    REAL DEFAULT 0.0,"
        "  overall_stddev_tpm REAL DEFAULT 0.0,"
        "  is_raw_table       INTEGER DEFAULT 0,"
        "  sort_order         INTEGER DEFAULT 0"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── samples ──────────────────────────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS samples ("
        "  id                  INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  test_id             INTEGER NOT NULL REFERENCES tests(id) ON DELETE CASCADE,"
        "  sort_order          INTEGER DEFAULT 0,"
        "  sample_name         TEXT,"
        "  sample_id           TEXT,"
        "  date                TEXT,"
        "  tester              TEXT,"
        "  media               TEXT,"
        "  viscosity           REAL DEFAULT 0.0,"
        "  resistance          REAL DEFAULT 0.0,"
        "  voltage             REAL DEFAULT 0.0,"
        "  power               REAL DEFAULT 0.0,"
        "  heating_technology  TEXT,"
        "  puffing_regime      TEXT,"
        "  initial_oil_mass    REAL DEFAULT 0.0,"
        "  average_tpm         REAL DEFAULT 0.0,"
        "  stddev_tpm          REAL DEFAULT 0.0,"
        "  avg_power_density   REAL DEFAULT 0.0,"
        "  efficiency_percent  REAL DEFAULT 0.0,"
        "  total_oil_consumed  REAL DEFAULT 0.0,"
        "  total_puffs         INTEGER DEFAULT 0,"
        "  normalized_tpm      REAL DEFAULT 0.0,"
        "  burn_status         TEXT,"
        "  clog_status         TEXT,"
        "  leak_status         TEXT"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── data_rows ────────────────────────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS data_rows ("
        "  id                INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  sample_id         INTEGER NOT NULL REFERENCES samples(id) ON DELETE CASCADE,"
        "  sort_order        INTEGER DEFAULT 0,"
        "  puffs             REAL DEFAULT 0.0,"
        "  before_weight     REAL DEFAULT 0.0,"
        "  after_weight      REAL DEFAULT 0.0,"
        "  draw_pressure     REAL DEFAULT 0.0,"
        "  resistance        REAL DEFAULT 0.0,"
        "  smell             TEXT,"
        "  clog              TEXT,"
        "  notes             TEXT,"
        "  tpm               REAL DEFAULT 0.0,"
        "  tpm_power_density REAL DEFAULT 0.0,"
        "  variation_tpm     REAL DEFAULT 0.0,"
        "  oil_consumed      REAL DEFAULT 0.0"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── images ───────────────────────────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS images ("
        "  id               INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  sample_id        INTEGER NOT NULL REFERENCES samples(id) ON DELETE CASCADE,"
        "  sort_order       INTEGER DEFAULT 0,"
        "  file_name        TEXT,"
        "  image_data       BLOB,"
        "  layout_x         REAL,"
        "  layout_y         REAL,"
        "  layout_w         REAL,"
        "  layout_h         REAL,"
        "  crop_x           REAL,"
        "  crop_y           REAL,"
        "  crop_w           REAL,"
        "  crop_h           REAL"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── sensory_sessions ─────────────────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS sensory_sessions ("
        "  id            INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  session_name  TEXT,"
        "  tester_name   TEXT,"
        "  assessor_name TEXT,"
        "  media         TEXT,"
        "  puff_length   TEXT,"
        "  date          TEXT,"
        "  timestamp     TEXT,"
        "  json_data     TEXT"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── sensory_images (same pattern as images table, FK → sensory_sessions) ──
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS sensory_images ("
        "  id               INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  session_id       INTEGER NOT NULL REFERENCES sensory_sessions(id) ON DELETE CASCADE,"
        "  sort_order       INTEGER DEFAULT 0,"
        "  file_name        TEXT,"
        "  image_data       BLOB,"
        "  layout_x         REAL,"
        "  layout_y         REAL,"
        "  layout_w         REAL,"
        "  layout_h         REAL,"
        "  crop_x           REAL,"
        "  crop_y           REAL,"
        "  crop_w           REAL,"
        "  crop_h           REAL"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── detailed_sensory_sessions ───────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS detailed_sensory_sessions ("
        "  id            INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  session_name  TEXT,"
        "  tester_name   TEXT,"
        "  assessor_name TEXT,"
        "  media         TEXT,"
        "  date          TEXT,"
        "  timestamp     TEXT,"
        "  json_data     TEXT"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── detailed_sensory_images ─────────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS detailed_sensory_images ("
        "  id               INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  session_id       INTEGER NOT NULL REFERENCES detailed_sensory_sessions(id) ON DELETE CASCADE,"
        "  sort_order       INTEGER DEFAULT 0,"
        "  file_name        TEXT,"
        "  image_data       BLOB,"
        "  layout_x         REAL,"
        "  layout_y         REAL,"
        "  layout_w         REAL,"
        "  layout_h         REAL,"
        "  crop_x           REAL,"
        "  crop_y           REAL,"
        "  crop_w           REAL,"
        "  crop_h           REAL"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── Migration: add tester_name column if missing (existing databases) ────
    q.exec("ALTER TABLE sensory_sessions ADD COLUMN tester_name TEXT");
    // Silently fails if column already exists — that's fine.

    // ── settings (unchanged) ─────────────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS settings ("
        "  key   TEXT PRIMARY KEY,"
        "  value TEXT"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    return true;
}

// ============================================================================
// Hierarchical file save
// ============================================================================
bool DatabaseManager::saveFile(const FileResult& result)
{
    if (!m_open) return false;

    int totalSamples = 0;
    for (const auto& sheet : result.sheets)
        totalSamples += sheet.samples.size();

    m_db.transaction();

    QSqlQuery q(m_db);

    // ── Delete existing entry for this file_path (CASCADE clears children) ──
    q.prepare("SELECT id FROM files WHERE file_path = ?");
    q.addBindValue(result.filePath);
    if (q.exec() && q.next()) {
        int oldId = q.value(0).toInt();
        QSqlQuery del(m_db);
        del.prepare("DELETE FROM files WHERE id = ?");
        del.addBindValue(oldId);
        del.exec();
    }

    // ── Insert file ─────────────────────────────────────────────────────────
    q.prepare(
        "INSERT INTO files (file_path, file_name, loaded_at, template_version, sheet_count, sample_count) "
        "VALUES (?, ?, ?, ?, ?, ?)"
    );
    q.addBindValue(result.filePath);
    q.addBindValue(result.fileName);
    q.addBindValue(QDateTime::currentDateTime().toString(Qt::ISODate));
    q.addBindValue(result.templateVersion);
    q.addBindValue(result.sheets.size());
    q.addBindValue(totalSamples);
    if (!q.exec()) {
        m_lastError = q.lastError().text();
        m_db.rollback();
        logDebug("saveFile INSERT files failed: " + m_lastError);
        return false;
    }
    const int fileId = q.lastInsertId().toInt();

    // ── Prepare reusable statements for inner loops ─────────────────────────
    QSqlQuery insertTest(m_db);
    insertTest.prepare(
        "INSERT INTO tests (file_id, sheet_name, template_version, "
        "overall_avg_tpm, overall_stddev_tpm, is_raw_table, sort_order) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)"
    );

    QSqlQuery insertSample(m_db);
    insertSample.prepare(
        "INSERT INTO samples (test_id, sort_order, sample_name, sample_id, date, tester, "
        "media, viscosity, resistance, voltage, power, heating_technology, puffing_regime, "
        "initial_oil_mass, average_tpm, stddev_tpm, avg_power_density, efficiency_percent, "
        "total_oil_consumed, total_puffs, normalized_tpm, burn_status, clog_status, leak_status) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    );

    QSqlQuery insertRow(m_db);
    insertRow.prepare(
        "INSERT INTO data_rows (sample_id, sort_order, puffs, before_weight, after_weight, "
        "draw_pressure, resistance, smell, clog, notes, tpm, tpm_power_density, "
        "variation_tpm, oil_consumed) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    );

    QSqlQuery insertImage(m_db);
    insertImage.prepare(
        "INSERT INTO images (sample_id, sort_order, file_name, image_data, "
        "layout_x, layout_y, layout_w, layout_h, crop_x, crop_y, crop_w, crop_h) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    );

    // ── Insert tests → samples → data_rows + images ────────────────────────
    for (int si = 0; si < result.sheets.size(); ++si) {
        const SheetResult& sheet = result.sheets[si];

        insertTest.addBindValue(fileId);
        insertTest.addBindValue(sheet.sheetName);
        insertTest.addBindValue(sheet.templateVersion);
        insertTest.addBindValue(sheet.overallAvgTPM);
        insertTest.addBindValue(sheet.overallStdDevTPM);
        insertTest.addBindValue(sheet.isRawTable ? 1 : 0);
        insertTest.addBindValue(si);
        if (!insertTest.exec()) {
            m_lastError = insertTest.lastError().text();
            m_db.rollback();
            logDebug("saveFile INSERT test failed: " + m_lastError);
            return false;
        }
        const int testId = insertTest.lastInsertId().toInt();

        for (int sj = 0; sj < sheet.samples.size(); ++sj) {
            const SampleResult& sr = sheet.samples[sj];

            insertSample.addBindValue(testId);
            insertSample.addBindValue(sj);
            insertSample.addBindValue(sr.sampleName);
            insertSample.addBindValue(sr.sampleID);
            insertSample.addBindValue(sr.date);
            insertSample.addBindValue(sr.tester);
            insertSample.addBindValue(sr.media);
            insertSample.addBindValue(sr.viscosity);
            insertSample.addBindValue(sr.resistance);
            insertSample.addBindValue(sr.voltage);
            insertSample.addBindValue(sr.power);
            insertSample.addBindValue(sr.heatingTechnology);
            insertSample.addBindValue(sr.puffingRegime);
            insertSample.addBindValue(sr.initialOilMass);
            insertSample.addBindValue(sr.averageTPM);
            insertSample.addBindValue(sr.stdDevTPM);
            insertSample.addBindValue(sr.averagePowerDensity);
            insertSample.addBindValue(sr.efficiencyPercent);
            insertSample.addBindValue(sr.totalOilConsumed);
            insertSample.addBindValue(sr.totalPuffs);
            insertSample.addBindValue(sr.normalizedTPM);
            insertSample.addBindValue(sr.burnStatus);
            insertSample.addBindValue(sr.clogStatus);
            insertSample.addBindValue(sr.leakStatus);
            if (!insertSample.exec()) {
                m_lastError = insertSample.lastError().text();
                m_db.rollback();
                logDebug("saveFile INSERT sample failed: " + m_lastError);
                return false;
            }
            const int sampleId = insertSample.lastInsertId().toInt();

            // Data rows
            for (int ri = 0; ri < sr.rows.size(); ++ri) {
                const DataRow& dr = sr.rows[ri];
                insertRow.addBindValue(sampleId);
                insertRow.addBindValue(ri);
                insertRow.addBindValue(dr.puffs);
                insertRow.addBindValue(dr.beforeWeight);
                insertRow.addBindValue(dr.afterWeight);
                insertRow.addBindValue(dr.drawPressure);
                insertRow.addBindValue(dr.resistance);
                insertRow.addBindValue(dr.smell);
                insertRow.addBindValue(dr.clog);
                insertRow.addBindValue(dr.notes);
                insertRow.addBindValue(dr.tpm);
                insertRow.addBindValue(dr.tpmPowerDensity);
                insertRow.addBindValue(dr.variationTPM);
                insertRow.addBindValue(dr.oilConsumed);
                if (!insertRow.exec()) {
                    m_lastError = insertRow.lastError().text();
                    m_db.rollback();
                    logDebug("saveFile INSERT data_row failed: " + m_lastError);
                    return false;
                }
            }

            // Images
            for (int ii = 0; ii < sr.imagePaths.size(); ++ii) {
                QByteArray imgData;
                QFile imgFile(sr.imagePaths[ii]);
                if (imgFile.open(QIODevice::ReadOnly))
                    imgData = imgFile.readAll();

                QRectF layout = (ii < sr.imageLayouts.size()) ? sr.imageLayouts[ii] : QRectF();
                QRectF crop   = (ii < sr.imageCrops.size())   ? sr.imageCrops[ii]   : QRectF(0,0,1,1);

                insertImage.addBindValue(sampleId);
                insertImage.addBindValue(ii);
                insertImage.addBindValue(QFileInfo(sr.imagePaths[ii]).fileName());
                insertImage.addBindValue(imgData);
                insertImage.addBindValue(layout.x());
                insertImage.addBindValue(layout.y());
                insertImage.addBindValue(layout.width());
                insertImage.addBindValue(layout.height());
                insertImage.addBindValue(crop.x());
                insertImage.addBindValue(crop.y());
                insertImage.addBindValue(crop.width());
                insertImage.addBindValue(crop.height());
                if (!insertImage.exec()) {
                    m_lastError = insertImage.lastError().text();
                    m_db.rollback();
                    logDebug("saveFile INSERT image failed: " + m_lastError);
                    return false;
                }
            }
        }
    }

    m_db.commit();
    logDebug(QString("Saved file '%1' to database (fileId=%2, %3 sheets, %4 samples)")
                 .arg(result.fileName).arg(fileId)
                 .arg(result.sheets.size()).arg(totalSamples));
    return true;
}

// ============================================================================
// hasFile
// ============================================================================
bool DatabaseManager::hasFile(const QString& filePath) const
{
    if (!m_open) return false;
    QSqlQuery q(m_db);
    q.prepare("SELECT 1 FROM files WHERE file_path = ? LIMIT 1");
    q.addBindValue(filePath);
    return q.exec() && q.next();
}

// ============================================================================
// Load full FileResult from database
// ============================================================================
FileResult DatabaseManager::loadFile(int id) const
{
    FileResult result;
    if (!m_open) return result;

    // Step 1: Load file metadata
    {
        QSqlQuery qf(m_db);
        qf.prepare("SELECT file_path, file_name, template_version FROM files WHERE id = ?");
        qf.addBindValue(id);
        if (!qf.exec() || !qf.next()) return result;
        result.filePath        = qf.value(0).toString();
        result.fileName        = qf.value(1).toString();
        result.templateVersion = qf.value(2).toString();
    }

    // Step 2: Collect all test IDs and metadata (avoid nested cursor issues)
    struct TestInfo { int id; QString sheetName; QString templateVersion;
                      double avgTPM; double stddevTPM; bool isRaw; };
    QVector<TestInfo> tests;
    {
        QSqlQuery qt(m_db);
        qt.prepare("SELECT id, sheet_name, template_version, overall_avg_tpm, "
                   "overall_stddev_tpm, is_raw_table FROM tests "
                   "WHERE file_id = ? ORDER BY sort_order");
        qt.addBindValue(id);
        if (!qt.exec()) return result;
        while (qt.next()) {
            tests.append({qt.value(0).toInt(), qt.value(1).toString(),
                          qt.value(2).toString(), qt.value(3).toDouble(),
                          qt.value(4).toDouble(), qt.value(5).toBool()});
        }
    }

    // Step 3: For each test, collect sample IDs, then data rows and images
    for (const TestInfo& ti : tests) {
        SheetResult sheet;
        sheet.sheetName        = ti.sheetName;
        sheet.templateVersion  = ti.templateVersion;
        sheet.overallAvgTPM    = ti.avgTPM;
        sheet.overallStdDevTPM = ti.stddevTPM;
        sheet.isRawTable       = ti.isRaw;
        result.sheetNames.append(ti.sheetName);

        // Collect sample IDs + metadata
        struct SampleInfo {
            int id; QString name, sampleID, date, tester, media;
            double visc, res, volt, pwr; QString heatingTech, puffRegime;
            double initOil, avgTPM, stdDev, avgPD, effPct, totOil;
            int totPuffs; double normTPM;
            QString burn, clog, leak;
        };
        QVector<SampleInfo> sampleInfos;
        {
            QSqlQuery qs(m_db);
            qs.prepare("SELECT id, sample_name, sample_id, date, tester, media, viscosity, "
                       "resistance, voltage, power, heating_technology, puffing_regime, "
                       "initial_oil_mass, average_tpm, stddev_tpm, avg_power_density, "
                       "efficiency_percent, total_oil_consumed, total_puffs, normalized_tpm, "
                       "burn_status, clog_status, leak_status "
                       "FROM samples WHERE test_id = ? ORDER BY sort_order");
            qs.addBindValue(ti.id);
            if (qs.exec()) {
                while (qs.next()) {
                    SampleInfo si;
                    si.id         = qs.value(0).toInt();
                    si.name       = qs.value(1).toString();
                    si.sampleID   = qs.value(2).toString();
                    si.date       = qs.value(3).toString();
                    si.tester     = qs.value(4).toString();
                    si.media      = qs.value(5).toString();
                    si.visc       = qs.value(6).toDouble();
                    si.res        = qs.value(7).toDouble();
                    si.volt       = qs.value(8).toDouble();
                    si.pwr        = qs.value(9).toDouble();
                    si.heatingTech= qs.value(10).toString();
                    si.puffRegime = qs.value(11).toString();
                    si.initOil    = qs.value(12).toDouble();
                    si.avgTPM     = qs.value(13).toDouble();
                    si.stdDev     = qs.value(14).toDouble();
                    si.avgPD      = qs.value(15).toDouble();
                    si.effPct     = qs.value(16).toDouble();
                    si.totOil     = qs.value(17).toDouble();
                    si.totPuffs   = qs.value(18).toInt();
                    si.normTPM    = qs.value(19).toDouble();
                    si.burn       = qs.value(20).toString();
                    si.clog       = qs.value(21).toString();
                    si.leak       = qs.value(22).toString();
                    sampleInfos.append(si);
                }
            }
        }

        // Build SampleResults with rows and images (each query is non-nested now)
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

            // Load data rows
            {
                QSqlQuery qr(m_db);
                qr.prepare("SELECT puffs, before_weight, after_weight, draw_pressure, resistance, "
                           "smell, clog, notes, tpm, tpm_power_density, variation_tpm, oil_consumed "
                           "FROM data_rows WHERE sample_id = ? ORDER BY sort_order");
                qr.addBindValue(si.id);
                if (qr.exec()) {
                    while (qr.next()) {
                        DataRow dr;
                        dr.puffs           = qr.value(0).toDouble();
                        dr.beforeWeight    = qr.value(1).toDouble();
                        dr.afterWeight     = qr.value(2).toDouble();
                        dr.drawPressure    = qr.value(3).toDouble();
                        dr.resistance      = qr.value(4).toDouble();
                        dr.smell           = qr.value(5).toString();
                        dr.clog            = qr.value(6).toString();
                        dr.notes           = qr.value(7).toString();
                        dr.tpm             = qr.value(8).toDouble();
                        dr.tpmPowerDensity = qr.value(9).toDouble();
                        dr.variationTPM    = qr.value(10).toDouble();
                        dr.oilConsumed     = qr.value(11).toDouble();
                        sr.rows.append(dr);
                    }
                }
            }

            // Load images — write BLOBs to temp files so imagePaths works
            {
                QSqlQuery qi(m_db);
                qi.prepare("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                           "crop_x, crop_y, crop_w, crop_h "
                           "FROM images WHERE sample_id = ? ORDER BY sort_order");
                qi.addBindValue(si.id);
                if (qi.exec()) {
                    QString tempDir = QStandardPaths::writableLocation(
                                         QStandardPaths::AppLocalDataLocation) + "/ImageCache";
                    QDir().mkpath(tempDir);
                    while (qi.next()) {
                        QString fileName = qi.value(0).toString();
                        QByteArray blob  = qi.value(1).toByteArray();
                        QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                                      qi.value(4).toDouble(), qi.value(5).toDouble());
                        QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                                    qi.value(8).toDouble(), qi.value(9).toDouble());

                        QString tempPath = tempDir + "/" + QString::number(si.id) + "_" + fileName;
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

        // Rebuild tpmTrend/puffCounts (not stored in DB)
        for (const SampleResult& sr : sheet.samples) {
            for (const DataRow& dr : sr.rows) {
                if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
                sheet.tpmTrend.append(dr.tpm);
                sheet.puffCounts.append(dr.puffs);
            }
        }

        result.sheets.append(sheet);
    }

    logDebug(QString("Loaded file id=%1 '%2' from database (%3 sheets)")
                 .arg(id).arg(result.fileName).arg(result.sheets.size()));
    return result;
}

FileResult DatabaseManager::loadFileByPath(const QString& filePath) const
{
    FileResult result;
    if (!m_open) return result;

    QSqlQuery q(m_db);
    q.prepare("SELECT id FROM files WHERE file_path = ? LIMIT 1");
    q.addBindValue(filePath);
    if (q.exec() && q.next())
        return loadFile(q.value(0).toInt());
    return result;
}

// ============================================================================
// Deduplication
// ============================================================================
int DatabaseManager::deduplicateFiles(int keepPerName)
{
    if (!m_open) return 0;

    int deleted = 0;

    // 1. Delete entries with "unknown" template version (corrupt/empty)
    {
        QSqlQuery q(m_db);
        q.prepare("SELECT id FROM files WHERE template_version = 'unknown'");
        if (q.exec()) {
            QVector<int> ids;
            while (q.next()) ids.append(q.value(0).toInt());
            for (int id : ids) {
                removeFile(id);
                ++deleted;
            }
        }
    }

    // 2. For each file_name, keep only the N most recent entries
    {
        // Get distinct file names
        QSqlQuery qNames(m_db);
        qNames.exec("SELECT DISTINCT file_name FROM files");
        QStringList names;
        while (qNames.next()) names.append(qNames.value(0).toString());

        for (const QString& name : names) {
            QSqlQuery q(m_db);
            q.prepare("SELECT id FROM files WHERE file_name = ? ORDER BY loaded_at DESC");
            q.addBindValue(name);
            if (!q.exec()) continue;

            QVector<int> ids;
            while (q.next()) ids.append(q.value(0).toInt());

            // Delete all but the first keepPerName
            for (int i = keepPerName; i < ids.size(); ++i) {
                removeFile(ids[i]);
                ++deleted;
            }
        }
    }

    logDebug(QString("Deduplicated files: removed %1 entries").arg(deleted));
    return deleted;
}

// ============================================================================
// List / remove / recent / settings (unchanged logic)
// ============================================================================
QVector<FileRecord> DatabaseManager::listFiles() const
{
    QVector<FileRecord> records;
    if (!m_open) return records;

    QSqlQuery q("SELECT id, file_path, file_name, loaded_at, template_version, sheet_count, sample_count "
                "FROM files ORDER BY loaded_at DESC", m_db);

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

bool DatabaseManager::removeFile(int id)
{
    if (!m_open) return false;
    QSqlQuery q(m_db);
    q.prepare("DELETE FROM files WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) { m_lastError = q.lastError().text(); return false; }
    return true;
}

QStringList DatabaseManager::recentFilePaths() const
{
    QStringList paths;
    if (!m_open) return paths;
    QSqlQuery q("SELECT file_path FROM files ORDER BY loaded_at DESC LIMIT 20", m_db);
    while (q.next()) paths << q.value(0).toString();
    return paths;
}

// ============================================================================
// Sensory sessions
// ============================================================================
bool DatabaseManager::saveSensorySession(const SensorySession& s)
{
    if (!m_open) return false;

    // Serialize the session to JSON
    QJsonObject root;
    root["session_name"]  = s.sessionName;
    root["test_title"]    = s.testTitle;
    root["assessor_name"] = s.assessorName;
    root["tester_name"]   = s.testerName;
    root["media"]         = s.media;
    root["date"]          = s.date;
    root["timestamp"]     = s.timestamp;

    // New session-level test properties
    root["control"]              = s.control;
    root["is_blind"]             = s.isBlind;
    root["primary_differences"]  = s.primaryDifferences;

    // Legacy fields (backward compat)
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
        // Per-sample device properties
        sObj["voltage"]            = sample.voltage;
        sObj["resistance"]         = sample.resistance;
        sObj["power"]              = sample.power;
        sObj["heating_technology"] = sample.heatingTechnology;
        samplesArr.append(sObj);
    }
    root["samples"] = samplesArr;

    QString jsonStr = QString::fromUtf8(
        QJsonDocument(root).toJson(QJsonDocument::Compact));

    // Upsert: delete existing record matching session_name + tester_name + date
    // (different testers for the same test must NOT overwrite each other)
    {
        logDebug(QString("saveSensorySession: name='%1' tester='%2' date='%3' samples=%4")
                     .arg(s.sessionName, s.testerName, s.date)
                     .arg(s.samples.size()));

        // First delete orphaned images for the old session row
        QSqlQuery findOld(m_db);
        findOld.prepare("SELECT id FROM sensory_sessions "
                        "WHERE session_name = ? AND tester_name = ? AND date = ?");
        findOld.addBindValue(s.sessionName);
        findOld.addBindValue(s.testerName);
        findOld.addBindValue(s.date);
        if (findOld.exec()) {
            while (findOld.next()) {
                int oldId = findOld.value(0).toInt();
                logDebug(QString("  deleting old session id=%1").arg(oldId));
                QSqlQuery delImg(m_db);
                delImg.prepare("DELETE FROM sensory_images WHERE session_id = ?");
                delImg.addBindValue(oldId);
                delImg.exec();
            }
        }

        QSqlQuery del(m_db);
        del.prepare("DELETE FROM sensory_sessions "
                    "WHERE session_name = ? AND tester_name = ? AND date = ?");
        del.addBindValue(s.sessionName);
        del.addBindValue(s.testerName);
        del.addBindValue(s.date);
        del.exec();
        logDebug(QString("  deleted %1 old rows").arg(del.numRowsAffected()));
    }

    QSqlQuery q(m_db);
    q.prepare("INSERT INTO sensory_sessions "
              "(session_name, tester_name, assessor_name, media, puff_length, date, timestamp, json_data) "
              "VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
    q.addBindValue(s.sessionName);
    q.addBindValue(s.testerName);
    q.addBindValue(s.assessorName);
    q.addBindValue(s.media);
    q.addBindValue(s.puffLength);
    q.addBindValue(s.date);
    q.addBindValue(s.timestamp);
    q.addBindValue(jsonStr);

    if (!q.exec()) {
        m_lastError = q.lastError().text();
        logDebug("saveSensorySession failed: " + m_lastError);
        return false;
    }

    // Save images linked to this session (same pattern as TPM images)
    int sessionId = q.lastInsertId().toInt();
    if (!s.imagePaths.isEmpty()) {
        QSqlQuery imgQ(m_db);
        imgQ.prepare("INSERT INTO sensory_images "
                     "(session_id, sort_order, file_name, image_data,"
                     " layout_x, layout_y, layout_w, layout_h,"
                     " crop_x, crop_y, crop_w, crop_h) "
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
        for (int i = 0; i < s.imagePaths.size(); ++i) {
            QByteArray imgData;
            QFile imgFile(s.imagePaths[i]);
            if (imgFile.open(QIODevice::ReadOnly))
                imgData = imgFile.readAll();

            QRectF layout = (i < s.imageLayouts.size()) ? s.imageLayouts[i] : QRectF();
            QRectF crop   = (i < s.imageCrops.size())   ? s.imageCrops[i]   : QRectF(0,0,1,1);

            imgQ.addBindValue(sessionId);
            imgQ.addBindValue(i);
            imgQ.addBindValue(QFileInfo(s.imagePaths[i]).fileName());
            imgQ.addBindValue(imgData);
            imgQ.addBindValue(layout.x());
            imgQ.addBindValue(layout.y());
            imgQ.addBindValue(layout.width());
            imgQ.addBindValue(layout.height());
            imgQ.addBindValue(crop.x());
            imgQ.addBindValue(crop.y());
            imgQ.addBindValue(crop.width());
            imgQ.addBindValue(crop.height());
            imgQ.exec();
        }
    }

    return true;
}

QVector<SensorySession> DatabaseManager::loadSensorySessions() const
{
    QVector<SensorySession> result;
    if (!m_open) return result;

    QSqlQuery q("SELECT id, json_data FROM sensory_sessions ORDER BY id DESC", m_db);
    while (q.next()) {
        int sessId = q.value(0).toInt();
        QByteArray jsonBytes = q.value(1).toString().toUtf8();
        QJsonDocument doc    = QJsonDocument::fromJson(jsonBytes);
        if (doc.isNull() || !doc.isObject()) continue;

        QJsonObject root = doc.object();
        SensorySession sess;
        sess.sessionName  = root["session_name"].toString();
        sess.testTitle    = root["test_title"].toString();
        sess.assessorName = root["assessor_name"].toString();
        sess.testerName   = root["tester_name"].toString();
        sess.media        = root["media"].toString();
        sess.date         = root["date"].toString();
        sess.timestamp    = root["timestamp"].toString();

        // New session-level properties
        sess.control             = root["control"].toString();
        sess.isBlind             = root["is_blind"].toBool(false);
        sess.primaryDifferences  = root["primary_differences"].toString();

        // Legacy fields (backward compat)
        sess.puffLength          = root["puff_length"].toString();
        sess.burnStatus          = root["burn_status"].toString();
        sess.clogStatus          = root["clog_status"].toString();
        sess.leakStatus          = root["leak_status"].toString();
        sess.resistance          = root["resistance"].toDouble();
        sess.voltage             = root["voltage"].toDouble();
        sess.power               = root["power"].toDouble();
        sess.heatingTechnology   = root["heating_technology"].toString();

        for (const QJsonValue& sv : root["samples"].toArray()) {
            QJsonObject sObj = sv.toObject();
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

        // Load images linked to this session
        QSqlQuery qi(m_db);
        qi.prepare("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                   "crop_x, crop_y, crop_w, crop_h "
                   "FROM sensory_images WHERE session_id = ? ORDER BY sort_order");
        qi.addBindValue(sessId);
        if (qi.exec()) {
            int ii = 0;
            while (qi.next()) {
                QByteArray blob = qi.value(1).toByteArray();
                QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                             qi.value(4).toDouble(), qi.value(5).toDouble());
                QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                           qi.value(8).toDouble(), qi.value(9).toDouble());

                QString tempPath = QDir::temp().filePath(
                    QString("dve_sensimg_%1_%2.png").arg(sessId).arg(ii++));
                QFile tmpFile(tempPath);
                if (tmpFile.open(QIODevice::WriteOnly)) {
                    tmpFile.write(blob);
                    tmpFile.close();
                }
                sess.imagePaths.append(tempPath);
                sess.imageLayouts.append(layout);
                sess.imageCrops.append(crop);
            }
        }

        result.append(sess);
    }
    return result;
}

SensorySession DatabaseManager::loadSensorySession(int id) const
{
    SensorySession sess;
    if (!m_open) return sess;

    QSqlQuery q(m_db);
    q.prepare("SELECT json_data FROM sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec() || !q.next()) return sess;

    QByteArray jsonBytes = q.value(0).toString().toUtf8();
    QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
    if (doc.isNull() || !doc.isObject()) return sess;

    QJsonObject root = doc.object();
    sess.sessionName  = root["session_name"].toString();
    sess.testTitle    = root["test_title"].toString();
    sess.assessorName = root["assessor_name"].toString();
    sess.testerName   = root["tester_name"].toString();
    sess.media        = root["media"].toString();
    sess.date         = root["date"].toString();
    sess.timestamp    = root["timestamp"].toString();

    // New session-level properties
    sess.control             = root["control"].toString();
    sess.isBlind             = root["is_blind"].toBool(false);
    sess.primaryDifferences  = root["primary_differences"].toString();

    // Legacy fields (backward compat)
    sess.puffLength          = root["puff_length"].toString();
    sess.burnStatus          = root["burn_status"].toString();
    sess.clogStatus          = root["clog_status"].toString();
    sess.leakStatus          = root["leak_status"].toString();
    sess.resistance          = root["resistance"].toDouble();
    sess.voltage             = root["voltage"].toDouble();
    sess.power               = root["power"].toDouble();
    sess.heatingTechnology   = root["heating_technology"].toString();

    for (const QJsonValue& sv : root["samples"].toArray()) {
        QJsonObject sObj = sv.toObject();
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

    // Load images linked to this session
    QSqlQuery qi(m_db);
    qi.prepare("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
               "crop_x, crop_y, crop_w, crop_h "
               "FROM sensory_images WHERE session_id = ? ORDER BY sort_order");
    qi.addBindValue(id);
    if (qi.exec()) {
        int ii = 0;
        while (qi.next()) {
            QByteArray blob = qi.value(1).toByteArray();
            QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                         qi.value(4).toDouble(), qi.value(5).toDouble());
            QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                       qi.value(8).toDouble(), qi.value(9).toDouble());

            QString tempPath = QDir::temp().filePath(
                QString("dve_sensimg_%1_%2.png").arg(id).arg(ii++));
            QFile tmpFile(tempPath);
            if (tmpFile.open(QIODevice::WriteOnly)) {
                tmpFile.write(blob);
                tmpFile.close();
            }
            sess.imagePaths.append(tempPath);
            sess.imageLayouts.append(layout);
            sess.imageCrops.append(crop);
        }
    }

    return sess;
}

QVector<SensoryRecord> DatabaseManager::listSensoryRecords() const
{
    QVector<SensoryRecord> result;
    if (!m_open) return result;

    QSqlQuery q("SELECT id, session_name, assessor_name, media, date, json_data "
                "FROM sensory_sessions ORDER BY id DESC", m_db);
    while (q.next()) {
        SensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();

        // Extract test_title, tester_name, and sample count from JSON
        QByteArray jsonBytes = q.value(5).toString().toUtf8();
        QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
        if (!doc.isNull() && doc.isObject()) {
            QJsonObject root = doc.object();
            rec.testTitle    = root["test_title"].toString();
            rec.testerName   = root["tester_name"].toString();
            rec.sampleCount  = root["samples"].toArray().size();
        } else {
            rec.sampleCount = 0;
        }

        result.append(rec);
    }
    return result;
}

bool DatabaseManager::removeSensorySession(int id)
{
    if (!m_open) return false;
    QSqlQuery q(m_db);
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
    if (!m_open) return QStringLiteral("test_0001");

    // Count how many test_NNNN entries already exist
    QSqlQuery q("SELECT json_data FROM sensory_sessions", m_db);
    int maxNum = 0;
    QRegularExpression rx(QStringLiteral("^test_(\\d+)$"));
    while (q.next()) {
        QJsonDocument doc = QJsonDocument::fromJson(q.value(0).toString().toUtf8());
        if (doc.isObject()) {
            QString title = doc.object()["test_title"].toString();
            auto match = rx.match(title);
            if (match.hasMatch()) {
                int num = match.captured(1).toInt();
                if (num > maxNum) maxNum = num;
            }
        }
    }
    return QString("test_%1").arg(maxNum + 1, 4, 10, QLatin1Char('0'));
}

bool DatabaseManager::setSetting(const QString& key, const QString& value)
{
    if (!m_open) return false;
    QSqlQuery q(m_db);
    q.prepare("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)");
    q.addBindValue(key);
    q.addBindValue(value);
    if (!q.exec()) { m_lastError = q.lastError().text(); return false; }
    return true;
}

QString DatabaseManager::getSetting(const QString& key, const QString& defaultVal) const
{
    if (!m_open) return defaultVal;
    QSqlQuery q(m_db);
    q.prepare("SELECT value FROM settings WHERE key = ?");
    q.addBindValue(key);
    if (q.exec() && q.next()) return q.value(0).toString();
    return defaultVal;
}

// ============================================================================
// Detailed Sensory Sessions
// ============================================================================

bool DatabaseManager::saveDetailedSensorySession(const DetailedSensorySession& s)
{
    if (!m_open) return false;

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

    QString jsonStr = QString::fromUtf8(
        QJsonDocument(root).toJson(QJsonDocument::Compact));

    // Upsert: delete existing matching record
    {
        QSqlQuery findOld(m_db);
        findOld.prepare("SELECT id FROM detailed_sensory_sessions "
                        "WHERE session_name = ? AND tester_name = ? AND date = ?");
        findOld.addBindValue(s.sessionName);
        findOld.addBindValue(s.testerName);
        findOld.addBindValue(s.date);
        if (findOld.exec()) {
            while (findOld.next()) {
                int oldId = findOld.value(0).toInt();
                QSqlQuery delImg(m_db);
                delImg.prepare("DELETE FROM detailed_sensory_images WHERE session_id = ?");
                delImg.addBindValue(oldId);
                delImg.exec();
            }
        }

        QSqlQuery del(m_db);
        del.prepare("DELETE FROM detailed_sensory_sessions "
                    "WHERE session_name = ? AND tester_name = ? AND date = ?");
        del.addBindValue(s.sessionName);
        del.addBindValue(s.testerName);
        del.addBindValue(s.date);
        del.exec();
    }

    QSqlQuery q(m_db);
    q.prepare("INSERT INTO detailed_sensory_sessions "
              "(session_name, tester_name, assessor_name, media, date, timestamp, json_data) "
              "VALUES (?, ?, ?, ?, ?, ?, ?)");
    q.addBindValue(s.sessionName);
    q.addBindValue(s.testerName);
    q.addBindValue(s.assessorName);
    q.addBindValue(s.media);
    q.addBindValue(s.date);
    q.addBindValue(s.timestamp);
    q.addBindValue(jsonStr);

    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }

    // Save images
    int sessionId = q.lastInsertId().toInt();
    if (!s.imagePaths.isEmpty()) {
        QSqlQuery imgQ(m_db);
        imgQ.prepare("INSERT INTO detailed_sensory_images "
                     "(session_id, sort_order, file_name, image_data,"
                     " layout_x, layout_y, layout_w, layout_h,"
                     " crop_x, crop_y, crop_w, crop_h) "
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
        for (int i = 0; i < s.imagePaths.size(); ++i) {
            QByteArray imgData;
            QFile imgFile(s.imagePaths[i]);
            if (imgFile.open(QIODevice::ReadOnly))
                imgData = imgFile.readAll();

            QRectF layout = (i < s.imageLayouts.size()) ? s.imageLayouts[i] : QRectF();
            QRectF crop   = (i < s.imageCrops.size())   ? s.imageCrops[i]   : QRectF(0,0,1,1);

            imgQ.addBindValue(sessionId);
            imgQ.addBindValue(i);
            imgQ.addBindValue(QFileInfo(s.imagePaths[i]).fileName());
            imgQ.addBindValue(imgData);
            imgQ.addBindValue(layout.x());
            imgQ.addBindValue(layout.y());
            imgQ.addBindValue(layout.width());
            imgQ.addBindValue(layout.height());
            imgQ.addBindValue(crop.x());
            imgQ.addBindValue(crop.y());
            imgQ.addBindValue(crop.width());
            imgQ.addBindValue(crop.height());
            imgQ.exec();
        }
    }

    return true;
}

QVector<DetailedSensorySession> DatabaseManager::loadDetailedSensorySessions() const
{
    QVector<DetailedSensorySession> result;
    if (!m_open) return result;

    QSqlQuery q("SELECT id, json_data FROM detailed_sensory_sessions ORDER BY id DESC", m_db);
    while (q.next()) {
        int sessId = q.value(0).toInt();
        QByteArray jsonBytes = q.value(1).toString().toUtf8();
        QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
        if (doc.isNull() || !doc.isObject()) continue;

        QJsonObject root = doc.object();
        DetailedSensorySession sess;
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
        sess.deviceReturnDate   = root["device_return_date"].toString();
        sess.viscosity          = root["viscosity"].toString();

        for (const QJsonValue& sv : root["samples"].toArray()) {
            QJsonObject sObj = sv.toObject();
            DetailedSensorySample sample;
            sample.name              = sObj["name"].toString();
            sample.comments          = sObj["comments"].toString();
            sample.voltage           = sObj["voltage"].toDouble();
            sample.resistance        = sObj["resistance"].toDouble();
            sample.power             = sObj["power"].toDouble();
            sample.heatingTechnology = sObj["heating_technology"].toString();
            for (const QString& metric : kDetailedAllMetrics) {
                double maxVal = kDetailedMetricMaxScore.value(metric, 9);
                sample.scores[metric] = qBound(1.0, sObj[metric].toDouble(1.0), maxVal);
            }
            sess.samples.append(sample);
        }

        // Load images
        QSqlQuery qi(m_db);
        qi.prepare("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                   "crop_x, crop_y, crop_w, crop_h "
                   "FROM detailed_sensory_images WHERE session_id = ? ORDER BY sort_order");
        qi.addBindValue(sessId);
        if (qi.exec()) {
            int ii = 0;
            while (qi.next()) {
                QByteArray blob = qi.value(1).toByteArray();
                QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                             qi.value(4).toDouble(), qi.value(5).toDouble());
                QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                           qi.value(8).toDouble(), qi.value(9).toDouble());

                QString tempPath = QDir::temp().filePath(
                    QString("dve_detsensimg_%1_%2.png").arg(sessId).arg(ii++));
                QFile tmpFile(tempPath);
                if (tmpFile.open(QIODevice::WriteOnly)) {
                    tmpFile.write(blob);
                    tmpFile.close();
                }
                sess.imagePaths.append(tempPath);
                sess.imageLayouts.append(layout);
                sess.imageCrops.append(crop);
            }
        }

        result.append(sess);
    }
    return result;
}

DetailedSensorySession DatabaseManager::loadDetailedSensorySession(int id) const
{
    DetailedSensorySession sess;
    if (!m_open) return sess;

    QSqlQuery q(m_db);
    q.prepare("SELECT json_data FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec() || !q.next()) return sess;

    QByteArray jsonBytes = q.value(0).toString().toUtf8();
    QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
    if (doc.isNull() || !doc.isObject()) return sess;

    QJsonObject root = doc.object();
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
    sess.deviceReturnDate   = root["device_return_date"].toString();
    sess.viscosity          = root["viscosity"].toString();

    for (const QJsonValue& sv : root["samples"].toArray()) {
        QJsonObject sObj = sv.toObject();
        DetailedSensorySample sample;
        sample.name              = sObj["name"].toString();
        sample.comments          = sObj["comments"].toString();
        sample.voltage           = sObj["voltage"].toDouble();
        sample.resistance        = sObj["resistance"].toDouble();
        sample.power             = sObj["power"].toDouble();
        sample.heatingTechnology = sObj["heating_technology"].toString();
        for (const QString& metric : kDetailedAllMetrics) {
            double maxVal = kDetailedMetricMaxScore.value(metric, 9);
            sample.scores[metric] = qBound(1.0, sObj[metric].toDouble(1.0), maxVal);
        }
        sess.samples.append(sample);
    }

    // Load images
    QSqlQuery qi(m_db);
    qi.prepare("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
               "crop_x, crop_y, crop_w, crop_h "
               "FROM detailed_sensory_images WHERE session_id = ? ORDER BY sort_order");
    qi.addBindValue(id);
    if (qi.exec()) {
        int ii = 0;
        while (qi.next()) {
            QByteArray blob = qi.value(1).toByteArray();
            QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                         qi.value(4).toDouble(), qi.value(5).toDouble());
            QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                       qi.value(8).toDouble(), qi.value(9).toDouble());

            QString tempPath = QDir::temp().filePath(
                QString("dve_detsensimg_%1_%2.png").arg(id).arg(ii++));
            QFile tmpFile(tempPath);
            if (tmpFile.open(QIODevice::WriteOnly)) {
                tmpFile.write(blob);
                tmpFile.close();
            }
            sess.imagePaths.append(tempPath);
            sess.imageLayouts.append(layout);
            sess.imageCrops.append(crop);
        }
    }

    return sess;
}

QVector<DetailedSensoryRecord> DatabaseManager::listDetailedSensoryRecords() const
{
    QVector<DetailedSensoryRecord> result;
    if (!m_open) return result;

    QSqlQuery q("SELECT id, session_name, assessor_name, media, date, json_data "
                "FROM detailed_sensory_sessions ORDER BY id DESC", m_db);
    while (q.next()) {
        DetailedSensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();

        QByteArray jsonBytes = q.value(5).toString().toUtf8();
        QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
        if (!doc.isNull() && doc.isObject()) {
            QJsonObject root = doc.object();
            rec.testTitle    = root["test_title"].toString();
            rec.testerName   = root["tester_name"].toString();
            rec.sampleCount  = root["samples"].toArray().size();
        } else {
            rec.sampleCount = 0;
        }

        result.append(rec);
    }
    return result;
}

bool DatabaseManager::removeDetailedSensorySession(int id)
{
    if (!m_open) return false;
    QSqlQuery q(m_db);
    q.prepare("DELETE FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }
    return true;
}

} // namespace DVE

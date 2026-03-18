#include "DatabaseManager.h"
#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>
#include <QDateTime>
#include <QFileInfo>
#include <QFile>
#include <QDir>

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
    pragma.exec("PRAGMA journal_mode = WAL");

    m_open = true;
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
                    QString tempDir = QDir::tempPath() + "/dve_images";
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

} // namespace DVE

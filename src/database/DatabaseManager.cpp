#include "DatabaseManager.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "ConfigLoader.h"
#include "OfflineSnapshot.h"

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

// ── settings keys ───────────────────────────────────────────────────────────
// Cumulative-layout JSON for the multi-session sensory radar lives in the
// settings table under this key. Defined once so save and load can't drift.
namespace {
constexpr const char* kCumulativeLayoutKey = "sensory.cumulative_layout";

// Postgres SQLSTATE for unique_violation. Mapped to WriteResult::UniqueViolation
// by every tryWrite* method's INSERT branch.
constexpr const char* kSqlStateUniqueViolation = "23505";
}

// --- ctor / dtor / open / close ---------------------------------------------
DatabaseManager::DatabaseManager(QObject* parent)
    : QObject(parent), m_pg(new PostgresConnection(this)) {}

DatabaseManager::~DatabaseManager() { close(); }

bool DatabaseManager::open(const DbConfig& cfg, IdentityManager* identity) {
    m_identity = identity;
    m_online = false;
    if (!m_pg->open(cfg)) {
        m_lastError = QStringLiteral("open(connect): ") + m_pg->lastError();
        m_open = false;
        return false;
    }
    m_lastError.clear();
    m_open = true;
    // Default state on successful open: online + no snapshot. ConnectionMonitor
    // (C3) flips m_online to false when ping detects the server is unreachable.
    m_online = true;
    return true;
}

void DatabaseManager::close() {
    m_online = false;
    if (m_pg) m_pg->close();
    m_open = false;
}

bool DatabaseManager::isOpen() const {
    return m_open && m_pg && m_pg->isOpen();
}

// ── Offline mode (Plan C) ───────────────────────────────────────────────────
// Lifetime: m_snapshot is owned by MainWindow; DatabaseManager just holds a
// raw pointer. MainWindow guarantees it outlives every save/read path.
void DatabaseManager::setOfflineSnapshot(OfflineSnapshot* snap) {
    m_snapshot = snap;
}

// Soft state — set by ConnectionMonitor in response to a ping failure /
// reconnect, not by close(). close() also clears it for hygiene.
void DatabaseManager::setOnline(bool b) {
    m_online = b;
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

// --- helper: post-UPDATE rowcount-zero diagnostic ---------------------------
// Called when an optimistic UPDATE returned numRowsAffected() == 0. Issues a
// SELECT on the same id to distinguish "row exists with newer version"
// (VersionMismatch) from "row no longer exists" (RowDeleted). Any SQL error
// in the diagnostic itself collapses to OtherError so we never silently
// upgrade a conflict to success.
static WriteResult classifyMissingUpdate(QSqlDatabase& db,
                                         const QString& table,
                                         qint64 id,
                                         QString* outDetail)
{
    QSqlQuery q(db);
    q.prepare(QString("SELECT 1 FROM %1 WHERE id = ?").arg(table));
    q.addBindValue(static_cast<qlonglong>(id));
    if (!q.exec()) {
        if (outDetail) *outDetail = q.lastError().text();
        return WriteResult::OtherError;
    }
    return q.next() ? WriteResult::VersionMismatch : WriteResult::RowDeleted;
}

// ============================================================================
//  Hierarchical file storage
// ============================================================================
WriteResult DatabaseManager::tryWriteFile(const FileResult& result) {
    // Delegate to the mutable-ref overload via a local copy. The writeback
    // (post-save id + version) is discarded — callers who need it must
    // pass a mutable reference. Offline guard lives in the mutable overload
    // so this wrapper auto-inherits it.
    FileResult copy = result;
    return tryWriteFile(copy);
}

WriteResult DatabaseManager::tryWriteFile(FileResult& result) {
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteFile: database not open");
        return WriteResult::OtherError;
    }

    int totalSamples = 0;
    for (const auto& sheet : result.sheets)
        totalSamples += sheet.samples.size();

    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteFile(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    // -- Upsert the file row by file_path with optimistic concurrency. There
    //    are two branches:
    //      (a) result.id != -1 && result.version > 0 — caller has a server-
    //          loaded struct and wants to UPDATE that exact row. We require
    //          WHERE id = ? AND version = ? to refuse stale writes.
    //      (b) result.id == -1 — fresh struct, INSERT. UniqueViolation on
    //          file_path collision; caller is expected to recover (e.g.,
    //          load-then-merge) and re-issue.
    int fileId  = -1;
    int newVer  = 0;  // server-assigned version, captured via RETURNING
    if (result.id != -1 && result.version > 0) {
        QSqlQuery q(db);
        q.prepare(
            "UPDATE files SET "
            "  file_path = ?, "
            "  file_name = ?, "
            "  loaded_at = ?, "
            "  template_version = ?, "
            "  sheet_count = ?, "
            "  sample_count = ?, "
            "  updated_by = ? "
            "WHERE id = ? AND version = ? "
            "RETURNING version");
        q.addBindValue(result.filePath);
        q.addBindValue(result.fileName);
        q.addBindValue(QDateTime::currentDateTime().toString(Qt::ISODate));
        q.addBindValue(result.templateVersion);
        q.addBindValue(result.sheets.size());
        q.addBindValue(totalSamples);
        q.addBindValue(who);
        q.addBindValue(result.id);
        q.addBindValue(result.version);
        if (!q.exec()) {
            const QString code = q.lastError().nativeErrorCode();
            m_lastError = QStringLiteral("tryWriteFile(UPDATE files): ")
                          + q.lastError().text();
            db.rollback();
            logDebug(m_lastError);
            if (code == QString::fromLatin1(kSqlStateUniqueViolation))
                return WriteResult::UniqueViolation;
            return WriteResult::OtherError;
        }
        // RETURNING means a Success row is available via q.next(); absence
        // signals the optimistic-concurrency miss (the WHERE didn't match).
        if (!q.next()) {
            QString detail;
            const WriteResult cls = classifyMissingUpdate(
                db, QStringLiteral("files"), result.id, &detail);
            db.rollback();
            if (cls == WriteResult::VersionMismatch) {
                m_lastError = QStringLiteral(
                    "tryWriteFile(UPDATE files): version mismatch (id=%1, "
                    "expected version=%2)").arg(result.id).arg(result.version);
            } else if (cls == WriteResult::RowDeleted) {
                m_lastError = QStringLiteral(
                    "tryWriteFile(UPDATE files): row deleted (id=%1)").arg(result.id);
            } else {
                m_lastError = QStringLiteral(
                    "tryWriteFile(UPDATE files): classify failed: ") + detail;
            }
            logDebug(m_lastError);
            return cls;
        }
        fileId = result.id;
        newVer = q.value(0).toInt();
    } else {
        QSqlQuery q(db);
        q.prepare(
            "INSERT INTO files (file_path, file_name, loaded_at, template_version, "
            "sheet_count, sample_count, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?) RETURNING id, version");
        q.addBindValue(result.filePath);
        q.addBindValue(result.fileName);
        q.addBindValue(QDateTime::currentDateTime().toString(Qt::ISODate));
        q.addBindValue(result.templateVersion);
        q.addBindValue(result.sheets.size());
        q.addBindValue(totalSamples);
        q.addBindValue(who);
        if (!q.exec() || !q.next()) {
            const QString code = q.lastError().nativeErrorCode();
            m_lastError = QStringLiteral("tryWriteFile(INSERT files): ")
                          + q.lastError().text();
            db.rollback();
            logDebug(m_lastError);
            if (code == QString::fromLatin1(kSqlStateUniqueViolation))
                return WriteResult::UniqueViolation;
            return WriteResult::OtherError;
        }
        fileId = q.value(0).toInt();
        newVer = q.value(1).toInt();
    }

    // -- Wipe all children of this file and rebuild them. CASCADE on tests
    //    handles samples / data_rows / images automatically.
    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM tests WHERE file_id = ?");
        del.addBindValue(fileId);
        if (!del.exec()) {
            m_lastError = QStringLiteral("tryWriteFile(DELETE tests): ")
                          + del.lastError().text();
            db.rollback();
            logDebug(m_lastError);
            return WriteResult::OtherError;
        }
    }

    // -- Insert tests -> samples -> data_rows + images ----------------------
    // Prepare the four insert statements once outside the loops, then rebind
    // per iteration.
    QSqlQuery insertTest(db);
    if (!insertTest.prepare(
            "INSERT INTO tests (file_id, sheet_name, template_version, "
            "overall_avg_tpm, overall_stddev_tpm, is_raw_table, sort_order, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?) RETURNING id")) {
        m_lastError = QStringLiteral("tryWriteFile(prepare INSERT test): ")
                      + insertTest.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
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
        m_lastError = QStringLiteral("tryWriteFile(prepare INSERT sample): ")
                      + insertSample.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    QSqlQuery insertRow(db);
    if (!insertRow.prepare(
            "INSERT INTO data_rows (sample_id, sort_order, puffs, before_weight, after_weight, "
            "draw_pressure, resistance, smell, clog, notes, tpm, tpm_power_density, "
            "variation_tpm, oil_consumed, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")) {
        m_lastError = QStringLiteral("tryWriteFile(prepare INSERT data_row): ")
                      + insertRow.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    QSqlQuery insertImage(db);
    if (!insertImage.prepare(
            "INSERT INTO images (sample_id, sort_order, file_name, image_data, "
            "layout_x, layout_y, layout_w, layout_h, crop_x, crop_y, crop_w, crop_h, "
            "updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")) {
        m_lastError = QStringLiteral("tryWriteFile(prepare INSERT image): ")
                      + insertImage.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
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
                m_lastError = QStringLiteral("tryWriteFile(INSERT test): ")
                              + insertTest.lastError().text();
                db.rollback();
                logDebug(m_lastError);
                return WriteResult::OtherError;
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
                    m_lastError = QStringLiteral("tryWriteFile(INSERT sample): ")
                                  + insertSample.lastError().text();
                    db.rollback();
                    logDebug(m_lastError);
                    return WriteResult::OtherError;
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
                    m_lastError = QStringLiteral("tryWriteFile(INSERT data_row): ")
                                  + insertRow.lastError().text();
                    db.rollback();
                    logDebug(m_lastError);
                    return WriteResult::OtherError;
                }
            }

            // -- images (per-sample) --------------------------------------
            // Reads each on-disk image into a BYTEA blob and stores the
            // layout/crop rectangles. Skips files we can't open or that
            // exceed 100 MB.
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
                    m_lastError = QStringLiteral("tryWriteFile(INSERT image): ")
                                  + insertImage.lastError().text();
                    db.rollback();
                    logDebug(m_lastError);
                    return WriteResult::OtherError;
                }
            }
        }
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteFile(commit): ") + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }
    logDebug(QString("Saved file '%1' (fileId=%2, version=%3, %4 sheets, %5 samples)")
                 .arg(result.fileName).arg(fileId).arg(newVer)
                 .arg(result.sheets.size()).arg(totalSamples));
    // Writeback: parent file id + version only. Child ids (sample/data_row/
    // image) are intentionally NOT cascaded back — callers that need fresh
    // child ids should issue a follow-up loadFile(result.id) after this
    // returns. This keeps the writeback footprint small while letting the
    // recreate handler in MainWindow round-trip the parent id correctly.
    result.id      = fileId;
    result.version = newVer;
    return WriteResult::Success;
}

bool DatabaseManager::saveFile(const FileResult& result) {
    // Const-ref delegates to the const-ref tryWriteFile, which itself
    // delegates to the mutable-ref variant via a local copy and discards
    // the writeback — matching the legacy fire-and-forget semantics of
    // this bool shim.
    return tryWriteFile(result) == WriteResult::Success;
}

// --- hasFile ----------------------------------------------------------------
bool DatabaseManager::hasFile(const QString& filePath) const {
    m_lastError.clear();
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return !m_snapshot->loadFileByPath(filePath).filePath.isEmpty();
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("hasFile: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT 1 FROM files WHERE file_path = ? LIMIT 1");
    q.addBindValue(filePath);
    if (!q.exec()) {
        m_lastError = QStringLiteral("hasFile(select): ") + q.lastError().text();
        return false;
    }
    return q.next();
}

// --- loadFile ---------------------------------------------------------------
// Pure read path - no transaction. Walks files -> tests -> samples ->
// data_rows + images. SELECT now pulls id+version so subsequent saves can
// participate in optimistic concurrency.
FileResult DatabaseManager::loadFile(int id) const {
    m_lastError.clear();
    FileResult result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->loadFile(id);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadFile: database not open");
        return result;
    }

    QSqlDatabase& db = m_pg->queryDb();

    // Step 1: file metadata + id + version
    {
        QSqlQuery q(db);
        q.prepare("SELECT id, version, file_path, file_name, template_version "
                  "FROM files WHERE id = ?");
        q.addBindValue(id);
        if (!q.exec()) {
            m_lastError = QStringLiteral("loadFile(SELECT files): ")
                          + q.lastError().text();
            return result;
        }
        if (!q.next()) return result;  // not found, leave result empty
        result.id              = q.value(0).toInt();
        result.version         = q.value(1).toInt();
        result.filePath        = q.value(2).toString();
        result.fileName        = q.value(3).toString();
        result.templateVersion = q.value(4).toString();
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
        if (!q.exec()) {
            m_lastError = QStringLiteral("loadFile(SELECT tests): ")
                          + q.lastError().text();
            return result;
        }
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
            } else {
                m_lastError = QStringLiteral("loadFile(SELECT samples): ")
                              + q.lastError().text();
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
                q.prepare("SELECT id, puffs, before_weight, after_weight, draw_pressure, resistance, "
                          "smell, clog, notes, tpm, tpm_power_density, variation_tpm, oil_consumed "
                          "FROM data_rows WHERE sample_id = ? ORDER BY sort_order");
                q.addBindValue(si.id);
                if (q.exec()) {
                    while (q.next()) {
                        DataRow dr;
                        dr.id              = q.value(0).toInt();
                        dr.puffs           = q.value(1).toDouble();
                        dr.beforeWeight    = q.value(2).toDouble();
                        dr.afterWeight     = q.value(3).toDouble();
                        dr.drawPressure   = q.value(4).toDouble();
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
                    m_lastError = QStringLiteral("loadFile(SELECT data_rows): ")
                                  + q.lastError().text();
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
                } else {
                    m_lastError = QStringLiteral("loadFile(SELECT images): ")
                                  + q.lastError().text();
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

    logDebug(QString("Loaded file id=%1 '%2' (%3 sheets, version=%4)")
                 .arg(id).arg(result.fileName).arg(result.sheets.size())
                 .arg(result.version));
    return result;
}

FileResult DatabaseManager::loadFileByPath(const QString& filePath) const {
    m_lastError.clear();
    FileResult result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->loadFileByPath(filePath);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadFileByPath: database not open");
        return result;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT id FROM files WHERE file_path = ? LIMIT 1");
    q.addBindValue(filePath);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadFileByPath(SELECT files): ")
                      + q.lastError().text();
        return result;
    }
    if (q.next())
        return loadFile(q.value(0).toInt());
    return result;
}

// --- listFiles --------------------------------------------------------------
QVector<FileRecord> DatabaseManager::listFiles() const {
    m_lastError.clear();
    QVector<FileRecord> records;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->listFiles();
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return records;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("listFiles: database not open");
        return records;
    }

    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT id, file_path, file_name, loaded_at, template_version, "
              "sheet_count, sample_count FROM files ORDER BY loaded_at DESC");
    if (!q.exec()) {
        m_lastError = QStringLiteral("listFiles(SELECT files): ")
                      + q.lastError().text();
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

// --- removeFile -------------------------------------------------------------
bool DatabaseManager::removeFile(int id) {
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("removeFile: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM files WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("removeFile(DELETE files): ")
                      + q.lastError().text();
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
    m_lastError.clear();
    if (!m_online) {
        // Returns 0 (consistent with existing "early-out" semantics — the
        // method also returns 0 when nothing is removed).
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return 0;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("deduplicateFiles: database not open");
        return 0;
    }

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
        } else {
            m_lastError = QStringLiteral("deduplicateFiles(unknown-scan): ")
                          + q.lastError().text();
        }
    }

    // 2. per-file_name retention
    QStringList names;
    {
        QSqlQuery q(db);
        if (q.exec("SELECT DISTINCT file_name FROM files")) {
            while (q.next()) names.append(q.value(0).toString());
        } else {
            m_lastError = QStringLiteral("deduplicateFiles(name-scan): ")
                          + q.lastError().text();
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
    m_lastError.clear();
    QStringList paths;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            // OfflineSnapshot has no dedicated recentFilePaths; derive from
            // listFiles() (already sorted by loaded_at DESC) and cap at 20 to
            // match the online SELECT.
            const auto recs = m_snapshot->listFiles();
            for (int i = 0; i < recs.size() && paths.size() < 20; ++i)
                paths << recs[i].filePath;
            return paths;
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return paths;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("recentFilePaths: database not open");
        return paths;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT file_path FROM files ORDER BY loaded_at DESC LIMIT 20");
    if (!q.exec()) {
        m_lastError = QStringLiteral("recentFilePaths(SELECT files): ")
                      + q.lastError().text();
        return paths;
    }
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
// preserves it independently of saveSensorySession (saveSensoryLayout below
// UPDATEs only layout_json).
//
// Optimistic concurrency: when s.id != -1 && s.version > 0 the row is
// UPDATEd with WHERE id = ? AND version = ?; rowcount == 0 triggers a follow-
// up SELECT classified into VersionMismatch / RowDeleted. Fresh sessions
// (s.id == -1) INSERT and map SQLSTATE 23505 (duplicate natural-key) to
// UniqueViolation.

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
        sObj["power_type"]         = sample.powerType;
        sObj["puff_length_sec"]    = sample.puffLengthSec;
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
        // #7: power_type/puff_length_sec backward-compatible defaults preserved
        // when reading older rows that pre-date these fields.
        sample.powerType         = sObj.contains("power_type")
            ? sObj["power_type"].toString()
            : QStringLiteral("Constant Voltage");
        sample.puffLengthSec     = sObj.contains("puff_length_sec")
            ? sObj["puff_length_sec"].toDouble(3.0) : 3.0;
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

// -- Sensory save core --------------------------------------------------------
// Both saveSensorySession overloads (const and by-ref) share this body. The
// caller passes optional pointers for the post-write id and version, which
// the by-ref overload then propagates back into its struct.
//
// Three branches mirror tryWriteFile:
//   (a) s.id != -1 && s.version > 0 → UPDATE WHERE id=? AND version=?
//   (b) s.id == -1                    → INSERT
//   (c) UPSERT by natural key — used as a fallback when the caller doesn't
//       have id+version. NOT used here: we explicitly forbid silent UPSERTs
//       once optimistic concurrency is in play, because they mask
//       VersionMismatch into a successful overwrite of someone else's
//       changes. The const overload still has to support the "save a fresh
//       struct with possibly-conflicting natural key" case — that goes
//       through INSERT and surfaces UniqueViolation. Callers must then
//       load-then-merge.
WriteResult tryWriteSensoryCore(QSqlDatabase& db,
                                const SensorySession& s,
                                const QString& who,
                                const QString& jsonStr,
                                qint64* outId,
                                int* outVersion,
                                QString* outError)
{
    auto setError = [outError](const QString& msg) {
        if (outError) *outError = msg;
    };

    if (s.id != -1 && s.version > 0) {
        QSqlQuery q(db);
        q.prepare(R"(
            UPDATE sensory_sessions SET
                session_name  = ?,
                tester_name   = ?,
                assessor_name = ?,
                media         = ?,
                puff_length   = ?,
                date          = ?,
                timestamp     = ?,
                json_data     = CAST(? AS JSONB),
                updated_by    = ?
            WHERE id = ? AND version = ?
            RETURNING id, version
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
        q.addBindValue(s.id);
        q.addBindValue(s.version);
        if (!q.exec()) {
            const QString code = q.lastError().nativeErrorCode();
            setError(QStringLiteral("UPDATE sensory_sessions: ") + q.lastError().text());
            if (code == QString::fromLatin1(kSqlStateUniqueViolation))
                return WriteResult::UniqueViolation;
            return WriteResult::OtherError;
        }
        if (!q.next()) {
            // No row matched id+version. classifyMissingUpdate handles its
            // own error-text on internal SQL failure.
            QString detail;
            const WriteResult cls = classifyMissingUpdate(
                db, QStringLiteral("sensory_sessions"), s.id, &detail);
            if (cls == WriteResult::VersionMismatch) {
                setError(QStringLiteral("UPDATE sensory_sessions: version mismatch "
                                        "(id=%1, expected version=%2)")
                             .arg(s.id).arg(s.version));
            } else if (cls == WriteResult::RowDeleted) {
                setError(QStringLiteral("UPDATE sensory_sessions: row deleted "
                                        "(id=%1)").arg(s.id));
            } else {
                setError(QStringLiteral("UPDATE sensory_sessions classify: ") + detail);
            }
            return cls;
        }
        if (outId)      *outId = q.value(0).toLongLong();
        if (outVersion) *outVersion = q.value(1).toInt();
        return WriteResult::Success;
    }

    // INSERT branch — fresh struct. layout_json is NULL on insert; the
    // separate saveSensoryLayout path UPDATEs it later.
    QSqlQuery q(db);
    q.prepare(R"(
        INSERT INTO sensory_sessions
            (session_name, tester_name, assessor_name, media, puff_length,
             date, timestamp, json_data, layout_json, updated_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, CAST(? AS JSONB), NULL, ?)
        RETURNING id, version
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
        const QString code = q.lastError().nativeErrorCode();
        setError(QStringLiteral("INSERT sensory_sessions: ") + q.lastError().text());
        if (code == QString::fromLatin1(kSqlStateUniqueViolation))
            return WriteResult::UniqueViolation;
        return WriteResult::OtherError;
    }
    if (outId)      *outId = q.value(0).toLongLong();
    if (outVersion) *outVersion = q.value(1).toInt();
    return WriteResult::Success;
}

} // namespace

WriteResult DatabaseManager::tryWriteSensorySession(const SensorySession& s)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteSensorySession: database not open");
        return WriteResult::OtherError;
    }

    const QString jsonStr = serializeSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteSensorySession(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    logDebug(QString("tryWriteSensorySession: name='%1' tester='%2' date='%3' samples=%4 id=%5 v=%6")
                 .arg(s.sessionName, s.testerName, s.date)
                 .arg(s.samples.size()).arg(s.id).arg(s.version));

    qint64 sessionId = -1;
    int    newVer   = 0;
    QString coreErr;
    const WriteResult coreResult = tryWriteSensoryCore(db, s, who, jsonStr,
                                                       &sessionId, &newVer, &coreErr);
    if (coreResult != WriteResult::Success) {
        m_lastError = QStringLiteral("tryWriteSensorySession(") + coreErr + QStringLiteral(")");
        db.rollback();
        logDebug(m_lastError);
        return coreResult;
    }

    // Wipe and rebuild this session's images. CASCADE wouldn't fire here
    // because the parent row was UPDATEd (not deleted) on the conflict path.
    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM sensory_images WHERE session_id = ?");
        del.addBindValue(static_cast<qlonglong>(sessionId));
        if (!del.exec()) {
            m_lastError = QStringLiteral("tryWriteSensorySession(DELETE sensory_images): ")
                          + del.lastError().text();
            db.rollback();
            logDebug(m_lastError);
            return WriteResult::OtherError;
        }
    }

    QString imgErr;
    if (!insertImagesFor(db, "sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops, who, &imgErr)) {
        m_lastError = QStringLiteral("tryWriteSensorySession(INSERT sensory_images): ") + imgErr;
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteSensorySession(commit): ")
                      + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }
    return WriteResult::Success;
}

WriteResult DatabaseManager::tryWriteSensorySession(SensorySession& s)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteSensorySession: database not open");
        return WriteResult::OtherError;
    }

    const QString jsonStr = serializeSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    qint64 sessionId = -1;
    int    newVer    = 0;
    QString coreErr;
    const WriteResult coreResult = tryWriteSensoryCore(db, s, who, jsonStr,
                                                       &sessionId, &newVer, &coreErr);
    if (coreResult != WriteResult::Success) {
        m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(") + coreErr + QStringLiteral(")");
        db.rollback();
        logDebug(m_lastError);
        return coreResult;
    }

    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM sensory_images WHERE session_id = ?");
        del.addBindValue(static_cast<qlonglong>(sessionId));
        if (!del.exec()) {
            m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(DELETE sensory_images): ")
                          + del.lastError().text();
            db.rollback();
            logDebug(m_lastError);
            return WriteResult::OtherError;
        }
    }

    QString imgErr;
    if (!insertImagesFor(db, "sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops, who, &imgErr)) {
        m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(INSERT sensory_images): ")
                      + imgErr;
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(commit): ")
                      + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    s.id      = static_cast<int>(sessionId);
    s.version = newVer;
    return WriteResult::Success;
}

bool DatabaseManager::saveSensorySession(const SensorySession& s) {
    return tryWriteSensorySession(s) == WriteResult::Success;
}

bool DatabaseManager::saveSensorySession(SensorySession& s) {
    return tryWriteSensorySession(s) == WriteResult::Success;
}

QVector<SensorySession> DatabaseManager::loadSensorySessions() const
{
    m_lastError.clear();
    QVector<SensorySession> result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            // OfflineSnapshot exposes listSensoryRecords + per-id loadSensorySession
            // but no bulk loader. Derive the full list one at a time. OK for
            // typical session counts (< thousands).
            const auto recs = m_snapshot->listSensoryRecords();
            result.reserve(recs.size());
            for (const auto& r : recs) {
                SensorySession s = m_snapshot->loadSensorySession(r.id);
                if (s.id > 0) result.append(s);
            }
            return result;
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadSensorySessions: database not open");
        return result;
    }

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT id, version, json_data FROM sensory_sessions ORDER BY id DESC");
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadSensorySessions(SELECT): ")
                      + q.lastError().text();
        return result;
    }

    // Step 1: read every row's (id, version, json) into memory before we issue
    // the per-session image queries — re-entering the cursor on the same
    // QSqlQuery while another QSqlQuery is in flight can confuse the QPSQL
    // driver.
    struct Row { qint64 id; int version; QByteArray json; };
    QVector<Row> rows;
    while (q.next()) {
        Row r;
        r.id      = q.value(0).toLongLong();
        r.version = q.value(1).toInt();
        r.json    = q.value(2).toString().toUtf8();
        rows.append(r);
    }

    for (const Row& r : rows) {
        SensorySession sess;
        if (!deserializeSensoryJson(r.json, sess)) continue;
        sess.id      = static_cast<int>(r.id);
        sess.version = r.version;

        loadImagesFor(db, "sensory_images", "dve_sensimg", r.id,
                      &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
        result.append(sess);
    }
    return result;
}

SensorySession DatabaseManager::loadSensorySession(int id) const
{
    m_lastError.clear();
    SensorySession sess;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->loadSensorySession(id);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return sess;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadSensorySession: database not open");
        return sess;
    }

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT version, json_data FROM sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadSensorySession(SELECT): ")
                      + q.lastError().text();
        return sess;
    }
    if (!q.next()) return sess;  // not found

    const int rowVersion = q.value(0).toInt();
    if (!deserializeSensoryJson(q.value(1).toString().toUtf8(), sess)) return sess;
    sess.id      = id;
    sess.version = rowVersion;

    loadImagesFor(db, "sensory_images", "dve_sensimg", id,
                  &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
    return sess;
}

QVector<SensoryRecord> DatabaseManager::listSensoryRecords() const
{
    m_lastError.clear();
    QVector<SensoryRecord> result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->listSensoryRecords();
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("listSensoryRecords: database not open");
        return result;
    }

    QSqlQuery q(m_pg->queryDb());
    // Pull from columns first; tap json_data only for the extras the listing
    // needs (test_title, tester_name, sample count). jsonb_array_length is
    // O(1) on a JSONB so it's cheap to compute server-side per row.
    q.prepare("SELECT id, session_name, assessor_name, media, date, "
              "       json_data->>'test_title'   AS test_title, "
              "       json_data->>'tester_name'  AS tester_name, "
              "       COALESCE(jsonb_array_length(json_data->'samples'), 0) AS sample_count "
              "FROM sensory_sessions ORDER BY id DESC");
    if (!q.exec()) {
        m_lastError = QStringLiteral("listSensoryRecords(SELECT): ")
                      + q.lastError().text();
        return result;
    }

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
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("removeSensorySession: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("removeSensorySession(DELETE): ")
                      + q.lastError().text();
        return false;
    }
    return true;
}

QString DatabaseManager::nextDefaultTestName() const
{
    m_lastError.clear();
    if (!m_online) {
        // Naming a brand-new test is implicitly a write-side operation; even
        // if we returned a value derived from the snapshot, the user can't
        // actually create the session offline. Return the sentinel default
        // and surface the offline state via lastError().
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return QStringLiteral("test_0001");
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("nextDefaultTestName: database not open");
        return QStringLiteral("test_0001");
    }

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
    } else {
        m_lastError = QStringLiteral("nextDefaultTestName(SELECT): ")
                      + q.lastError().text();
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

namespace {

// Shared core for both tryWriteDetailedSensorySession overloads. Lives in the
// transaction the caller has already started; on Success populates outId and
// outVersion (server-assigned) so the mutable-ref overload can write them back.
//
// Branches mirror tryWriteSensoryCore:
//   (a) s.id != -1 && s.version > 0 → UPDATE WHERE id=? AND version=?
//                                     RETURNING id, version
//   (b) s.id == -1                    → INSERT RETURNING id, version
WriteResult tryWriteDetailedSensoryCore(QSqlDatabase& db,
                                        const DetailedSensorySession& s,
                                        const QString& who,
                                        const QString& jsonStr,
                                        qint64* outId,
                                        int* outVersion,
                                        QString* outError)
{
    auto setError = [outError](const QString& msg) {
        if (outError) *outError = msg;
    };

    if (s.id != -1 && s.version > 0) {
        QSqlQuery q(db);
        q.prepare(R"(
            UPDATE detailed_sensory_sessions SET
                session_name  = ?,
                tester_name   = ?,
                assessor_name = ?,
                media         = ?,
                date          = ?,
                timestamp     = ?,
                json_data     = CAST(? AS JSONB),
                updated_by    = ?
            WHERE id = ? AND version = ?
            RETURNING id, version
        )");
        q.addBindValue(s.sessionName);
        q.addBindValue(s.testerName);
        q.addBindValue(s.assessorName);
        q.addBindValue(s.media);
        q.addBindValue(s.date);
        q.addBindValue(s.timestamp);
        q.addBindValue(jsonStr);
        q.addBindValue(who);
        q.addBindValue(s.id);
        q.addBindValue(s.version);
        if (!q.exec()) {
            const QString code = q.lastError().nativeErrorCode();
            setError(QStringLiteral("UPDATE detailed_sensory_sessions: ") + q.lastError().text());
            if (code == QString::fromLatin1(kSqlStateUniqueViolation))
                return WriteResult::UniqueViolation;
            return WriteResult::OtherError;
        }
        if (!q.next()) {
            QString detail;
            const WriteResult cls = classifyMissingUpdate(
                db, QStringLiteral("detailed_sensory_sessions"), s.id, &detail);
            if (cls == WriteResult::VersionMismatch) {
                setError(QStringLiteral("UPDATE detailed_sensory_sessions: version mismatch "
                                        "(id=%1, expected version=%2)")
                             .arg(s.id).arg(s.version));
            } else if (cls == WriteResult::RowDeleted) {
                setError(QStringLiteral("UPDATE detailed_sensory_sessions: row deleted "
                                        "(id=%1)").arg(s.id));
            } else {
                setError(QStringLiteral("UPDATE detailed_sensory_sessions classify: ") + detail);
            }
            return cls;
        }
        if (outId)      *outId      = q.value(0).toLongLong();
        if (outVersion) *outVersion = q.value(1).toInt();
        return WriteResult::Success;
    }

    // INSERT branch — fresh struct.
    QSqlQuery q(db);
    q.prepare(R"(
        INSERT INTO detailed_sensory_sessions
            (session_name, tester_name, assessor_name, media, date, timestamp,
             json_data, updated_by)
        VALUES (?, ?, ?, ?, ?, ?, CAST(? AS JSONB), ?)
        RETURNING id, version
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
        const QString code = q.lastError().nativeErrorCode();
        setError(QStringLiteral("INSERT detailed_sensory_sessions: ") + q.lastError().text());
        if (code == QString::fromLatin1(kSqlStateUniqueViolation))
            return WriteResult::UniqueViolation;
        return WriteResult::OtherError;
    }
    if (outId)      *outId      = q.value(0).toLongLong();
    if (outVersion) *outVersion = q.value(1).toInt();
    return WriteResult::Success;
}

} // namespace

WriteResult DatabaseManager::tryWriteDetailedSensorySession(const DetailedSensorySession& s)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession: database not open");
        return WriteResult::OtherError;
    }

    const QString jsonStr = serializeDetailedSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    qint64 sessionId = -1;
    int    newVer   = 0;
    QString coreErr;
    const WriteResult coreResult = tryWriteDetailedSensoryCore(
        db, s, who, jsonStr, &sessionId, &newVer, &coreErr);
    if (coreResult != WriteResult::Success) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(") + coreErr + QStringLiteral(")");
        db.rollback();
        logDebug(m_lastError);
        return coreResult;
    }

    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM detailed_sensory_images WHERE session_id = ?");
        del.addBindValue(static_cast<qlonglong>(sessionId));
        if (!del.exec()) {
            m_lastError = QStringLiteral("tryWriteDetailedSensorySession(DELETE images): ")
                          + del.lastError().text();
            db.rollback();
            logDebug(m_lastError);
            return WriteResult::OtherError;
        }
    }

    QString imgErr;
    if (!insertImagesFor(db, "detailed_sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops, who, &imgErr)) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(INSERT images): ") + imgErr;
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(commit): ")
                      + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }
    return WriteResult::Success;
}

// Mutable-ref overload — see header. Mirrors the const-ref path but writes the
// post-save id+version back into `s`.
WriteResult DatabaseManager::tryWriteDetailedSensorySession(DetailedSensorySession& s)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession: database not open");
        return WriteResult::OtherError;
    }

    const QString jsonStr = serializeDetailedSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    qint64 sessionId = -1;
    int    newVer    = 0;
    QString coreErr;
    const WriteResult coreResult = tryWriteDetailedSensoryCore(
        db, s, who, jsonStr, &sessionId, &newVer, &coreErr);
    if (coreResult != WriteResult::Success) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(") + coreErr + QStringLiteral(")");
        db.rollback();
        logDebug(m_lastError);
        return coreResult;
    }

    {
        QSqlQuery del(db);
        del.prepare("DELETE FROM detailed_sensory_images WHERE session_id = ?");
        del.addBindValue(static_cast<qlonglong>(sessionId));
        if (!del.exec()) {
            m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(DELETE images): ")
                          + del.lastError().text();
            db.rollback();
            logDebug(m_lastError);
            return WriteResult::OtherError;
        }
    }

    QString imgErr;
    if (!insertImagesFor(db, "detailed_sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops, who, &imgErr)) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(INSERT images): ") + imgErr;
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(commit): ")
                      + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    s.id      = static_cast<int>(sessionId);
    s.version = newVer;
    return WriteResult::Success;
}

bool DatabaseManager::saveDetailedSensorySession(const DetailedSensorySession& s) {
    return tryWriteDetailedSensorySession(s) == WriteResult::Success;
}

QVector<DetailedSensorySession> DatabaseManager::loadDetailedSensorySessions() const
{
    m_lastError.clear();
    QVector<DetailedSensorySession> result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            const auto recs = m_snapshot->listDetailedSensoryRecords();
            result.reserve(recs.size());
            for (const auto& r : recs) {
                DetailedSensorySession s = m_snapshot->loadDetailedSensorySession(r.id);
                if (s.id > 0) result.append(s);
            }
            return result;
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadDetailedSensorySessions: database not open");
        return result;
    }

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT id, version, json_data FROM detailed_sensory_sessions ORDER BY id DESC");
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadDetailedSensorySessions(SELECT): ")
                      + q.lastError().text();
        return result;
    }

    struct Row { qint64 id; int version; QByteArray json; };
    QVector<Row> rows;
    while (q.next()) {
        Row r;
        r.id      = q.value(0).toLongLong();
        r.version = q.value(1).toInt();
        r.json    = q.value(2).toString().toUtf8();
        rows.append(r);
    }

    for (const Row& r : rows) {
        DetailedSensorySession sess;
        if (!deserializeDetailedSensoryJson(r.json, sess)) continue;
        sess.id      = static_cast<int>(r.id);
        sess.version = r.version;
        loadImagesFor(db, "detailed_sensory_images", "dve_detsensimg", r.id,
                      &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
        result.append(sess);
    }
    return result;
}

DetailedSensorySession DatabaseManager::loadDetailedSensorySession(int id) const
{
    m_lastError.clear();
    DetailedSensorySession sess;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->loadDetailedSensorySession(id);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return sess;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadDetailedSensorySession: database not open");
        return sess;
    }

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT version, json_data FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadDetailedSensorySession(SELECT): ")
                      + q.lastError().text();
        return sess;
    }
    if (!q.next()) return sess;

    const int rowVersion = q.value(0).toInt();
    if (!deserializeDetailedSensoryJson(q.value(1).toString().toUtf8(), sess)) return sess;
    sess.id      = id;
    sess.version = rowVersion;

    loadImagesFor(db, "detailed_sensory_images", "dve_detsensimg", id,
                  &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops);
    return sess;
}

QVector<DetailedSensoryRecord> DatabaseManager::listDetailedSensoryRecords() const
{
    m_lastError.clear();
    QVector<DetailedSensoryRecord> result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->listDetailedSensoryRecords();
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("listDetailedSensoryRecords: database not open");
        return result;
    }

    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT id, session_name, assessor_name, media, date, "
              "       json_data->>'test_title'  AS test_title, "
              "       json_data->>'tester_name' AS tester_name, "
              "       COALESCE(jsonb_array_length(json_data->'samples'), 0) AS sample_count "
              "FROM detailed_sensory_sessions ORDER BY id DESC");
    if (!q.exec()) {
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
        rec.testTitle    = q.value(5).toString();
        rec.testerName   = q.value(6).toString();
        rec.sampleCount  = q.value(7).toInt();
        result.append(rec);
    }
    return result;
}

bool DatabaseManager::removeDetailedSensorySession(int id)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("removeDetailedSensorySession: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("removeDetailedSensorySession(DELETE): ")
                      + q.lastError().text();
        return false;
    }
    return true;
}

// ============================================================================
//  Layout JSON persistence (sensory report preview)
// ============================================================================

QString DatabaseManager::loadSensoryLayout(int sessionId) const
{
    m_lastError.clear();
    if (!m_online) {
        // SensorySession has no layoutJson member, and OfflineSnapshot does
        // not currently expose layout_json through its loadSensorySession
        // accessor. The sensory report preview that consumes this layout is
        // a write-targeted UI (Save Layout button writes back), so offline
        // we surface an empty layout — the report builder degrades to the
        // default placement. Plan C C4/C5 may add a dedicated snapshot
        // accessor if the offline preview turns out to need real layouts.
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        Q_UNUSED(sessionId);
        return {};
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadSensoryLayout: database not open");
        return {};
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT layout_json FROM sensory_sessions WHERE id = ?");
    q.addBindValue(sessionId);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadSensoryLayout(SELECT): ")
                      + q.lastError().text();
        return {};
    }
    if (!q.next()) return {};
    return q.value(0).toString();
}

bool DatabaseManager::saveSensoryLayout(int sessionId, const QString& layoutJson)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("saveSensoryLayout: database not open");
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
        m_lastError = QStringLiteral("saveSensoryLayout(UPDATE): ")
                      + q.lastError().text();
        return false;
    }
    return true;
}

QString DatabaseManager::loadCumulativeLayout() const
{
    return getSetting(QString::fromLatin1(kCumulativeLayoutKey));
}

bool DatabaseManager::saveCumulativeLayout(const QString& layoutJson)
{
    return setSetting(QString::fromLatin1(kCumulativeLayoutKey), layoutJson);
}

// ============================================================================
//  Settings key/value store
// ============================================================================

bool DatabaseManager::setSetting(const QString& key, const QString& value)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("setSetting: database not open");
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
        m_lastError = QStringLiteral("setSetting(UPSERT): ")
                      + q.lastError().text();
        return false;
    }
    return true;
}

QString DatabaseManager::getSetting(const QString& key, const QString& defaultVal) const
{
    m_lastError.clear();
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->getSetting(key, defaultVal);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return defaultVal;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("getSetting: database not open");
        return defaultVal;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT value FROM settings WHERE key = ?");
    q.addBindValue(key);
    if (!q.exec()) {
        m_lastError = QStringLiteral("getSetting(SELECT): ")
                      + q.lastError().text();
        return defaultVal;
    }
    if (q.next()) return q.value(0).toString();
    return defaultVal;
}

} // namespace DVE

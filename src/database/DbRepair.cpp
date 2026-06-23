#include "DbRepair.h"
#include "DatabaseManager.h"
#include "../pipeline/DataProcessor.h"
#include "../pipeline/ReportData.h"

#include <QDebug>
#include <QDir>
#include <QDirIterator>
#include <QFile>
#include <QFileInfo>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QTextStream>

namespace DVE {

// ── isFileComplete ───────────────────────────────────────────────────────────
// A FileResult loaded from the DB is considered complete when EVERY normal
// (non-raw-table) sheet with at least one sample that has averageTPM > 0 also
// has at least one DataRow. Raw-table sheets are complete when rawHeaders is
// populated (or dbDataIncomplete is false).
static bool isFileComplete(const FileResult& fr)
{
    for (const SheetResult& sheet : fr.sheets) {
        if (sheet.isRawTable) {
            if (sheet.dbDataIncomplete)
                return false;
            continue;
        }
        for (const SampleResult& sample : sheet.samples) {
            if (sample.averageTPM > 0.0 && sample.rows.isEmpty())
                return false;
        }
    }
    return true;
}

// ── locateXlsx ───────────────────────────────────────────────────────────────
// Strategy 1: stored file_path
// Strategy 2: recursive search by file name under opts.sourceDir
static QString locateXlsx(const FileRecord& rec, const QString& sourceDir)
{
    if (!rec.filePath.isEmpty() && QFile::exists(rec.filePath))
        return rec.filePath;

    if (!sourceDir.isEmpty() && !rec.fileName.isEmpty()) {
        QDirIterator it(sourceDir,
                        QStringList{rec.fileName},
                        QDir::Files,
                        QDirIterator::Subdirectories);
        if (it.hasNext())
            return it.next();
    }
    return {};
}

// ── mergeIdVersion ───────────────────────────────────────────────────────────
// Copy id/version anchors from the DB record into the live result so the
// UPDATE path fires instead of INSERT.  Also match sheet/sample ids by name.
static FileResult mergeIdVersion(const FileResult& live, const FileResult& dbRec)
{
    FileResult merged  = live;
    merged.id          = dbRec.id;
    merged.version     = dbRec.version;

    for (SheetResult& mSheet : merged.sheets) {
        for (const SheetResult& dSheet : dbRec.sheets) {
            if (dSheet.sheetName == mSheet.sheetName) {
                mSheet.id      = dSheet.id;
                mSheet.version = dSheet.version;
                for (SampleResult& mSamp : mSheet.samples) {
                    for (const SampleResult& dSamp : dSheet.samples) {
                        if (dSamp.sampleName == mSamp.sampleName) {
                            mSamp.id      = dSamp.id;
                            mSamp.version = dSamp.version;
                            break;
                        }
                    }
                }
                break;
            }
        }
    }
    return merged;
}

// ── backfillFile (public, testable core) ─────────────────────────────────────
bool backfillFile(DatabaseManager& db,
                  const FileResult& liveResult,
                  const FileResult& dbRecord,
                  bool dryRun,
                  QString& detail)
{
    if (dryRun) {
        detail = QStringLiteral("dryRun -- would heal %1 (db id=%2)")
                     .arg(dbRecord.fileName)
                     .arg(dbRecord.id);
        return true;
    }

    FileResult toSave = mergeIdVersion(liveResult, dbRecord);
    const WriteResult wr = db.tryWriteFile(toSave);
    if (wr == WriteResult::Success) {
        detail = QStringLiteral("healed %1 (db id=%2 -> new version=%3)")
                     .arg(dbRecord.fileName)
                     .arg(dbRecord.id)
                     .arg(toSave.version);
        return true;
    }

    detail = QStringLiteral("FAILED %1 (db id=%2): writeResult=%3 err=%4")
                 .arg(dbRecord.fileName)
                 .arg(dbRecord.id)
                 .arg(static_cast<int>(wr))
                 .arg(db.lastError());
    return false;
}

// ── runDbRepair (full loop) ───────────────────────────────────────────────────
RepairSummary runDbRepair(DatabaseManager& db,
                          DataProcessor&   proc,
                          const RepairOptions& opts)
{
    RepairSummary summary;

    const QVector<FileRecord> records = db.listFiles();
    qInfo().noquote() << QStringLiteral("[DbRepair] Checking %1 DB file records...")
                             .arg(records.size());

    int processed = 0;
    const int total = static_cast<int>(records.size());
    for (const FileRecord& rec : records) {
        if (opts.progress) opts.progress(processed++, total, rec.fileName);
        const FileResult dbFr = db.loadFile(rec.id);
        if (dbFr.filePath.isEmpty()) {
            const QString d = QStringLiteral("FAILED to load DB record id=%1 (%2): %3")
                                  .arg(rec.id).arg(rec.fileName).arg(db.lastError());
            qWarning().noquote() << "[DbRepair]" << d;
            summary.details << d;
            ++summary.failed;
            continue;
        }

        if (isFileComplete(dbFr)) {
            const QString d = QStringLiteral("already-complete: %1 (id=%2)")
                                  .arg(rec.fileName).arg(rec.id);
            qDebug().noquote() << "[DbRepair]" << d;
            summary.details << d;
            ++summary.alreadyComplete;
            continue;
        }

        const QString xlsxPath = locateXlsx(rec, opts.sourceDir);
        if (xlsxPath.isEmpty()) {
            const QString d = QStringLiteral("skipped-no-xlsx: %1 (id=%2, stored=%3)")
                                  .arg(rec.fileName).arg(rec.id).arg(rec.filePath);
            qWarning().noquote() << "[DbRepair]" << d;
            summary.details << d;
            ++summary.skippedNoXlsx;
            continue;
        }

        const FileResult liveFr = proc.processFile(xlsxPath);
        if (liveFr.filePath.isEmpty()) {
            const QString d = QStringLiteral("FAILED processFile %1 (id=%2): %3")
                                  .arg(xlsxPath).arg(rec.id).arg(proc.lastError());
            qWarning().noquote() << "[DbRepair]" << d;
            summary.details << d;
            ++summary.failed;
            continue;
        }

        QString detail;
        const bool ok = backfillFile(db, liveFr, dbFr, opts.dryRun, detail);
        qInfo().noquote() << "[DbRepair]" << detail;
        summary.details << detail;
        if (ok)
            ++summary.healed;
        else
            ++summary.failed;
    }

    if (!opts.reportPath.isEmpty()) {
        QJsonArray detailArr;
        for (const QString& d : summary.details)
            detailArr.append(d);

        QJsonObject root;
        root[QStringLiteral("healed")]          = summary.healed;
        root[QStringLiteral("alreadyComplete")] = summary.alreadyComplete;
        root[QStringLiteral("skippedNoXlsx")]   = summary.skippedNoXlsx;
        root[QStringLiteral("failed")]          = summary.failed;
        root[QStringLiteral("dryRun")]          = opts.dryRun;
        root[QStringLiteral("details")]         = detailArr;

        QFile f(opts.reportPath);
        if (f.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
            f.write(QJsonDocument(root).toJson(QJsonDocument::Indented));
            f.close();
            qInfo().noquote() << "[DbRepair] Report written to:" << opts.reportPath;
        } else {
            qWarning().noquote() << "[DbRepair] Could not write report to:"
                                 << opts.reportPath;
        }
    }

    return summary;
}

} // namespace DVE

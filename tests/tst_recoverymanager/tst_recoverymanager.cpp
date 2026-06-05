#include <QtTest/QtTest>

#include <QFile>
#include <QJsonObject>
#include <QStandardPaths>
#include <QTemporaryDir>

#include "pipeline/ReportData.h"
#include "pipeline/ReportDataJson.h"
#include "utils/RecoveryManager.h"

using namespace DVE;

// Task C3: store primitives for the Plan C auto-recovery on-disk store.
// Round-trips two entries (one Tpm whose payload is a fileResultToJson tree,
// one Sensory with a plain QJsonObject) through writeItem -> readAll and
// asserts kind/id/displayName/dirty + the full payload survive, and that
// index.json lands on disk.
class TstRecoveryManager : public QObject
{
    Q_OBJECT

private:
    // A small but non-trivial FileResult to serialize into the Tpm blob.
    static FileResult makeFile()
    {
        FileResult f;
        f.filePath        = QStringLiteral("C:/data/Acme Device.xlsx");
        f.fileName        = QStringLiteral("Acme Device.xlsx");
        f.templateVersion = QStringLiteral("new");
        f.sheetNames      << QStringLiteral("Test 1");
        f.id              = static_cast<qint64>(500000000123LL);
        f.version         = 9;

        SheetResult sh;
        sh.sheetName       = QStringLiteral("Test 1");
        sh.templateVersion = QStringLiteral("new");
        sh.overallAvgTPM   = 6.66;
        sh.id              = static_cast<qint64>(600000000050LL);
        sh.version         = 4;

        SampleResult s;
        s.sampleName = QStringLiteral("Sample Alpha");
        s.sampleID   = QStringLiteral("SA-001");
        s.averageTPM = 5.55;
        s.id         = static_cast<qint64>(800000000099LL);
        s.version    = 7;

        DataRow r;
        r.puffs        = 50.0;
        r.beforeWeight = 12.5;
        r.afterWeight  = 11.5;
        r.tpm          = 5.0;
        r.id           = static_cast<qint64>(900000000001LL);
        r.version      = 2;
        s.rows.append(r);

        sh.samples.append(s);
        f.sheets.append(sh);
        return f;
    }

private slots:
    void initTestCase()
    {
        // Hygiene: keep any AppLocalDataLocation fallback inside test scratch
        // space rather than the real per-user profile.
        QStandardPaths::setTestModeEnabled(true);
    }

    void storeRoundTrip()
    {
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());

        RecoveryManager mgr;
        mgr.setDirOverride(tmp.path());

        // ---- Entry 1: a TPM file, payload = fileResultToJson(tree) ----
        const FileResult file = makeFile();
        const QJsonObject tpmPayload = fileResultToJson(file);

        RecoveryEntry tpmEntry;
        tpmEntry.kind        = RecoveryKind::Tpm;
        tpmEntry.id          = file.filePath;
        tpmEntry.displayName = file.fileName;
        tpmEntry.sourcePath  = file.filePath;
        tpmEntry.dirty       = true;
        QVERIFY(mgr.writeItem(tpmEntry, tpmPayload));

        // ---- Entry 2: a sensory session, small QJsonObject payload ----
        QJsonObject sensPayload;
        sensPayload["session_name"] = QStringLiteral("Panel 2026-06-04");
        sensPayload["tester"]       = QStringLiteral("CB");
        sensPayload["score"]        = 4.5;

        RecoveryEntry sensEntry;
        sensEntry.kind        = RecoveryKind::Sensory;
        sensEntry.id          = QStringLiteral("session-7");
        sensEntry.displayName = QStringLiteral("Panel 2026-06-04");
        sensEntry.sourcePath  = QString();
        sensEntry.dirty       = false;
        QVERIFY(mgr.writeItem(sensEntry, sensPayload));

        // ---- index.json must exist on disk ----
        QVERIFY(QFile::exists(mgr.liveDir() + "/index.json"));

        // ---- readAll returns both, with matching metadata + payloads ----
        const QVector<RecoveryEntry> all = mgr.readAll(mgr.liveDir());
        QCOMPARE(all.size(), 2);

        // Index order is insertion order (replace-or-append), so [0]=tpm, [1]=sensory.
        const RecoveryEntry& got0 = all[0];
        QCOMPARE(got0.kind, RecoveryKind::Tpm);
        QCOMPARE(got0.id, tpmEntry.id);
        QCOMPARE(got0.displayName, tpmEntry.displayName);
        QCOMPARE(got0.sourcePath, tpmEntry.sourcePath);
        QCOMPARE(got0.dirty, true);
        QCOMPARE(got0.payload, tpmPayload);

        const RecoveryEntry& got1 = all[1];
        QCOMPARE(got1.kind, RecoveryKind::Sensory);
        QCOMPARE(got1.id, sensEntry.id);
        QCOMPARE(got1.displayName, sensEntry.displayName);
        QCOMPARE(got1.dirty, false);
        QCOMPARE(got1.payload, sensPayload);
    }
};

QTEST_MAIN(TstRecoveryManager)
#include "tst_recoverymanager.moc"

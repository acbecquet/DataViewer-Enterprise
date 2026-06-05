#include <QtTest/QtTest>

#include <QDir>
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

    // Helper: write two distinct entries (one Tpm, one Sensory) into the live
    // store of a freshly-overridden manager. Returns the two payloads so callers
    // can assert they survive a move-to-_prev.
    static void seedTwo(RecoveryManager& mgr,
                        QJsonObject& tpmPayloadOut,
                        QJsonObject& sensPayloadOut)
    {
        tpmPayloadOut = fileResultToJson(makeFile());
        RecoveryEntry tpmEntry;
        tpmEntry.kind        = RecoveryKind::Tpm;
        tpmEntry.id          = QStringLiteral("C:/data/Acme Device.xlsx");
        tpmEntry.displayName = QStringLiteral("Acme Device.xlsx");
        tpmEntry.sourcePath  = QStringLiteral("C:/data/Acme Device.xlsx");
        tpmEntry.dirty       = true;
        QVERIFY(mgr.writeItem(tpmEntry, tpmPayloadOut));

        sensPayloadOut = QJsonObject();
        sensPayloadOut["session_name"] = QStringLiteral("Panel 2026-06-04");
        sensPayloadOut["score"]        = 4.5;
        RecoveryEntry sensEntry;
        sensEntry.kind        = RecoveryKind::Sensory;
        sensEntry.id          = QStringLiteral("session-7");
        sensEntry.displayName = QStringLiteral("Panel 2026-06-04");
        sensEntry.dirty       = false;
        QVERIFY(mgr.writeItem(sensEntry, sensPayloadOut));
    }

    // Task C4: crash/clean detection + move-to-_prev lifecycle.
    //  - writeItem two entries into live;
    //  - adoptPreviousSession() moves live -> _prev and leaves a fresh live dir;
    //  - hasRecoverable() is true, recoverableItems() returns both with payloads;
    //  - clear() removes both dirs and hasRecoverable() goes false.
    void detectionLifecycle()
    {
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());

        RecoveryManager mgr;
        mgr.setDirOverride(tmp.path());

        QJsonObject tpmPayload;
        QJsonObject sensPayload;
        seedTwo(mgr, tpmPayload, sensPayload);

        // Pre-adopt: nothing recoverable yet (the store is *live*, not _prev).
        QVERIFY(QFile::exists(mgr.liveDir() + "/index.json"));
        QVERIFY(!mgr.hasRecoverable());

        // ---- adopt: the prior (now orphaned) live store becomes _prev ----
        // Happy path: the promote must succeed (returns true).
        QVERIFY(mgr.adoptPreviousSession());

        // The live dir is fully gone after a clean promote (not merely
        // index-less): the first post-adopt writeItem() recreates it via mkpath.
        QVERIFY(!QDir(mgr.liveDir()).exists());
        QVERIFY(mgr.readAll(mgr.liveDir()).isEmpty());

        // _prev now holds the recoverable set with payloads intact.
        QVERIFY(mgr.hasRecoverable());
        const QVector<RecoveryEntry> rec = mgr.recoverableItems();
        QCOMPARE(rec.size(), 2);
        QCOMPARE(rec[0].kind, RecoveryKind::Tpm);
        QCOMPARE(rec[0].payload, tpmPayload);
        QCOMPARE(rec[1].kind, RecoveryKind::Sensory);
        QCOMPARE(rec[1].payload, sensPayload);

        // A fresh write after adoption lands in the new (empty) live store and
        // does not disturb the recoverable _prev set.
        RecoveryEntry fresh;
        fresh.kind        = RecoveryKind::Detailed;
        fresh.id          = QStringLiteral("ds-1");
        fresh.displayName = QStringLiteral("DS Session");
        QJsonObject freshPayload;
        freshPayload["q1"] = 3;
        QVERIFY(mgr.writeItem(fresh, freshPayload));
        QCOMPARE(mgr.readAll(mgr.liveDir()).size(), 1);
        QCOMPARE(mgr.recoverableItems().size(), 2);

        // ---- clear: both stores gone, nothing recoverable ----
        mgr.clear();
        QVERIFY(!QDir(mgr.liveDir()).exists());
        QVERIFY(!QDir(mgr.prevDir()).exists());
        QVERIFY(!mgr.hasRecoverable());
        QVERIFY(mgr.recoverableItems().isEmpty());
    }

    // adoptPreviousSession when an *old* _prev already exists: the previous
    // _prev is discarded and replaced by the current live store (we only ever
    // keep one generation back).
    void adoptReplacesStalePrev()
    {
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());

        RecoveryManager mgr;
        mgr.setDirOverride(tmp.path());

        // Stale _prev from two sessions ago. Build it the same way production
        // does: write into a live store, then adopt to push it down to _prev.
        {
            RecoveryManager stale;
            stale.setDirOverride(tmp.path());
            RecoveryEntry old;
            old.kind        = RecoveryKind::Sensory;
            old.id          = QStringLiteral("stale-session");
            old.displayName = QStringLiteral("Stale");
            QJsonObject p;
            p["marker"] = QStringLiteral("STALE");
            QVERIFY(stale.writeItem(old, p));
            QVERIFY(stale.adoptPreviousSession());   // live -> _prev (now _prev holds "stale")
        }
        QVERIFY(QFile::exists(mgr.prevDir() + "/index.json"));

        // Current session's live store: one Tpm item.
        RecoveryEntry cur;
        cur.kind        = RecoveryKind::Tpm;
        cur.id          = QStringLiteral("C:/cur.xlsx");
        cur.displayName = QStringLiteral("cur.xlsx");
        QJsonObject curP = fileResultToJson(makeFile());
        QVERIFY(mgr.writeItem(cur, curP));

        // Adopt: stale _prev is discarded, current live becomes the only _prev.
        QVERIFY(mgr.adoptPreviousSession());

        const QVector<RecoveryEntry> rec = mgr.recoverableItems();
        QCOMPARE(rec.size(), 1);
        QCOMPARE(rec[0].kind, RecoveryKind::Tpm);
        QCOMPARE(rec[0].id, QStringLiteral("C:/cur.xlsx"));
        // The stale marker must be gone.
        QVERIFY(!rec[0].payload.contains(QStringLiteral("marker")));
    }

    // adoptPreviousSession with no prior live store: a no-op that leaves nothing
    // recoverable (a clean first run, or a run after clear()).
    void adoptWithNoLiveStore()
    {
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());

        RecoveryManager mgr;
        mgr.setDirOverride(tmp.path());

        QVERIFY(!QDir(mgr.liveDir()).exists());
        QVERIFY(mgr.adoptPreviousSession());   // nothing to promote, but a clean success
        QVERIFY(!mgr.hasRecoverable());
        QVERIFY(mgr.recoverableItems().isEmpty());
    }

    // hasRecoverable() must be false when _prev has an index but it lists no
    // entries (an empty store should not trigger a recovery prompt).
    void hasRecoverableFalseForEmptyPrev()
    {
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());

        RecoveryManager mgr;
        mgr.setDirOverride(tmp.path());

        // Empty live store (index.json exists, but no items) -> adopt -> _prev.
        // Force an empty index by writing then removing the only item.
        RecoveryEntry e;
        e.kind = RecoveryKind::Tpm;
        e.id   = QStringLiteral("temp");
        QJsonObject p;
        p["x"] = 1;
        QVERIFY(mgr.writeItem(e, p));
        QVERIFY(mgr.removeItem(RecoveryKind::Tpm, QStringLiteral("temp")));
        QVERIFY(QFile::exists(mgr.liveDir() + "/index.json"));

        QVERIFY(mgr.adoptPreviousSession());
        QVERIFY(QFile::exists(mgr.prevDir() + "/index.json"));
        QVERIFY(!mgr.hasRecoverable());        // index present but empty
        QVERIFY(mgr.recoverableItems().isEmpty());
    }

    // Robustness (folded in from the C3 review, since recoverableItems() reads a
    // possibly-damaged _prev store): a blob file deleted out from under the index
    // still yields its entry, with an empty payload; the surviving entry is intact.
    void readAllSurvivesDeletedBlob()
    {
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());

        RecoveryManager mgr;
        mgr.setDirOverride(tmp.path());

        QJsonObject tpmPayload;
        QJsonObject sensPayload;
        seedTwo(mgr, tpmPayload, sensPayload);

        // Delete the Sensory blob from disk, leaving its index entry behind.
        QVector<RecoveryEntry> before = mgr.readAll(mgr.liveDir());
        QCOMPARE(before.size(), 2);
        const QString sensBlob = mgr.liveDir() + "/" + before[1].blobFile;
        QVERIFY(QFile::exists(sensBlob));
        QVERIFY(QFile::remove(sensBlob));

        const QVector<RecoveryEntry> after = mgr.readAll(mgr.liveDir());
        QCOMPARE(after.size(), 2);                       // both entries surface
        QCOMPARE(after[0].kind, RecoveryKind::Tpm);
        QCOMPARE(after[0].payload, tpmPayload);          // survivor intact
        QCOMPARE(after[1].kind, RecoveryKind::Sensory);
        QVERIFY(after[1].payload.isEmpty());             // deleted blob -> empty
    }

    // Robustness: a corrupt (non-JSON) index.json yields an empty vector, not a
    // crash. (readAll over a damaged _prev must degrade gracefully.)
    void readAllSurvivesCorruptIndex()
    {
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());

        RecoveryManager mgr;
        mgr.setDirOverride(tmp.path());

        QJsonObject a;
        QJsonObject b;
        seedTwo(mgr, a, b);

        // Stomp index.json with garbage, including embedded NUL/control bytes
        // (use an explicit QByteArray so the NUL doesn't truncate a C-string).
        const QString idxPath = mgr.liveDir() + "/index.json";
        QByteArray garbage("this is not json {{{ ]]] ");
        garbage.append('\0');
        garbage.append('\x01');
        garbage.append(" garbage");
        QFile idx(idxPath);
        QVERIFY(idx.open(QIODevice::WriteOnly | QIODevice::Truncate));
        QVERIFY(idx.write(garbage) == garbage.size());
        idx.close();

        const QVector<RecoveryEntry> after = mgr.readAll(mgr.liveDir());
        QVERIFY(after.isEmpty());
    }

    // Robustness: removeItem deletes the blob from disk and drops the entry from
    // a subsequent readAll.
    void removeItemDropsBlobAndEntry()
    {
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());

        RecoveryManager mgr;
        mgr.setDirOverride(tmp.path());

        QJsonObject tpmPayload;
        QJsonObject sensPayload;
        seedTwo(mgr, tpmPayload, sensPayload);

        const QVector<RecoveryEntry> before = mgr.readAll(mgr.liveDir());
        QCOMPARE(before.size(), 2);
        const QString tpmBlob = mgr.liveDir() + "/" + before[0].blobFile;
        QVERIFY(QFile::exists(tpmBlob));

        QVERIFY(mgr.removeItem(RecoveryKind::Tpm,
                               QStringLiteral("C:/data/Acme Device.xlsx")));

        QVERIFY(!QFile::exists(tpmBlob));                // blob gone from disk
        const QVector<RecoveryEntry> after = mgr.readAll(mgr.liveDir());
        QCOMPARE(after.size(), 1);                       // entry dropped
        QCOMPARE(after[0].kind, RecoveryKind::Sensory);
        QCOMPARE(after[0].payload, sensPayload);
    }
};

QTEST_MAIN(TstRecoveryManager)
#include "tst_recoverymanager.moc"

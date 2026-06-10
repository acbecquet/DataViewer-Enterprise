#include <QtTest>
#include <QJsonArray>
#include <QJsonObject>
#include "pipeline/SensoryData.h"
#include "ui/TesterRound.h"

using namespace DVE;

class TstSensoryDataPlaceholder : public QObject
{
    Q_OBJECT
private slots:
    void defaultNewSessionIsPlaceholder();
    void typedNameMakesItNonPlaceholder();
    void typedTesterMakesItNonPlaceholder();
    void typedSampleNameMakesItNonPlaceholder();
    void renamedSessionIsNonPlaceholder();
    void persistedSessionNeverPlaceholder();
    void typedCommentMakesItNonPlaceholder();
    void voltageMakesItNonPlaceholder();
    void resistanceMakesItNonPlaceholder();
    void heatingTechMakesItNonPlaceholder();
    void adjustedScoreMakesItNonPlaceholder();
    void controlFieldMakesItNonPlaceholder();
    void imageAttachedMakesItNonPlaceholder();
    void primaryDifferencesMakesItNonPlaceholder();
    void multiSampleSecondHasContent();
    void typedAssessorMakesItNonPlaceholder();
    void typedMediaMakesItNonPlaceholder();
    void changedPowerTypeMakesItNonPlaceholder();
    void changedPuffLengthMakesItNonPlaceholder();

    // Canonical JSON helpers (sensorySessionToJson / sensorySessionFromJson).
    // The three persistence paths (Postgres JSONB, offline SQLite,
    // user-facing .json export) all route through these, so a regression
    // here is a regression in all three.
    void jsonRoundTripPreservesAllFields();
    void deserializeOldBlobMissingNewKeysGetsDefaults();

    // DATAVIEWER-4: DB-authoritative score-merge helper. A whole-session
    // save (or export) must never clobber per-cell scores LiveSync already
    // wrote to the DB; everything else stays in-memory-authoritative.
    void mergeSensory_dbScoreWinsOverInMemoryDefault();
    void mergeSensory_nonScoreKeysComeFromInMemory();
    void mergeSensory_newInMemorySampleKeepsItsScores();
    void mergeSensory_missingDbScoreKeyLeavesInMemoryValue();

    // v2.5.0 Task 3 (RC2): dirty-aware merge. The DATAVIEWER-4 merge was
    // UNCONDITIONALLY DB-authoritative for scores, which reverted a local
    // edit on any session whose LiveSync stream never ran (id<=0 or broken
    // sync). The dirty-cell overload keeps the in-memory value for cells the
    // user touched this run, while untouched cells stay DB-authoritative.
    void merge_keepsDbForUntouchedCells();
    void merge_keepsMemoryForDirtyCells();

    // v2.5.0 Task 3 (RC2 review, CRITICAL 1): dirty-cell paths embed the sample
    // index at edit time, so removing a sample must remap them or later samples'
    // edits point at the wrong row and get reverted. remapDirtyCellsAfterSample-
    // Removal() is the shared pure helper both panels call on removal.
    void remap_dropsRemovedSampleAndShiftsLater();
    void remap_removingLastSampleDropsOnlyIt();
    void remap_earlierAndNonSamplePathsUntouched();

    // v2.5.0 Task 3 (RC2 review, CRITICAL 2): the dirty-set adoption rule
    // syncSavedSessionState() applies. A previously-persisted session keeps id>0
    // even when this tick's save FAILED, so the old unconditional clear stripped
    // a failed session's protection and the retry reverted the edit. The caller
    // clears the LOCAL copy only on Success; the panel adopts that copy.
    void adopt_failedSaveKeepsDirtySet();
    void adopt_successfulSaveClearsDirtySet();
    void adopt_neverPersistedKeepsPanelSet();

    // DATAVIEWER-4 (Task 6): the export contract that SensoryPanel's
    // dbAuthoritativeSessions() relies on — per-metric scores come from the
    // DB blob while every other field (session + sample metadata) stays
    // in-memory-authoritative. Pins the JSON-layer behaviour the panel
    // helper overlays back onto the in-memory struct before any export.
    void export_usesDbScoresWithInMemoryMetadata();

    // DATAVIEWER-8 (Task 2): pure savability predicate. A session needs a
    // non-empty test title AND a non-empty tester (round stripped) for a valid
    // DB natural key; later tasks gate interactive/background saves on this.
    void isSensorySavable_requiresTitleAndTester();
};

static QJsonObject oneSampleBlob(const QString& sampleName, double score,
                                 const QString& comments = QString())
{
    QJsonObject sample;
    sample["name"]     = sampleName;
    sample["comments"] = comments;
    for (const QString& m : DVE::kSensoryMetrics) sample[m] = score;
    QJsonArray samples; samples.append(sample);
    QJsonObject root;
    root["session_name"] = "S";
    root["samples"]      = samples;
    return root;
}

void TstSensoryDataPlaceholder::defaultNewSessionIsPlaceholder()
{
    // Mirrors what SensoryPanel::newSession() produces: default
    // name, today's date+timestamp, plus one empty SampleCard.
    SensorySession s;
    s.sessionName = "New Session";
    s.date        = "2026-05-15";
    s.timestamp   = "2026-05-15T00:00:00Z";
    s.samples.append(SensorySample{});  // empty default card
    QVERIFY(isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedNameMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.testTitle = "Strawberry vs Mint";  // user typed something
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedTesterMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.testerName = "Alice";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedSampleNameMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.name = "DeviceA";
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::renamedSessionIsNonPlaceholder()
{
    // Once the user renames the session, it's no longer a placeholder
    // regardless of content emptiness.
    SensorySession s;
    s.sessionName = "TPM run 2026-05-15";
    s.samples.append(SensorySample{});
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::persistedSessionNeverPlaceholder()
{
    // Issue 2: once a session has an id, the gate must let it through
    // even if the user blanks out the content -- otherwise NOTIFY would
    // silently roll back the blank state.
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.id = 42;
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedCommentMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.comments = "tasted like burnt rubber";
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::voltageMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.voltage = 3.7;
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::resistanceMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.resistance = 1.2;
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::heatingTechMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.heatingTechnology = "Ceramic Coil";
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::adjustedScoreMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.scores["Burnt Taste"] = 7.0;
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::controlFieldMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.control = "DeviceX baseline";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::imageAttachedMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.imagePaths.append("C:/scans/form.png");
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::primaryDifferencesMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.primaryDifferences = "flavor density";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::multiSampleSecondHasContent()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});      // empty
    SensorySample second;
    second.name = "DeviceB";
    s.samples.append(second);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedAssessorMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.assessorName = "Reviewer 3";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedMediaMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.media = "Tobacco";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::changedPowerTypeMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.powerType = "Variable Power";  // changed from default
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::changedPuffLengthMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.puffLengthSec = 5.0;  // changed from default
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::jsonRoundTripPreservesAllFields()
{
    SensorySession orig;
    orig.sessionName        = "RT";
    orig.testTitle          = "Mint vs Berry";
    orig.assessorName       = "Reviewer 3";
    orig.testerName         = "Alice";
    orig.media              = "Tobacco";
    orig.date               = "2026-05-16";
    orig.timestamp          = "2026-05-16T00:00:00Z";
    orig.control            = "Baseline";
    orig.isBlind            = true;
    orig.primaryDifferences = "flavor density";
    orig.puffLength         = "3s";
    orig.id                 = 7;
    orig.version            = 2;
    orig.imagePaths         << "C:/scans/a.png" << "C:/scans/b.png";

    SensorySample s1;
    s1.name              = "DeviceA";
    s1.voltage           = 3.7;
    s1.resistance        = 1.2;
    s1.power             = 11.4;
    s1.heatingTechnology = "Ceramic Coil";
    s1.comments          = "tasted clean";
    s1.powerType         = "Variable Voltage";
    s1.puffLengthSec     = 2.5;
    s1.scores["Burnt Taste"]    = 7.0;
    s1.scores["Vapor Volume"]   = 8.0;
    s1.scores["Overall Flavor"] = 6.5;
    s1.scores["Smoothness"]     = 7.5;
    s1.scores["Overall Liking"] = 7.0;
    orig.samples.append(s1);

    const QJsonObject json = sensorySessionToJson(orig);
    const SensorySession decoded = sensorySessionFromJson(json);

    QCOMPARE(decoded.sessionName,        orig.sessionName);
    QCOMPARE(decoded.testTitle,          orig.testTitle);
    QCOMPARE(decoded.assessorName,       orig.assessorName);
    QCOMPARE(decoded.testerName,         orig.testerName);
    QCOMPARE(decoded.media,              orig.media);
    QCOMPARE(decoded.date,               orig.date);
    QCOMPARE(decoded.timestamp,          orig.timestamp);
    QCOMPARE(decoded.control,            orig.control);
    QCOMPARE(decoded.isBlind,            orig.isBlind);
    QCOMPARE(decoded.primaryDifferences, orig.primaryDifferences);
    QCOMPARE(decoded.samples.size(),     orig.samples.size());

    const SensorySample& d = decoded.samples[0];
    QCOMPARE(d.name,              s1.name);
    QCOMPARE(d.voltage,           s1.voltage);
    QCOMPARE(d.resistance,        s1.resistance);
    QCOMPARE(d.heatingTechnology, s1.heatingTechnology);
    QCOMPARE(d.comments,          s1.comments);
    QCOMPARE(d.powerType,         s1.powerType);
    QCOMPARE(d.puffLengthSec,     s1.puffLengthSec);
    QCOMPARE(d.scores.value("Burnt Taste"),    s1.scores["Burnt Taste"]);
    QCOMPARE(d.scores.value("Vapor Volume"),   s1.scores["Vapor Volume"]);
    QCOMPARE(d.scores.value("Overall Flavor"), s1.scores["Overall Flavor"]);
    QCOMPARE(d.scores.value("Smoothness"),     s1.scores["Smoothness"]);
    QCOMPARE(d.scores.value("Overall Liking"), s1.scores["Overall Liking"]);
}

void TstSensoryDataPlaceholder::deserializeOldBlobMissingNewKeysGetsDefaults()
{
    // Simulate a JSON blob written by pre-#7 code. Loading must give
    // back the struct defaults for powerType and puffLengthSec.
    QJsonObject obj;
    obj["session_name"] = "OldRow";
    QJsonArray samples;
    QJsonObject s;
    s["name"]    = "OldDevice";
    s["voltage"] = 3.5;
    samples.append(s);
    obj["samples"] = samples;

    const SensorySession decoded = sensorySessionFromJson(obj);
    QCOMPARE(decoded.samples.size(), 1);
    QCOMPARE(decoded.samples[0].powerType,     QStringLiteral("Constant Voltage"));
    QCOMPARE(decoded.samples[0].puffLengthSec, 3.0);
}

void TstSensoryDataPlaceholder::mergeSensory_dbScoreWinsOverInMemoryDefault()
{
    QJsonObject mem = oneSampleBlob("A", 5.0);
    QJsonObject db  = oneSampleBlob("A", 8.0);
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    const QJsonObject s0 = merged["samples"].toArray()[0].toObject();
    for (const QString& m : DVE::kSensoryMetrics)
        QCOMPARE(s0[m].toDouble(), 8.0);
}

void TstSensoryDataPlaceholder::mergeSensory_nonScoreKeysComeFromInMemory()
{
    QJsonObject mem = oneSampleBlob("NewName", 5.0, "new comment");
    mem["media"] = "MediaX";
    QJsonObject db  = oneSampleBlob("OldName", 8.0, "old comment");
    db["media"] = "MediaOld";
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    QCOMPARE(merged["media"].toString(), QString("MediaX"));
    const QJsonObject s0 = merged["samples"].toArray()[0].toObject();
    QCOMPARE(s0["name"].toString(), QString("NewName"));
    QCOMPARE(s0["comments"].toString(), QString("new comment"));
    QCOMPARE(s0[DVE::kSensoryMetrics.first()].toDouble(), 8.0);
}

void TstSensoryDataPlaceholder::mergeSensory_newInMemorySampleKeepsItsScores()
{
    QJsonObject mem = oneSampleBlob("A", 5.0);
    QJsonArray ms = mem["samples"].toArray();
    QJsonObject s1; s1["name"] = "B";
    for (const QString& m : DVE::kSensoryMetrics) s1[m] = 7.0;
    ms.append(s1); mem["samples"] = ms;
    QJsonObject db = oneSampleBlob("A", 8.0);
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    const QJsonArray out = merged["samples"].toArray();
    QCOMPARE(out.size(), 2);
    QCOMPARE(out[0].toObject()[DVE::kSensoryMetrics.first()].toDouble(), 8.0);
    QCOMPARE(out[1].toObject()[DVE::kSensoryMetrics.first()].toDouble(), 7.0);
}

void TstSensoryDataPlaceholder::mergeSensory_missingDbScoreKeyLeavesInMemoryValue()
{
    QJsonObject mem = oneSampleBlob("A", 5.0);
    QJsonObject db  = oneSampleBlob("A", 8.0);
    QJsonObject s0 = db["samples"].toArray()[0].toObject();
    s0.remove("Smoothness");                 // real metric name from kSensoryMetrics
    QJsonArray dbs; dbs.append(s0); db["samples"] = dbs;
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    const QJsonObject m0 = merged["samples"].toArray()[0].toObject();
    QCOMPARE(m0["Smoothness"].toDouble(), 5.0);     // removed metric -> in-memory 5.0
    QCOMPARE(m0["Burnt Taste"].toDouble(), 8.0);    // present metric -> DB 8.0
}

void TstSensoryDataPlaceholder::merge_keepsDbForUntouchedCells()
{
    // Regression-lock: empty dirty set behaves exactly like the legacy
    // unconditional merge -- every score key is DB-authoritative.
    QJsonObject mem = oneSampleBlob("A", 5.0);
    QJsonObject db  = oneSampleBlob("A", 8.0);
    QJsonObject merged =
        DVE::mergeSensoryPreservingDbScores(mem, db, /*dirtyCells=*/{});
    const QJsonObject s0 = merged["samples"].toArray()[0].toObject();
    for (const QString& m : DVE::kSensoryMetrics)
        QCOMPARE(s0[m].toDouble(), 8.0);
}

void TstSensoryDataPlaceholder::merge_keepsMemoryForDirtyCells()
{
    // samples[1].Smoothness was edited locally this run (dirty). DB has 5.0,
    // memory has 7.5 -> the merged blob must keep 7.5 for THAT cell while
    // every other metric (and samples[0] entirely) still takes the DB value.
    QJsonObject mem = oneSampleBlob("A", 1.0);   // sample 0: all metrics 1.0
    QJsonArray ms = mem["samples"].toArray();
    QJsonObject s1; s1["name"] = "B";
    for (const QString& m : DVE::kSensoryMetrics) s1[m] = 2.0;
    s1["Smoothness"] = 7.5;                       // the locally-edited cell
    ms.append(s1); mem["samples"] = ms;

    QJsonObject db = oneSampleBlob("A", 8.0);     // db sample 0: all 8.0
    QJsonArray dbs = db["samples"].toArray();
    QJsonObject d1; d1["name"] = "B";
    for (const QString& m : DVE::kSensoryMetrics) d1[m] = 5.0;  // incl Smoothness 5.0
    dbs.append(d1); db["samples"] = dbs;

    const QSet<QString> dirty = { QStringLiteral("samples[1].Smoothness") };
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db, dirty);
    const QJsonArray out = merged["samples"].toArray();

    // sample 0 untouched -> DB-authoritative throughout
    const QJsonObject o0 = out[0].toObject();
    for (const QString& m : DVE::kSensoryMetrics)
        QCOMPARE(o0[m].toDouble(), 8.0);

    // sample 1: only Smoothness is dirty -> memory wins (7.5); every other
    // metric still takes the DB value (5.0).
    const QJsonObject o1 = out[1].toObject();
    QCOMPARE(o1["Smoothness"].toDouble(), 7.5);
    for (const QString& m : DVE::kSensoryMetrics) {
        if (m == QLatin1String("Smoothness")) continue;
        QCOMPARE(o1[m].toDouble(), 5.0);
    }
}

void TstSensoryDataPlaceholder::remap_dropsRemovedSampleAndShiftsLater()
{
    // Remove sample index 0. The removed sample's paths vanish; every later
    // path's index drops by one (samples[2].X -> samples[1].X, samples[1].Y ->
    // samples[0].Y).
    const QSet<QString> dirty = {
        QStringLiteral("samples[0].Smoothness"),     // removed -> dropped
        QStringLiteral("samples[1].Burnt Taste"),    // -> samples[0].Burnt Taste
        QStringLiteral("samples[2].Vapor Volume"),   // -> samples[1].Vapor Volume
    };
    const QSet<QString> out =
        DVE::remapDirtyCellsAfterSampleRemoval(dirty, /*removedIdx=*/0);

    const QSet<QString> expected = {
        QStringLiteral("samples[0].Burnt Taste"),
        QStringLiteral("samples[1].Vapor Volume"),
    };
    QCOMPARE(out, expected);
    QVERIFY(!out.contains(QStringLiteral("samples[0].Smoothness")));  // metric of dropped row not re-added
}

void TstSensoryDataPlaceholder::remap_removingLastSampleDropsOnlyIt()
{
    // Removing the highest index drops only that sample's paths; lower indices
    // are unchanged (nothing shifts).
    const QSet<QString> dirty = {
        QStringLiteral("samples[0].Smoothness"),
        QStringLiteral("samples[1].Overall Liking"),
        QStringLiteral("samples[2].Burnt Taste"),    // the last sample -> dropped
        QStringLiteral("samples[2].Vapor Volume"),    // ditto
    };
    const QSet<QString> out =
        DVE::remapDirtyCellsAfterSampleRemoval(dirty, /*removedIdx=*/2);

    const QSet<QString> expected = {
        QStringLiteral("samples[0].Smoothness"),
        QStringLiteral("samples[1].Overall Liking"),
    };
    QCOMPARE(out, expected);
}

void TstSensoryDataPlaceholder::remap_earlierAndNonSamplePathsUntouched()
{
    // Removing index 2: paths for earlier samples (0,1) keep their index; a
    // non-sample path passes through verbatim.
    const QSet<QString> dirty = {
        QStringLiteral("samples[0].Smoothness"),     // earlier -> unchanged
        QStringLiteral("samples[1].Burnt Taste"),    // earlier -> unchanged
        QStringLiteral("samples[3].Vapor Volume"),   // later  -> samples[2].Vapor Volume
        QStringLiteral("notASamplePath"),            // non-matching -> verbatim
    };
    const QSet<QString> out =
        DVE::remapDirtyCellsAfterSampleRemoval(dirty, /*removedIdx=*/2);

    const QSet<QString> expected = {
        QStringLiteral("samples[0].Smoothness"),
        QStringLiteral("samples[1].Burnt Taste"),
        QStringLiteral("samples[2].Vapor Volume"),
        QStringLiteral("notASamplePath"),
    };
    QCOMPARE(out, expected);
}

void TstSensoryDataPlaceholder::adopt_failedSaveKeepsDirtySet()
{
    // A previously-persisted session (savedId > 0) whose write FAILED this tick:
    // the caller did NOT clear the local copy, so savedDirty is still non-empty.
    // The panel must KEEP the protection (adopt the still-dirty set) — this is
    // the exact regression the review flagged.
    const QSet<QString> savedDirty = { QStringLiteral("samples[0].Smoothness") };
    const QSet<QString> panelDirty = { QStringLiteral("samples[0].Smoothness") };
    const QSet<QString> out =
        DVE::adoptedDirtyCellsAfterSave(/*savedId=*/7, savedDirty, panelDirty);
    QCOMPARE(out, savedDirty);
    QVERIFY(!out.isEmpty());                 // protection retained on failure
}

void TstSensoryDataPlaceholder::adopt_successfulSaveClearsDirtySet()
{
    // savedId > 0 and the caller cleared the local copy on Success (savedDirty
    // empty). The panel adopts the empty set -> dirty cleared.
    const QSet<QString> savedDirty;          // caller cleared it on Success
    const QSet<QString> panelDirty = { QStringLiteral("samples[0].Smoothness") };
    const QSet<QString> out =
        DVE::adoptedDirtyCellsAfterSave(/*savedId=*/7, savedDirty, panelDirty);
    QVERIFY(out.isEmpty());                  // edits are in the DB now -> cleared
}

void TstSensoryDataPlaceholder::adopt_neverPersistedKeepsPanelSet()
{
    // savedId <= 0: the write never landed (placeholder / hard error). The panel
    // keeps its existing set untouched (no spurious clear, no spurious adopt).
    const QSet<QString> savedDirty;          // irrelevant when id<=0
    const QSet<QString> panelDirty = { QStringLiteral("samples[1].Burnt Taste") };
    const QSet<QString> out =
        DVE::adoptedDirtyCellsAfterSave(/*savedId=*/-1, savedDirty, panelDirty);
    QCOMPARE(out, panelDirty);               // panel's own set preserved
}

void TstSensoryDataPlaceholder::export_usesDbScoresWithInMemoryMetadata()
{
    QJsonObject mem = oneSampleBlob("A", 5.0, "memo");
    mem["media"] = "FreshMedia";
    QJsonObject db = oneSampleBlob("A", 8.0, "old");
    db["media"] = "OldMedia";
    QJsonObject exportBlob = DVE::mergeSensoryPreservingDbScores(mem, db);
    QCOMPARE(exportBlob["media"].toString(), QString("FreshMedia"));   // metadata in-memory
    const QJsonObject s0 = exportBlob["samples"].toArray()[0].toObject();
    QCOMPARE(s0["comments"].toString(), QString("memo"));              // comment in-memory
    QCOMPARE(s0[DVE::kSensoryMetrics.first()].toDouble(), 8.0);        // score from DB
}

void TstSensoryDataPlaceholder::isSensorySavable_requiresTitleAndTester()
{
    DVE::SensorySession s;
    s.testTitle = "";  s.testerName = DVE::combineTesterRound("Alice", "1");
    QVERIFY(!DVE::isSensorySessionSavable(s));          // no title
    s.testTitle = "T"; s.testerName = DVE::combineTesterRound("", "1");
    QVERIFY(!DVE::isSensorySessionSavable(s));          // no tester (round only)
    s.testTitle = "T"; s.testerName = DVE::combineTesterRound("Alice", "1");
    QVERIFY(DVE::isSensorySessionSavable(s));           // both present
    s.testTitle = "   "; s.testerName = DVE::combineTesterRound("Alice", "1");
    QVERIFY(!DVE::isSensorySessionSavable(s));          // whitespace title
}

QTEST_MAIN(TstSensoryDataPlaceholder)
#include "tst_sensorydataplaceholder.moc"

#pragma once

#include <QString>
#include <QStringList>
#include <QMap>
#include <QSet>
#include <QVector>
#include <QRectF>

namespace DVE {

// Data-entry / table order (matches the physical evaluation sheet)
static const QStringList kSensoryMetrics = {
    "Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness", "Overall Liking"
};

// Plot order — Overall Liking at the top (first axis = 12-o'clock)
static const QStringList kSensoryMetricsPlot = {
    "Overall Liking", "Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness"
};

struct SensorySample {
    QString name;
    QMap<QString, double> scores;   // metric → 1.0–9.0, default 5.0
    QString comments;

    // Per-sample device properties
    double  voltage    = 0.0;
    double  resistance = 0.0;
    double  power      = 0.0;       // calculated: V²/(R + Roffset)
    QString heatingTechnology;

    // #7: per-sample test conditions added 2026-05-15.
    QString powerType    = QStringLiteral("Constant Voltage");
    double  puffLengthSec = 3.0;
};

struct SensorySession {
    QString  sessionName;
    QString  testTitle;
    QString  assessorName;
    QString  testerName;
    QString  media;
    QString  date;
    QVector<SensorySample> samples;
    QString  timestamp;       // ISO8601
    QString  sourceImagePath; // path of scanned form image (may be empty)

    // Session-level test properties
    QString  control;            // control sample name
    bool     isBlind = false;    // blind test?
    QString  primaryDifferences; // what is being tested

    // Legacy fields kept for backward compatibility with old JSON/DB
    QString  puffLength;
    QString  burnStatus;
    QString  clogStatus;
    QString  leakStatus;
    double   resistance = 0.0;
    double   voltage    = 0.0;
    double   power      = 0.0;
    QString  heatingTechnology;

    // Images linked to this session (same pattern as TPM SampleResult)
    QStringList     imagePaths;
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;
    // C3: server-assigned image-row identities, parallel to imagePaths. -1/0
    // means "fresh"; the id-aware upsert in tryWriteSensorySession back-fills
    // these from RETURNING and uses them to UPDATE existing image rows in
    // place rather than wiping the whole session_id subtree.
    QVector<qint64> imageIds;
    QVector<int>    imageVersions;

    // Persistence anchors (added for the report-preview feature, Phase 1A Task 4)
    int     id      = -1;        // -1 if not yet persisted; populated by DB loaders
    int     version = 0;         // server-assigned row version; 0 if unknown
    QString sourceFilePath;      // the .xlsx file the session was loaded from (empty if DB-only)
    QString excelLayoutJson;     // dve_layout custom property pulled by ExcelReader
                                 // (empty until Phase 2 lands)

    // v2.1.0+: sessionName at the moment this session was loaded (or last
    // synced with the DB). When a user edits the Test Title and the new name
    // differs from this snapshot, the save flow treats it as a brand-new
    // session (forces INSERT) instead of UPDATE-in-place, so the old DB row
    // is preserved for bookkeeping and a new row lands with the new name.
    // Empty until the session has been persisted at least once.
    QString originalSessionName;

    // v2.5.0 Task 3 (RC2): cells the user edited THIS run, recorded by the
    // panel's per-cell commit handler. The dirty-aware merge keeps the
    // in-memory value for these cells instead of letting the DB blob win, so a
    // local edit is never reverted on save/export even when LiveSync never
    // streamed it (id<=0 at edit time, or a broken sync connection).
    //
    // PATH FORMAT: "samples[<idx>].<MetricKey>" where <idx> is the 0-based
    // sample index and <MetricKey> is the literal kSensoryMetrics string the
    // serializer/merge use as the flat score key (e.g. "samples[1].Smoothness").
    // This mirrors the LiveSync json_path minus its "json_path:" prefix.
    //
    // NOT serialized: sensorySessionToJson/fromJson deliberately ignore this —
    // it is a per-run UI-edit signal, not persisted state.
    QSet<QString> dirtyCells;
};

inline bool isPlaceholderSession(const SensorySession& s)
{
    // A session is a "placeholder" -- not yet meaningfully filled in --
    // when ALL of the following hold:
    //   1. It has never been persisted (id < 0). Once it has an id,
    //      subsequent saves must always go through so the user can
    //      legitimately blank a row out without it silently rolling
    //      back via NOTIFY.
    //   2. It still has the default new-session name.
    //   3. No header / session-level field has user content.
    //   4. No sample card has user content (name, comments, V/R,
    //      heating tech, image, or a non-default score).
    //
    // Without this gate, the 5-second auto-save persists the freshly-
    // created session to Postgres, which NOTIFY-broadcasts it to every
    // other client's navigator as a spurious "New Session" entry.
    if (s.id >= 0) return false;
    if (s.sessionName != QStringLiteral("New Session")) return false;

    const bool noHeaderContent = s.testerName.isEmpty()
                              && s.assessorName.isEmpty()
                              && s.media.isEmpty()
                              && s.testTitle.isEmpty()
                              && s.control.isEmpty()
                              && s.primaryDifferences.isEmpty()
                              && s.imagePaths.isEmpty();
    if (!noHeaderContent) return false;

    for (const SensorySample& sample : s.samples) {
        if (!sample.name.isEmpty())              return false;
        if (!sample.comments.isEmpty())          return false;
        if (sample.voltage    > 0.0)             return false;
        if (sample.resistance > 0.0)             return false;
        if (!sample.heatingTechnology.isEmpty()) return false;
        for (auto it = sample.scores.constBegin();
             it != sample.scores.constEnd(); ++it) {
            if (it.value() != 5.0) return false;
        }
        // #7: per-sample test conditions. Defaults MUST match the
        // SensorySample struct defaults above ("Constant Voltage" / 3.0);
        // if those change, update both sides together.
        if (sample.powerType != QStringLiteral("Constant Voltage")) return false;
        if (sample.puffLengthSec != 3.0)                            return false;
    }
    return true;
}

} // namespace DVE

#include <QJsonObject>

namespace DVE {

// Canonical JSON contract for SensorySession. All three persistence paths
// (Postgres JSONB column, offline SQLite snapshot, user-facing .json file
// export/import) route through these two helpers. Keep the field set + key
// names symmetric: any new SensorySample / SensorySession field must be
// added to BOTH functions and to the round-trip test in
// tst_sensorydataplaceholder.
QJsonObject sensorySessionToJson(const SensorySession& s);
SensorySession sensorySessionFromJson(const QJsonObject& obj);

// DATAVIEWER-4: merge an in-memory sensory blob with the current DB blob so a
// whole-session write OR an export never clobbers LiveSync-owned per-cell SCORE
// values. Scores are DB-authoritative: for every sample present in BOTH (matched
// by array index), each kSensoryMetrics score key is taken from `dbCurrent`; all
// other keys (metadata, structure, name, comments, device props) come from
// `inMemory`. Samples in `inMemory` beyond `dbCurrent`'s array keep their
// in-memory scores. Pure / no DB.
//
// v2.5.0 Task 3 (RC2): `dirtyCells` holds paths "samples[<idx>].<MetricKey>"
// (SensorySession::dirtyCells format) for cells the user edited this run. For
// any such (sample idx, metric) the DB-preserve is SKIPPED so the in-memory
// (locally-edited) value wins — fixing the revert-on-save data loss for
// sessions whose LiveSync stream never ran. An empty set reproduces the legacy
// unconditional DB-authoritative behavior exactly.
QJsonObject mergeSensoryPreservingDbScores(const QJsonObject& inMemory,
                                           const QJsonObject& dbCurrent,
                                           const QSet<QString>& dirtyCells = {});

// v2.5.0 Task 3 (RC2 review, CRITICAL 1): dirty-cell paths embed the sample
// index at edit time ("samples[<idx>].<MetricKey>"), so removing a sample
// invalidates every path that named a LATER sample. This pure helper remaps a
// dirty set across one sample removal at `removedIdx`:
//   - paths for the removed sample ("samples[removedIdx].*") are DROPPED;
//   - paths for samples after it ("samples[k].*", k > removedIdx) have their
//     index decremented by one to track the array shift;
//   - paths for samples before it (k < removedIdx) and any non-matching path
//     pass through unchanged.
// Shared by SensoryPanel (card removal) and DetailedSensoryPanel (sample
// removal) because both use the identical path format. Sample ADD only ever
// appends, which shifts no existing index, so no add-side remap is needed.
// Pure / no DB / no Qt widgets.
QSet<QString> remapDirtyCellsAfterSampleRemoval(const QSet<QString>& dirty,
                                                int removedIdx);

// v2.5.0 Task 3 (RC2 review, CRITICAL 2): the per-session dirty-set adoption
// rule that syncSavedSessionState() applies after a save loop. Encodes the fix
// for the "failed save still clears the dirty set" data loss:
//
//   * The caller (MainWindow's save loop) runs synchronously on the UI thread,
//     so no edits interleave. It clears the LOCAL copy's dirtyCells ONLY when
//     that session's WriteResult == Success, leaving a FAILED session's local
//     copy still carrying the set (a previously-persisted session keeps id>0
//     even on failure because the write wrapper back-fills only on Success).
//   * The panel then ADOPTS the local copy's set whenever savedId > 0.
//
// Net per session: savedId<=0 (never landed) -> keep the panel's existing set
// untouched; savedId>0 + Success -> caller cleared savedDirty -> panel cleared;
// savedId>0 + failure -> savedDirty still set -> panel keeps its protection so
// the retry still treats the local edits as authoritative (never DB-over-memory
// reverts them). Pure; mirrors the inline adopt in both panels'
// syncSavedSessionState().
QSet<QString> adoptedDirtyCellsAfterSave(int savedId,
                                         const QSet<QString>& savedDirty,
                                         const QSet<QString>& panelDirty);

// DATAVIEWER-8: True iff the session has the non-empty test title AND tester
// needed for a valid natural key (session_name, tester_name, date). Round is
// stripped before the tester check (tester+round are folded into testerName).
// Whitespace-only counts as empty. Pure. Gates saves (block interactive /
// skip background auto-save / refuse close of unnamed sessions).
bool isSensorySessionSavable(const SensorySession& s);

} // namespace DVE

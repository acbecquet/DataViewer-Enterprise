// Two harnesses over one Postgres fixture: the TPM v3 Phase 3b wide -> long
// data migration, and the Phase 3c END-TO-END GATE.
//
// Under test (3b): deploy/postgres/migrations/2026-07-31-v3-long-format.sql -
//   dve_migrate_to_long_format(), and the data_rows_v / samples_v compat views
//   it has to stay value-exact against.
// Under test (3c): the whole open-metric chain, end to end - see THE GATE
//   below.
// Plans: docs/superpowers/plans/2026-07-31-tpm-v3-phase3b-schema-migration-plan.md, Task 5
//        docs/superpowers/plans/2026-07-31-tpm-v3-phase3c-open-metrics-plan.md, Task 5
//
// WHY THIS EXISTS
//   The migration was first verified by hand, through throwaway psql scripts.
//   That proves it worked once. This suite makes the same proof permanent,
//   automated and re-runnable, so the migration stays honest through 3c and 3d
//   and the eventual supervised production run.
//
// WHAT IT PROVES
//   1. Value parity  - data_rows_v is row-for-row identical to data_rows on all
//      13 value columns, and samples_v to samples on all 22 header columns,
//      compared with IS NOT DISTINCT FROM so NULLs compare equal.
//   2. Per-metric checksums - count / sum(::numeric) / min / max plus an ordered
//      md5 of the text form of every value, on both sides, for all 35 keys.
//   3. The sparse rule (index D2) in BOTH directions: a source NULL produces no
//      measurement row at all, and a source 0.0 produces a row holding 0.
//   4. hasPerRowRegime preservation - a sheet whose puffing_regime is entirely
//      NULL yields zero regime measurements, so `puffing_regime IS NOT NULL`
//      still derives false.
//   5. Bit-exact doubles - 0.1 + 0.2 survives as the same 8 bytes.
//   6. Idempotence - a second run inserts nothing and churns no versions.
//   7. All three pre-flight aborts fire with a message naming the offending row,
//      and a missing metric_defs seed row raises a named exception instead of
//      silently skipping the column.
//   8. The H22 / D9 orphan tripwire - no measurement sits at a (sample_id,
//      sort_order) that data_rows does not have. Vacuous in 3b by design; it is
//      here before 3c's save path can break the contract it states.
//   9. Column arity - the wide tables' value-column sets in the live catalog are
//      exactly the lists this harness and the migration pin, so a column added
//      by a later ALTER cannot be silently skipped by both.
//
// DESIGN NOTES
//   - Every value comparison happens INSIDE Postgres (md5 over ::text,
//     float8send for bit equality, IS NOT DISTINCT FROM). Only counts and
//     booleans cross the wire. Qt's QPSQL driver renders bound doubles through
//     QSqlDriver::formatValue, which is not guaranteed to emit the 17
//     significant digits a double needs to round-trip; seeding therefore builds
//     SQL literals with QString::number(v, 'g', 17) and reading back never
//     converts a double to a C++ double at all.
//   - The four abort tests each run inside a transaction that is ALWAYS rolled
//     back, so the deliberate violation (and the deliberately deleted
//     metric_defs row) never survives the test and test order does not matter.
//
// THE GATE (v3 Phase 3c Task 5) - the reason Phase 3c exists
//   tests/data/manifest_demo.xlsx has a "Custom Coil Test" sheet whose
//   very-hidden _dve_schema manifest declares a custom `coil_temp` column.
//   Phase 2 has parsed it into DataRow::extra since 2c, where it has been
//   invisible and unpersisted. Three sub-phases of work exist to make ONE
//   sentence true, and the last three slots of this file are that sentence:
//
//     1. process the workbook, save it to Postgres, DISCARD the in-memory
//        result, reload from the database, and find `coil_temp` present,
//        correct and correctly typed on every data row;
//     2. do it AGAIN - reload, save, reload - because a reload that returns no
//        extras followed by a save whose orphan prune then deletes them is the
//        exact sequence that silently lost data before 3c (hazard H1);
//     3. the same round trip through the offline SQLite snapshot instead of
//        Postgres.
//
//   Those three slots drive the REAL chain (ExcelReader's openpyxl subprocess
//   -> DataProcessor -> SchemaResolver -> LegacyAdapter -> DatabaseManager /
//   DatabaseOps / OfflineSnapshot), mock nothing, and refuse to pass vacuously:
//   every phase re-asserts the fixture's known shape and the exact NUMBER of
//   values compared before any comparison is allowed to count as a pass. Every
//   failure message names the phase, because "coil_temp is missing" is useless
//   without knowing which link dropped it.
//
// Skips cleanly when DVE_TEST_PG_CONN is unset. Hard-fails, with the
// remediation command, if the container is missing the long-format objects.
// The gate slots skip cleanly when python + openpyxl are unavailable (the
// pipeline suites' established idiom) and hard-fail if the fixture is missing.

#include <QtTest/QtTest>

#include <QCoreApplication>
#include <QDateTime>
#include <QFileInfo>
#include <QMetaType>
#include <QScopeGuard>
#include <QSqlDatabase>
#include <QSqlError>
#include <QSqlQuery>
#include <QString>
#include <QStringList>
#include <QTemporaryDir>
#include <QVariant>

#include "MetricDefCache.h"   // v3 Phase 3c, plan Task 1

// v3 Phase 3c Task 5 - the end-to-end gate drives the production chain.
#include "ConfigLoader.h"     // DbConfig
#include "DataProcessor.h"
#include "DatabaseManager.h"
#include "IdentityManager.h"
#include "OfflineSnapshot.h"
#include "PostgresConnection.h"
#include "ReportData.h"
#include "TestHelpers.h"      // testDataFile / fuzzyEqual

namespace {

const char* const kConn = "tst_v3longformat";

// ---------------------------------------------------------------------------
// Connection
// ---------------------------------------------------------------------------

struct PgCfg {
    QString host     = QStringLiteral("127.0.0.1");
    int     port     = 5432;
    QString database;
    QString user;
    QString password;
};

PgCfg pgConfig()
{
    PgCfg c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
    const QStringList parts = env.split(' ', Qt::SkipEmptyParts);
    for (const QString& part : parts) {
        const QStringList kv = part.split('=');
        if (kv.size() != 2) continue;
        const QString k = kv.at(0);
        const QString v = kv.at(1);
        if      (k == QLatin1String("host"))     c.host     = v;
        else if (k == QLatin1String("port"))     c.port     = v.toInt();
        else if (k == QLatin1String("dbname"))   c.database = v;
        else if (k == QLatin1String("user"))     c.user     = v;
        else if (k == QLatin1String("password")) c.password = v;
    }
    return c;
}

// The same environment variable as a DVE::DbConfig, for the production
// components the 3c gate drives (DatabaseManager, PostgresConnection). The two
// structs are kept separate rather than merged because the 3b half of this file
// deliberately links no application code beyond MetricDefCache and must stay
// buildable if the gate below is ever split out again.
DVE::DbConfig pgDbConfig()
{
    const PgCfg c = pgConfig();
    DVE::DbConfig d;
    d.host     = c.host;
    d.port     = c.port;
    d.database = c.database;
    d.user     = c.user;
    d.password = c.password;
    return d;
}

// ---------------------------------------------------------------------------
// v3 Phase 3c Task 5 - the manifest demo fixture, pinned
//
// These are the fixture's KNOWN shape, transcribed from tests/generate_fixtures.py
// (gen_manifest_demo) and already asserted independently by
// tst_v3inference::manifestDemoEndToEnd. Restating them here is what makes the
// gate incapable of passing vacuously: a reload that returns an empty tree, or
// rows carrying empty `extra` maps, fails a count assertion instead of sailing
// through a loop that never executes.
// ---------------------------------------------------------------------------

const char* const kDemoSheetName    = "Custom Coil Test";
const char* const kCoilTempKey      = "coil_temp";
constexpr int     kDemoSamples      = 2;
constexpr int     kDemoRowsPerSample = 6;
// Sheet cell = base + rowIndex, per block. Sample 2 starts at 210 precisely so
// a block-1/block-2 mix-up is visible rather than silently self-consistent.
const double      kDemoCoilBase[kDemoSamples] = { 200.0, 210.0 };
// The recomputed TPM every row of the demo carries (the sheet's own TPM column
// holds a bogus 999 that the Derived recompute overwrites).
constexpr double  kDemoTpm          = 3.5;
const char* const kDemoRegime       = "60mL/3s/30s";

// The gate saves the demo under a TEST-OWNED file_path rather than the
// fixture's real one: `files.file_path` is the business key the save path
// upserts on, so owning it makes the residue cleanup below exact and keeps the
// gate from colliding with a corpus run that loaded the same workbook.
const char* const kE2eFilePath = "/tmp/tst_v3longformat_e2e_manifest_demo.xlsx";

// HAZARD H24. metric_defs is deliberately never wiped (see wipe()), it outlives
// this process, and tst_databasemanager::v3LongFormat_ensureSchemaMatchesRegistry
// asserts an EXACT row count against the compiled registry. The gate below
// auto-registers `coil_temp` on first save - it is in NEITHER the registry nor
// the migration seed, which is the whole point - so leaking it would turn a
// DIFFERENT suite red for a reason that has nothing to do with that suite.
// These are the keys this file's scenarios invent, deleted before each gate slot
// AND on teardown. Same pattern as tst_saveintegrity_e2e::openMetricKeys() and
// tst_databasemanager::v3OpenMetricKeys().
//
// The gate does not merely trust this list: it diffs metric_defs across the save
// and asserts the difference is EXACTLY this, so a fixture that starts inventing
// a second key fails here instead of leaking silently.
QStringList e2eInventedKeys()
{
    return QStringList{ QStringLiteral("metric|coil_temp") };
}

// ---------------------------------------------------------------------------
// The wide surface, in column order. These lists are the harness's own copy of
// what the migration claims to cover; they are deliberately NOT read out of
// metric_defs, so a column silently dropped from the migration's pinned array
// shows up here as a missing measurement rather than as a shrunken comparison.
// ---------------------------------------------------------------------------

const QStringList kMetricCols = {
    "puffs", "before_weight", "after_weight", "draw_pressure", "resistance",
    "smell", "clog", "notes", "tpm", "tpm_power_density", "variation_tpm",
    "oil_consumed", "puffing_regime"
};
const QStringList kTextMetricCols = { "smell", "clog", "notes", "puffing_regime" };

const QStringList kHeaderCols = {
    "sample_name", "sample_id", "date", "tester", "media", "viscosity",
    "resistance", "voltage", "power", "heating_technology", "puffing_regime",
    "initial_oil_mass", "average_tpm", "stddev_tpm", "avg_power_density",
    "efficiency_percent", "total_oil_consumed", "total_puffs",
    "normalized_tpm", "burn_status", "clog_status", "leak_status"
};
const QStringList kTextHeaderCols = {
    "sample_name", "sample_id", "date", "tester", "media", "heating_technology",
    "puffing_regime", "burn_status", "clog_status", "leak_status"
};

// ---------------------------------------------------------------------------
// SQL literal helpers. 17 significant digits is the shortest precision at which
// every IEEE-754 double round-trips through decimal text.
// ---------------------------------------------------------------------------

QString sqlNum(double v)
{
    return QString::number(v, 'g', 17);
}

QString sqlNumOrNull(bool present, double v)
{
    return present ? sqlNum(v) : QStringLiteral("NULL");
}

// A null QString means SQL NULL; an empty-but-non-null one means ''.
QString sqlTextOrNull(const QString& s)
{
    if (s.isNull()) return QStringLiteral("NULL");
    QString escaped = s;
    escaped.replace(QLatin1Char('\''), QLatin1String("''"));
    return QLatin1Char('\'') + escaped + QLatin1Char('\'');
}

QString sqlKeyList(const QStringList& keys)
{
    QStringList quoted;
    quoted.reserve(keys.size());
    for (const QString& k : keys) quoted << sqlTextOrNull(k);
    return quoted.join(QLatin1String(", "));
}

// The double every float-exactness argument is made of: 0.1 + 0.2 is the
// canonical value whose shortest round-trip form needs all 17 digits.
const double kTrickyDouble = 0.1 + 0.2;
const char* const kTrickyDoubleText = "0.30000000000000004";

} // namespace

#define DB_EXEC(query, sqlText)                                                    \
    do {                                                                           \
        const QString _sql = (sqlText);                                            \
        if (!(query).exec(_sql)) {                                                 \
            QFAIL(qPrintable(QStringLiteral("SQL failed: %1\n%2")                  \
                                 .arg((query).lastError().text(),                  \
                                      _sql.left(600))));                           \
        }                                                                          \
    } while (false)

class TstV3LongFormat : public QObject
{
    Q_OBJECT

private slots:
    void initTestCase();
    void cleanupTestCase();

    // --- the migrated dataset -------------------------------------------------
    void migration_returnsExpectedCounts();
    void dataRowsView_matchesDataRowsOnAllValueColumns();
    void samplesView_matchesSamplesOnAllHeaderColumns();
    void perMetricChecksums_match();
    void perHeaderFieldChecksums_match();
    void sparseRule_sourceNullProducesNoMeasurementRow();
    void sparseRule_sourceZeroProducesMeasurementRowWithZero();
    void hasPerRowRegime_derivationPreserved();
    void doubleValues_roundTripBitExact();
    void migratedRows_carrySourceUpdatedBy();
    void migration_isIdempotent();

    // --- the H22 / D9 lifecycle tripwire --------------------------------------
    // Stated here in 3b, BEFORE 3c's save path can depend on it. See the long
    // banner on the first one if you arrived because it turned red.
    void orphanTripwire_noMeasurementWithoutItsDataRow();
    void orphanTripwire_dataRowDeleteDoesNotCascadeToMeasurements();

    // --- v3 Phase 3c Task 1: MetricDefCache -----------------------------------
    // The (kind, key) -> metric_defs.id resolver every extras write goes
    // through. Registry and seeded keys are already there; a custom column like
    // coil_temp is in NEITHER and has to be registered on first save or its data
    // has nowhere to go. Like the poka-yokes below, the two slots that write run
    // inside an always-rolled-back transaction, so no auto-registered key leaks
    // into the shared container (metric_defs is deliberately never wiped).
    void metricDefCache_resolvesSeededKeyWithoutInserting();
    void metricDefCache_registersUnknownKeyExactlyOnce();
    void metricDefCache_neverRewritesAnExistingRow();
    void metricDefCache_infersValueTypeInEnvelopeOrder();
    void metricDefCache_encodeDecodeRoundTripsEveryType();

    // --- the poka-yokes -------------------------------------------------------
    // Each of these mutates the database inside a transaction that is always
    // rolled back, so they leave the migrated dataset above untouched and can
    // run in any order.
    void preflight_abortsOnNullSortOrder();
    void preflight_abortsOnDuplicateSortOrder();
    void preflight_abortsOnAllNullDataRow();
    void migration_abortsOnMissingMetricDefSeed();

    // --- v3 Phase 3c Task 5: THE END-TO-END GATE ------------------------------
    // DECLARED LAST ON PURPOSE. Qt Test runs slots in declaration order, and
    // these three are the only ones in this file that COMMIT rows outside the
    // migrated fixture - a file whose data_rows have no measurements for the
    // standard columns, which is exactly what the parity slots above would
    // (correctly) call a broken migration. They also clean up after themselves,
    // before and after, so a single-slot run is still self-contained.
    void e2eGate_manifestDemoCustomColumnSurvivesPostgres();
    void e2eGate_manifestDemoSurvivesASecondRoundTrip();
    void e2eGate_manifestDemoCustomColumnSurvivesOfflineSnapshot();

private:
    QSqlDatabase db() const { return QSqlDatabase::database(QString::fromLatin1(kConn)); }

    void    wipe();
    void    seed();
    qint64  scalar(const QString& sql);
    QString textScalar(const QString& sql);

    void checkWideColumnArity(const QString& table, const QStringList& structuralCols,
                              const QStringList& valueCols, const QString& harnessList,
                              const QString& migrationArray);
    void checkParity(const QString& srcTable, const QString& viewName,
                     const QString& joinPred, const QStringList& cols);
    void checkChecksum(const QString& kind, const QString& key, bool numeric,
                       const QString& srcTable, const QString& srcOrder,
                       const QString& longTable, const QString& defCol,
                       const QString& longOrder);
    void expectAbort(const QStringList& setup, const QStringList& mustContain);

    // ---- v3 Phase 3c Task 5 -------------------------------------------------
    enum class DemoState { NotTried, Unavailable, Ready };

    void        ensureDemoWorkbook();
    void        verifyDemoCoilTemp(const DVE::FileResult& fr, const char* phase);
    void        clearE2eResidue();
    QStringList metricDefKeys();
    qint64      storedCoilTempCount(qint64 fileId);
    qint64      strandedMeasurements(qint64 fileId);
    qint64      sqliteScalar(const QString& file, const QString& sql);

    // The demo workbook, parsed ONCE. ExcelReader spawns a python subprocess
    // per processFile() and those can be slow under load, so all three gate
    // slots share one parse; each takes its own copy before mutating anything.
    DemoState       m_demoState = DemoState::NotTried;
    DVE::FileResult m_demo;
    QString         m_demoError;

    // Shape of the seeded dataset, filled by seed().
    int    m_files      = 0;
    int    m_sheets     = 0;
    int    m_samples    = 0;
    int    m_dataRows   = 0;
    qint64 m_regimeSampleId = -1;   // a sample on a sheet WITH per-row regime
    qint64 m_plainSampleId  = -1;   // a sample on a sheet with NO per-row regime
    qint64 m_trickySampleId = -1;   // holds kTrickyDouble in tpm at sort_order 0

    // First-run migration return values, captured in initTestCase().
    qint64 m_measInserted = -1;
    qint64 m_hdrInserted  = -1;
    qint64 m_dataRowsSeen = -1;
    qint64 m_samplesSeen  = -1;
};

// ===========================================================================
// Fixture
// ===========================================================================

void TstV3LongFormat::wipe()
{
    // measurements / sample_headers go FIRST. ON DELETE CASCADE from samples
    // would take them anyway, but deleting them explicitly keeps the intent
    // visible and stays correct if that cascade is ever tightened.
    //
    // metric_defs is deliberately NOT in this list. It is the VOCABULARY table,
    // not data: dve_migrate_to_long_format() RAISEs if any of the 35 keys is
    // missing. Since the 2026-08-03 "cover all the columns" ruling every one of
    // those 35 is also a compiled-registry key, so ensureSchema() would now put
    // them back - but wiping it would still orphan every historical
    // measurement, whose metric_id resolves through these rows, and would break
    // any suite that runs before the next connect.
    static const QStringList tables = {
        "measurements", "sample_headers",
        "data_rows", "images", "samples", "tests", "files",
        "sensory_images", "sensory_sessions",
        "detailed_sensory_images", "detailed_sensory_sessions",
        "settings"
    };
    for (const QString& t : tables) {
        QSqlQuery q(db());
        q.exec(QStringLiteral("DELETE FROM ") + t);
    }

    // ... but the KEYS the 3c gate invents are not vocabulary anybody else owns,
    // and they outlive this process. Hazard H24 in full is on e2eInventedKeys().
    // After the table loop, so the measurements FK is already clear.
    QSqlQuery q(db());
    for (const QString& k : e2eInventedKeys()) {
        q.prepare(QStringLiteral("DELETE FROM metric_defs WHERE kind = ? AND key = ?"));
        q.addBindValue(k.section(QLatin1Char('|'), 0, 0));
        q.addBindValue(k.section(QLatin1Char('|'), 1, 1));
        q.exec();
    }
}

// Seeds a prod-shaped dataset straight into the wide tables. Shapes come from
// the 2026-07-30 corpus audit: 1-6 sheets per file, ~4-9 samples per sheet,
// 12-33 data rows per sample. Everything is a pure function of the indices, so
// the fixture is deterministic and any failure is reproducible.
//
// The mix is deliberately awkward:
//   - smell / clog / notes are mostly NULL (the sparse rule's main exercise),
//     and one note carries an apostrophe;
//   - resistance and draw_pressure carry scattered numeric NULLs;
//   - draw_pressure / tpm_power_density / variation_tpm / oil_consumed carry
//     GENUINE 0.0 values, which must migrate as rows holding 0;
//   - data_rows.puffing_regime is entirely NULL on every even-indexed sheet
//     (old-template shape) and populated on every odd one, so hasPerRowRegime
//     has both cases to preserve;
//   - updated_by varies per sample so the carry-through check is not vacuous;
//   - one tpm value is 0.1 + 0.2.
void TstV3LongFormat::seed()
{
    static const int kSheetsPerFile[] = { 2, 4, 3 };
    const int fileCount = static_cast<int>(sizeof(kSheetsPerFile) / sizeof(kSheetsPerFile[0]));

    QSqlDatabase d = db();
    QVERIFY2(d.transaction(), qPrintable(d.lastError().text()));
    QSqlQuery q(d);

    int gSheet  = 0;   // global sheet index
    int gSample = 0;   // global sample index
    int gRow    = 0;   // global data-row index

    for (int f = 0; f < fileCount; ++f) {
        DB_EXEC(q, QStringLiteral(
            "INSERT INTO files (file_path, file_name, loaded_at, template_version, "
            "sheet_count, sample_count, updated_by) "
            "VALUES ('/tmp/tst_v3longformat_f%1.xlsx', 'f%1.xlsx', '2026-07-31', "
            "'new', %2, 0, 'seed-file%1') RETURNING id")
            .arg(f).arg(kSheetsPerFile[f]));
        QVERIFY(q.next());
        const qint64 fileId = q.value(0).toLongLong();

        for (int s = 0; s < kSheetsPerFile[f]; ++s, ++gSheet) {
            // Even sheets are the old-template shape: no per-row puffing regime.
            const bool sheetHasRegime = (gSheet % 2 == 1);

            DB_EXEC(q, QStringLiteral(
                "INSERT INTO tests (file_id, sheet_name, template_version, sort_order, "
                "updated_by) VALUES (%1, 'Sheet %2', 'new', %3, 'seed-file%4') RETURNING id")
                .arg(fileId).arg(gSheet).arg(s).arg(f));
            QVERIFY(q.next());
            const qint64 testId = q.value(0).toLongLong();

            const int samplesInSheet = 4 + (gSheet % 6);
            for (int i = 0; i < samplesInSheet; ++i, ++gSample) {
                const int si = gSample;

                const QString sampleName = QStringLiteral("S%1-%2").arg(gSheet).arg(i);
                const QString sampleIdT  = (si % 7 == 3)
                                             ? QString()
                                             : QStringLiteral("SID-%1").arg(1000 + si);
                const QString dateT      = QStringLiteral("2026-0%1-1%2")
                                             .arg(1 + si % 9).arg(si % 10);
                const QString testerT    = (si % 5 == 2)
                                             ? QString()
                                             : QStringLiteral("Tester %1")
                                                   .arg(QChar(QLatin1Char('A' + si % 6)));
                const QString mediaT     = (si % 3 == 0)
                                             ? QString()
                                             : QStringLiteral("Media-%1").arg(si % 4);
                const QString heatT      = (si % 4 == 1)
                                             ? QString()
                                             : QStringLiteral("Ceramic-%1").arg(si % 3);
                const QString regimeT    = (si % 6 == 5)
                                             ? QString()
                                             : QStringLiteral("MRTP %1s").arg(3 + si % 2);
                const QString burnT      = (si % 9 == 1) ? QStringLiteral("none")  : QString();
                const QString clogT      = (si % 9 == 2) ? QStringLiteral("light") : QString();
                const QString leakT      = (si % 9 == 3) ? QStringLiteral("none")  : QString();

                const QStringList sampleVals = {
                    QString::number(testId),
                    QString::number(i),
                    sqlTextOrNull(sampleName),
                    sqlTextOrNull(sampleIdT),
                    sqlTextOrNull(dateT),
                    sqlTextOrNull(testerT),
                    sqlTextOrNull(mediaT),
                    sqlNumOrNull(si % 8 != 5, (si % 6 == 0) ? 0.0 : 1200.0 + si * 3.5),
                    sqlNum(1.15 + si * 0.007),
                    sqlNum((si % 9 == 4) ? 0.0 : 3.2 + si * 0.011),
                    sqlNum(8.0 + si * 0.13),
                    sqlTextOrNull(heatT),
                    sqlTextOrNull(regimeT),
                    sqlNum(1.0 + si * 0.021),
                    sqlNum(3.1 + si * 0.0179),
                    sqlNum((si % 7 == 0) ? 0.0 : 0.11 + si * 0.0009),
                    sqlNumOrNull(si % 11 != 6, 0.0043 + si * 1e-5),
                    sqlNum(55.0 + si * 0.37),
                    sqlNum(0.8 + si * 0.019),
                    QString::number(200 + si * 7),
                    sqlNum(1.05 + si * 0.003),
                    sqlTextOrNull(burnT),
                    sqlTextOrNull(clogT),
                    sqlTextOrNull(leakT),
                    sqlTextOrNull(QStringLiteral("seed-hdr-s%1").arg(si))
                };

                DB_EXEC(q, QStringLiteral(
                    "INSERT INTO samples (test_id, sort_order, sample_name, sample_id, "
                    "date, tester, media, viscosity, resistance, voltage, power, "
                    "heating_technology, puffing_regime, initial_oil_mass, average_tpm, "
                    "stddev_tpm, avg_power_density, efficiency_percent, "
                    "total_oil_consumed, total_puffs, normalized_tpm, burn_status, "
                    "clog_status, leak_status, updated_by) VALUES (%1) RETURNING id")
                    .arg(sampleVals.join(QLatin1String(", "))));
                QVERIFY(q.next());
                const qint64 sampleId = q.value(0).toLongLong();

                if (sheetHasRegime && m_regimeSampleId < 0) m_regimeSampleId = sampleId;
                if (!sheetHasRegime && m_plainSampleId  < 0) m_plainSampleId  = sampleId;

                const int rowsInSample = 12 + (si * 7) % 22;
                const QString rowUpdatedBy = QStringLiteral("seed-f%1-s%2").arg(f).arg(si);

                QStringList tuples;
                tuples.reserve(rowsInSample);
                for (int k = 0; k < rowsInSample; ++k, ++gRow) {
                    const int g = gRow;
                    if (g == 0) m_trickySampleId = sampleId;

                    const double beforeW = 12.0 + g * 0.0031;
                    const QStringList vals = {
                        QString::number(sampleId),
                        QString::number(k),
                        sqlNum(50.0 + (g % 5) * 10.0),                       // puffs
                        sqlNum(beforeW),                                     // before_weight
                        sqlNum(beforeW + 0.0007 * ((g % 17) + 1)),           // after_weight
                        // NULL wins over the genuine zero where the two rules
                        // collide, so both cases stay populated.
                        sqlNumOrNull(g % 13 != 0,
                                     (g % 7 == 0) ? 0.0 : 1.05 + (g % 23) * 0.017),
                        sqlNumOrNull(g % 9 != 4, 1.2 + (g % 11) * 0.013),    // resistance
                        sqlTextOrNull((g % 11 == 3)
                                          ? QStringLiteral("burnt-%1").arg(g % 4)
                                          : QString()),                      // smell
                        sqlTextOrNull((g % 17 == 5)
                                          ? QStringLiteral("partial")
                                          : QString()),                      // clog
                        sqlTextOrNull((g % 23 == 7)
                                          ? QStringLiteral("row %1 o'clock check").arg(g)
                                          : QString()),                      // notes
                        sqlNum((g == 0) ? kTrickyDouble : 2.5 + g * 0.0137), // tpm
                        sqlNum((g % 6 == 0) ? 0.0 : 0.0041 + g * 1e-6),      // tpm_power_density
                        sqlNum((g % 4 == 0) ? 0.0 : 3.7 + (g % 19) * 0.11),  // variation_tpm
                        sqlNum((g % 5 == 0) ? 0.0 : 0.0123 + (g % 29) * 0.0004), // oil_consumed
                        sqlTextOrNull(sheetHasRegime
                                          ? QStringLiteral("%1s on / 30s off").arg(3 + (g % 3))
                                          : QString()),                      // puffing_regime
                        sqlTextOrNull(rowUpdatedBy)
                    };
                    tuples << QLatin1Char('(') + vals.join(QLatin1String(", ")) + QLatin1Char(')');
                }

                DB_EXEC(q, QStringLiteral(
                    "INSERT INTO data_rows (sample_id, sort_order, puffs, before_weight, "
                    "after_weight, draw_pressure, resistance, smell, clog, notes, tpm, "
                    "tpm_power_density, variation_tpm, oil_consumed, puffing_regime, "
                    "updated_by) VALUES %1")
                    .arg(tuples.join(QLatin1String(", "))));
            }
        }
    }

    QVERIFY2(d.commit(), qPrintable(d.lastError().text()));

    m_files    = fileCount;
    m_sheets   = gSheet;
    m_samples  = gSample;
    m_dataRows = gRow;

    QVERIFY(m_regimeSampleId > 0);
    QVERIFY(m_plainSampleId  > 0);
    QVERIFY(m_trickySampleId > 0);
}

// Pins a wide table's VALUE-column set - the catalog minus the structural
// columns - to the list this harness carries, and by extension to the migration's
// pinned CONSTANT array of the same names.
//
// WHY: the migration drives off two hand-written CONSTANT arrays (c_metric_cols /
// c_header_cols in 2026-07-31-v3-long-format.sql), the compat views hand-write the
// same names again as FILTER clauses, and kMetricCols / kHeaderCols above are a
// third hand copy. All three are exact today. The exposure is DRIFT: a column
// added to data_rows in a later phase is skipped by the migration, is not exposed
// by the views, and is absent from this harness's copy too - so it produces ZERO
// failures anywhere and just silently stops migrating. That is not hypothetical:
// ensureSchema's kAdditiveColumns[] (src/database/DatabaseManager.cpp:149-166)
// already carries data_rows.puffing_regime, so ALTER is the ESTABLISHED way a
// column arrives. The catalog is the only source of truth none of the three
// copies can drift away from.
//
// Compared as an unordered SET on purpose, and that is not fussiness. Physical
// column order is database-VINTAGE-dependent: a database built from a current
// init.sql has puffing_regime inline in canonical position, while one that
// predates that column got it by ALTER and carries it physically LAST. A freshly
// provisioned dve-test-pg is the first shape; production is the second. An
// order-sensitive assertion here would therefore pass on this container and
// break on the NAS, for a reason that has nothing to do with column coverage.
void TstV3LongFormat::checkWideColumnArity(const QString& table,
                                           const QStringList& structuralCols,
                                           const QStringList& valueCols,
                                           const QString& harnessList,
                                           const QString& migrationArray)
{
    const QString diff = textScalar(QStringLiteral(
        "WITH catalog AS ("
        "  SELECT column_name FROM information_schema.columns "
        "   WHERE table_schema = 'public' AND table_name = %1 "
        "     AND column_name NOT IN (%2)), "
        "     pinned AS (SELECT unnest(ARRAY[%3]) AS column_name) "
        "SELECT coalesce(string_agg(msg, '; ' ORDER BY msg), '') FROM ("
        "  SELECT 'a real column NOT covered: ' || column_name AS msg "
        "    FROM (SELECT column_name FROM catalog EXCEPT SELECT column_name FROM pinned) a "
        "  UNION ALL "
        "  SELECT 'listed but NOT a real column: ' || column_name "
        "    FROM (SELECT column_name FROM pinned EXCEPT SELECT column_name FROM catalog) b) x")
        .arg(sqlTextOrNull(table), sqlKeyList(structuralCols), sqlKeyList(valueCols)));

    QVERIFY2(diff.isEmpty(),
             qPrintable(QStringLiteral(
                 "%1's value-column set has drifted from what the v3 long-format migration "
                 "covers -> %2.\n"
                 "Update ALL THREE hand-written copies in lockstep, or the column is migrated "
                 "by nobody and missed by every test:\n"
                 "  1. %3 in deploy/postgres/migrations/2026-07-31-v3-long-format.sql\n"
                 "  2. the matching FILTER list in the compat view in that same file\n"
                 "  3. %4 in this harness\n"
                 "and add a metric_defs seed row for the new key, or the migration will RAISE.")
                 .arg(table, diff, migrationArray, harnessList)));
}

void TstV3LongFormat::initTestCase()
{
    if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
        QSKIP("DVE_TEST_PG_CONN not set; skipping");

    QCoreApplication::setOrganizationName("DataViewerTest");
    QCoreApplication::setApplicationName("tst_v3longformat");

    const PgCfg cfg = pgConfig();
    QSqlDatabase pg = QSqlDatabase::addDatabase("QPSQL", QString::fromLatin1(kConn));
    pg.setHostName(cfg.host);
    pg.setPort(cfg.port);
    pg.setDatabaseName(cfg.database);
    pg.setUserName(cfg.user);
    pg.setPassword(cfg.password);
    QVERIFY2(pg.open(), qPrintable(pg.lastError().text()));

    // The long-format objects must be present. A stale container is a silently
    // wrong gate, so this fails loudly with the remediation rather than skipping.
    static const char* const kRemedy =
        "\nRe-provision the container:\n"
        "  docker rm -f dve-test-pg\n"
        "  powershell -ExecutionPolicy Bypass -File tests/start-test-postgres.ps1";
    for (const QString& obj : QStringList{ "measurements", "sample_headers",
                                           "metric_defs", "data_rows_v", "samples_v" }) {
        QVERIFY2(scalar(QStringLiteral("SELECT (to_regclass('%1') IS NOT NULL)::int").arg(obj)) == 1,
                 qPrintable(QStringLiteral("%1 is missing - "
                            "2026-07-31-v3-long-format.sql was not applied.%2")
                            .arg(obj, QString::fromLatin1(kRemedy))));
    }
    QVERIFY2(scalar("SELECT count(*) FROM pg_proc WHERE proname='dve_migrate_to_long_format'") == 1,
             qPrintable(QStringLiteral("dve_migrate_to_long_format() is missing.%1")
                        .arg(QString::fromLatin1(kRemedy))));

    // metric_defs coverage. NOT a bare row count: a migration-only container has
    // 22 kind='header' rows, but after the app's ensureSchema() upsert it has
    // the full compiled registry for both kinds. Only the 35 keys the migration
    // actually maps are asserted, scoped by kind, so this stays correct as the
    // registry grows (it gained 11 header fields and the per-row burn / leak
    // metrics on 2026-08-03).
    QCOMPARE(scalar(QStringLiteral(
                 "SELECT count(*) FROM metric_defs WHERE kind='metric' AND key IN (%1)")
                 .arg(sqlKeyList(kMetricCols))),
             static_cast<qint64>(kMetricCols.size()));
    QCOMPARE(scalar(QStringLiteral(
                 "SELECT count(*) FROM metric_defs WHERE kind='header' AND key IN (%1)")
                 .arg(sqlKeyList(kHeaderCols))),
             static_cast<qint64>(kHeaderCols.size()));

    // Column arity: the migration must still cover every value column that
    // actually exists. Checked before seeding, because a drifted schema makes
    // every downstream parity/checksum result meaningless rather than merely
    // wrong - the comparisons would just shrink in lockstep with the omission.
    checkWideColumnArity("data_rows",
                         QStringList{ "id", "sample_id", "sort_order",
                                      "updated_at", "updated_by", "version" },
                         kMetricCols, "kMetricCols", "c_metric_cols");
    if (QTest::currentTestFailed()) return;
    checkWideColumnArity("samples",
                         QStringList{ "id", "test_id", "sort_order",
                                      "updated_at", "updated_by", "version" },
                         kHeaderCols, "kHeaderCols", "c_header_cols");
    if (QTest::currentTestFailed()) return;

    wipe();
    seed();

    QSqlQuery q(db());
    QVERIFY2(q.exec("SELECT * FROM dve_migrate_to_long_format()"),
             qPrintable(QStringLiteral("migration failed: %1").arg(q.lastError().text())));
    QVERIFY(q.next());
    m_measInserted = q.value(0).toLongLong();
    m_hdrInserted  = q.value(1).toLongLong();
    m_dataRowsSeen = q.value(2).toLongLong();
    m_samplesSeen  = q.value(3).toLongLong();

    qInfo().noquote() << QStringLiteral(
        "seed: %1 files / %2 sheets / %3 samples / %4 data rows -> "
        "%5 measurements, %6 sample_headers")
        .arg(m_files).arg(m_sheets).arg(m_samples).arg(m_dataRows)
        .arg(m_measInserted).arg(m_hdrInserted);
}

void TstV3LongFormat::cleanupTestCase()
{
    if (!QSqlDatabase::contains(QString::fromLatin1(kConn))) return;
    {
        QSqlDatabase d = db();
        if (d.isOpen()) {
            wipe();
            d.close();
        }
    }
    QSqlDatabase::removeDatabase(QString::fromLatin1(kConn));
}

// ===========================================================================
// Small query helpers
// ===========================================================================

qint64 TstV3LongFormat::scalar(const QString& sql)
{
    QSqlQuery q(db());
    if (!q.exec(sql) || !q.next()) {
        qWarning().noquote() << "scalar() failed:" << q.lastError().text() << "\n" << sql;
        return -1;
    }
    return q.value(0).toLongLong();
}

QString TstV3LongFormat::textScalar(const QString& sql)
{
    QSqlQuery q(db());
    if (!q.exec(sql) || !q.next()) {
        qWarning().noquote() << "textScalar() failed:" << q.lastError().text() << "\n" << sql;
        return QStringLiteral("<query failed>");
    }
    return q.value(0).toString();
}

// ===========================================================================
// Migration output
// ===========================================================================

void TstV3LongFormat::migration_returnsExpectedCounts()
{
    QCOMPARE(m_dataRowsSeen, static_cast<qint64>(m_dataRows));
    QCOMPARE(m_samplesSeen,  static_cast<qint64>(m_samples));

    // count(<col>) counts non-NULLs, which under the sparse rule is exactly the
    // number of rows the migration owes for that column.
    QStringList metricTerms;
    for (const QString& c : kMetricCols)
        metricTerms << QStringLiteral("count(%1)").arg(c);
    const qint64 expectedMeas =
        scalar(QStringLiteral("SELECT %1 FROM data_rows").arg(metricTerms.join(QLatin1String(" + "))));

    QStringList headerTerms;
    for (const QString& c : kHeaderCols)
        headerTerms << QStringLiteral("count(%1)").arg(c);
    const qint64 expectedHdr =
        scalar(QStringLiteral("SELECT %1 FROM samples").arg(headerTerms.join(QLatin1String(" + "))));

    QCOMPARE(m_measInserted, expectedMeas);
    QCOMPARE(m_hdrInserted,  expectedHdr);

    QCOMPARE(scalar("SELECT count(*) FROM measurements"),    expectedMeas);
    QCOMPARE(scalar("SELECT count(*) FROM sample_headers"),  expectedHdr);

    // A dense migration would be several thousand rows larger; assert the sparse
    // rule actually bit rather than just that the numbers agree with each other.
    QVERIFY2(expectedMeas < static_cast<qint64>(m_dataRows) * kMetricCols.size(),
             "seed defect: no source NULLs at all, so the sparse rule is untested");
}

// ===========================================================================
// Value parity
// ===========================================================================

void TstV3LongFormat::checkParity(const QString& srcTable, const QString& viewName,
                                  const QString& joinPred, const QStringList& cols)
{
    const qint64 nSrc  = scalar(QStringLiteral("SELECT count(*) FROM %1").arg(srcTable));
    const qint64 nView = scalar(QStringLiteral("SELECT count(*) FROM %1").arg(viewName));
    const qint64 nJoin = scalar(QStringLiteral("SELECT count(*) FROM %1 d JOIN %2 v ON %3")
                                    .arg(srcTable, viewName, joinPred));

    QVERIFY2(nSrc > 0, qPrintable(QStringLiteral("%1 is empty - nothing was compared").arg(srcTable)));
    // H18: a source row that produced no long rows at all vanishes from the
    // view, so row-count parity is a first-class assertion, not a formality.
    QVERIFY2(nView == nSrc,
             qPrintable(QStringLiteral("%1 has %2 row(s) but %3 has %4")
                            .arg(srcTable).arg(nSrc).arg(viewName).arg(nView)));
    QVERIFY2(nJoin == nSrc,
             qPrintable(QStringLiteral("only %1 of %2 %3 row(s) found a %4 partner on %5")
                            .arg(nJoin).arg(nSrc).arg(srcTable, viewName, joinPred)));

    QStringList sel;
    sel.reserve(cols.size());
    for (const QString& c : cols)
        sel << QStringLiteral("count(*) FILTER (WHERE NOT (d.%1 IS NOT DISTINCT FROM v.%1))").arg(c);

    QSqlQuery q(db());
    const QString sql = QStringLiteral("SELECT %1 FROM %2 d JOIN %3 v ON %4")
                            .arg(sel.join(QLatin1String(", ")), srcTable, viewName, joinPred);
    QVERIFY2(q.exec(sql) && q.next(),
             qPrintable(q.lastError().text() + QLatin1Char('\n') + sql));

    for (int i = 0; i < cols.size(); ++i) {
        const qint64 bad = q.value(i).toLongLong();
        QVERIFY2(bad == 0,
                 qPrintable(QStringLiteral("%1.%2 differs from %3.%2 on %4 row(s)")
                                .arg(srcTable, cols.at(i), viewName).arg(bad)));
    }
}

void TstV3LongFormat::dataRowsView_matchesDataRowsOnAllValueColumns()
{
    checkParity("data_rows", "data_rows_v",
                "v.sample_id = d.sample_id AND v.sort_order = d.sort_order",
                kMetricCols);
}

void TstV3LongFormat::samplesView_matchesSamplesOnAllHeaderColumns()
{
    // samples_v's group key is `id` (it IS samples.id); `sample_id` is the TEXT
    // business identifier and one of the 22 header columns.
    checkParity("samples", "samples_v", "v.id = d.id", kHeaderCols);
}

// ===========================================================================
// Per-metric checksums
// ===========================================================================

void TstV3LongFormat::checkChecksum(const QString& kind, const QString& key, bool numeric,
                                    const QString& srcTable, const QString& srcOrder,
                                    const QString& longTable, const QString& defCol,
                                    const QString& longOrder)
{
    const QString longVal = numeric ? QStringLiteral("m.value_num") : QStringLiteral("m.value_text");

    // md5 over the ordered ::text form is a BIT-exact checksum: since PG 12,
    // float8 output is the shortest representation that round-trips, so the text
    // form is injective over doubles. sum(::numeric) is exact decimal addition
    // and therefore order-independent, unlike sum(float8).
    QString srcAgg = QStringLiteral(
        "count(*) AS n, md5(coalesce(string_agg(%1::text, '|' ORDER BY %2), '')) AS h")
        .arg(key, srcOrder);
    QString migAgg = QStringLiteral(
        "count(*) AS n, md5(coalesce(string_agg(%1::text, '|' ORDER BY %2), '')) AS h")
        .arg(longVal, longOrder);
    QString cmp = QStringLiteral("src.n, mig.n, (src.n = mig.n)::int, (src.h = mig.h)::int");

    if (numeric) {
        srcAgg += QStringLiteral(", sum(%1::numeric) AS s, min(%1::float8) AS mn, "
                                 "max(%1::float8) AS mx").arg(key);
        migAgg += QStringLiteral(", sum(%1::numeric) AS s, min(%1) AS mn, max(%1) AS mx").arg(longVal);
        cmp    += QStringLiteral(", (src.s IS NOT DISTINCT FROM mig.s)::int"
                                 ", (src.mn IS NOT DISTINCT FROM mig.mn)::int"
                                 ", (src.mx IS NOT DISTINCT FROM mig.mx)::int");
    }

    const QString sql =
        QStringLiteral("WITH src AS (SELECT ") + srcAgg
        + QStringLiteral(" FROM ") + srcTable
        + QStringLiteral(" WHERE ") + key + QStringLiteral(" IS NOT NULL), mig AS (SELECT ")
        + migAgg
        + QStringLiteral(" FROM ") + longTable
        + QStringLiteral(" m JOIN metric_defs md ON md.id = m.") + defCol
        + QStringLiteral(" WHERE md.kind = ") + sqlTextOrNull(kind)
        + QStringLiteral(" AND md.key = ") + sqlTextOrNull(key)
        + QStringLiteral(") SELECT ") + cmp + QStringLiteral(" FROM src, mig");

    QSqlQuery q(db());
    QVERIFY2(q.exec(sql) && q.next(),
             qPrintable(q.lastError().text() + QLatin1Char('\n') + sql));

    const QString label = kind + QLatin1Char('.') + key;
    const qint64 nSrc = q.value(0).toLongLong();
    const qint64 nMig = q.value(1).toLongLong();

    QVERIFY2(nSrc > 0,
             qPrintable(QStringLiteral("seed defect: %1 has no non-NULL source values, "
                                       "so its checksum comparison is vacuous").arg(label)));
    QVERIFY2(q.value(2).toInt() == 1,
             qPrintable(QStringLiteral("%1: source has %2 non-NULL value(s), migrated has %3")
                            .arg(label).arg(nSrc).arg(nMig)));
    QVERIFY2(q.value(3).toInt() == 1,
             qPrintable(QStringLiteral("%1: ordered md5 of the migrated values differs from "
                                       "the source").arg(label)));
    if (numeric) {
        QVERIFY2(q.value(4).toInt() == 1,
                 qPrintable(QStringLiteral("%1: sum(::numeric) differs").arg(label)));
        QVERIFY2(q.value(5).toInt() == 1, qPrintable(QStringLiteral("%1: min differs").arg(label)));
        QVERIFY2(q.value(6).toInt() == 1, qPrintable(QStringLiteral("%1: max differs").arg(label)));
    }
}

void TstV3LongFormat::perMetricChecksums_match()
{
    for (const QString& key : kMetricCols) {
        checkChecksum("metric", key, !kTextMetricCols.contains(key),
                      "data_rows", "sample_id, sort_order",
                      "measurements", "metric_id", "m.sample_id, m.sort_order");
        if (QTest::currentTestFailed()) return;
    }
}

void TstV3LongFormat::perHeaderFieldChecksums_match()
{
    for (const QString& key : kHeaderCols) {
        checkChecksum("header", key, !kTextHeaderCols.contains(key),
                      "samples", "id",
                      "sample_headers", "field_id", "m.sample_id");
        if (QTest::currentTestFailed()) return;
    }
}

// ===========================================================================
// The sparse rule (index D2)
// ===========================================================================

void TstV3LongFormat::sparseRule_sourceNullProducesNoMeasurementRow()
{
    // The rule stated globally: no long row anywhere holds nothing. A dense
    // migration would fail here immediately.
    QCOMPARE(scalar("SELECT count(*) FROM measurements "
                    "WHERE value_num IS NULL AND value_text IS NULL"), 0LL);
    QCOMPARE(scalar("SELECT count(*) FROM sample_headers "
                    "WHERE value_num IS NULL AND value_text IS NULL"), 0LL);

    // And per column, in the direction that matters: a source NULL must have no
    // measurement row at all, not a row carrying NULL.
    for (const QString& key : QStringList{ "smell", "clog", "notes", "puffing_regime",
                                           "resistance", "draw_pressure" }) {
        const qint64 srcNulls = scalar(
            QStringLiteral("SELECT count(*) FROM data_rows WHERE %1 IS NULL").arg(key));
        QVERIFY2(srcNulls > 0,
                 qPrintable(QStringLiteral("seed defect: data_rows.%1 has no NULLs, so the "
                                           "sparse rule is untested for it").arg(key)));

        const qint64 rowsForNulls = scalar(QStringLiteral(
            "SELECT count(*) FROM measurements m "
            "JOIN metric_defs md ON md.id = m.metric_id "
            "JOIN data_rows d ON d.sample_id = m.sample_id AND d.sort_order = m.sort_order "
            "WHERE md.kind = 'metric' AND md.key = '%1' AND d.%1 IS NULL").arg(key));
        QVERIFY2(rowsForNulls == 0,
                 qPrintable(QStringLiteral("data_rows.%1 is NULL on %2 row(s) but the migration "
                                           "still produced %3 measurement row(s) for them")
                                .arg(key).arg(srcNulls).arg(rowsForNulls)));
    }

    // Same on the header side.
    for (const QString& key : QStringList{ "media", "burn_status", "viscosity",
                                           "avg_power_density" }) {
        const qint64 srcNulls = scalar(
            QStringLiteral("SELECT count(*) FROM samples WHERE %1 IS NULL").arg(key));
        QVERIFY2(srcNulls > 0,
                 qPrintable(QStringLiteral("seed defect: samples.%1 has no NULLs").arg(key)));

        const qint64 rowsForNulls = scalar(QStringLiteral(
            "SELECT count(*) FROM sample_headers sh "
            "JOIN metric_defs md ON md.id = sh.field_id "
            "JOIN samples s ON s.id = sh.sample_id "
            "WHERE md.kind = 'header' AND md.key = '%1' AND s.%1 IS NULL").arg(key));
        QVERIFY2(rowsForNulls == 0,
                 qPrintable(QStringLiteral("samples.%1 is NULL on %2 row(s) but %3 "
                                           "sample_headers row(s) exist for them")
                                .arg(key).arg(srcNulls).arg(rowsForNulls)));
    }
}

void TstV3LongFormat::sparseRule_sourceZeroProducesMeasurementRowWithZero()
{
    // The other half of D2: a numeric 0.0 is a real measurement and must survive
    // as a row holding 0. Getting this wrong (treating 0 as "empty") would
    // silently delete real data.
    for (const QString& key : QStringList{ "draw_pressure", "oil_consumed",
                                           "tpm_power_density", "variation_tpm" }) {
        const qint64 srcZeros = scalar(
            QStringLiteral("SELECT count(*) FROM data_rows WHERE %1 = 0").arg(key));
        QVERIFY2(srcZeros > 0,
                 qPrintable(QStringLiteral("seed defect: data_rows.%1 has no genuine zeros")
                                .arg(key)));

        const qint64 migratedZeros = scalar(QStringLiteral(
            "SELECT count(*) FROM data_rows d "
            "JOIN metric_defs md ON md.kind = 'metric' AND md.key = '%1' "
            "JOIN measurements m ON m.sample_id = d.sample_id "
            "                   AND m.sort_order = d.sort_order "
            "                   AND m.metric_id = md.id "
            "WHERE d.%1 = 0 AND m.value_num = 0").arg(key));
        QVERIFY2(migratedZeros == srcZeros,
                 qPrintable(QStringLiteral("data_rows.%1 holds %2 genuine zero(s) but only %3 "
                                           "migrated to a measurement with value_num = 0")
                                .arg(key).arg(srcZeros).arg(migratedZeros)));
    }

    // And on the header side.
    const qint64 srcZeros = scalar("SELECT count(*) FROM samples WHERE stddev_tpm = 0");
    QVERIFY2(srcZeros > 0, "seed defect: samples.stddev_tpm has no genuine zeros");
    QCOMPARE(scalar(
        "SELECT count(*) FROM samples s "
        "JOIN metric_defs md ON md.kind = 'header' AND md.key = 'stddev_tpm' "
        "JOIN sample_headers sh ON sh.sample_id = s.id AND sh.field_id = md.id "
        "WHERE s.stddev_tpm = 0 AND sh.value_num = 0"), srcZeros);
}

// ===========================================================================
// hasPerRowRegime preservation
// ===========================================================================

void TstV3LongFormat::hasPerRowRegime_derivationPreserved()
{
    // SheetResult::hasPerRowRegime is derived at load time from
    // `puffing_regime IS NOT NULL` (DatabaseManager.cpp:1117-1118,
    // OfflineSnapshot.cpp:1695-1696). A dense migration would materialise an
    // empty regime for every row and flip every old-template sheet into
    // per-row-regime mode. Assert the derivation lands identically on both
    // shapes, per sample, for every sample in the database.
    const qint64 withRegime = scalar(
        "SELECT count(*) FROM samples s WHERE EXISTS "
        "(SELECT 1 FROM data_rows d WHERE d.sample_id = s.id AND d.puffing_regime IS NOT NULL)");
    const qint64 withoutRegime = scalar(
        "SELECT count(*) FROM samples s WHERE NOT EXISTS "
        "(SELECT 1 FROM data_rows d WHERE d.sample_id = s.id AND d.puffing_regime IS NOT NULL)");
    QVERIFY2(withRegime > 0 && withoutRegime > 0,
             qPrintable(QStringLiteral("seed defect: need both shapes, got %1 with regime and "
                                       "%2 without").arg(withRegime).arg(withoutRegime)));

    const qint64 disagreements = scalar(
        "SELECT count(*) FROM ("
        "  SELECT EXISTS (SELECT 1 FROM data_rows d "
        "                  WHERE d.sample_id = s.id AND d.puffing_regime IS NOT NULL) AS wide,"
        "         EXISTS (SELECT 1 FROM data_rows_v v "
        "                  WHERE v.sample_id = s.id AND v.puffing_regime IS NOT NULL) AS lng"
        "    FROM samples s) x "
        "WHERE x.wide <> x.lng");
    QVERIFY2(disagreements == 0,
             qPrintable(QStringLiteral("hasPerRowRegime would change for %1 sample(s) after "
                                       "migration").arg(disagreements)));

    // The concrete statement: samples on a no-regime sheet produce ZERO regime
    // measurement rows - not rows holding NULL.
    QCOMPARE(scalar(
        "SELECT count(*) FROM measurements m "
        "JOIN metric_defs md ON md.id = m.metric_id "
        "WHERE md.kind = 'metric' AND md.key = 'puffing_regime' AND m.sample_id IN ("
        "  SELECT s.id FROM samples s WHERE NOT EXISTS "
        "  (SELECT 1 FROM data_rows d WHERE d.sample_id = s.id "
        "                               AND d.puffing_regime IS NOT NULL))"), 0LL);

    QCOMPARE(scalar(QStringLiteral(
        "SELECT count(*) FROM measurements m "
        "JOIN metric_defs md ON md.id = m.metric_id "
        "WHERE md.kind = 'metric' AND md.key = 'puffing_regime' AND m.sample_id = %1")
        .arg(m_plainSampleId)), 0LL);
    QVERIFY(scalar(QStringLiteral(
        "SELECT count(*) FROM measurements m "
        "JOIN metric_defs md ON md.id = m.metric_id "
        "WHERE md.kind = 'metric' AND md.key = 'puffing_regime' AND m.sample_id = %1")
        .arg(m_regimeSampleId)) > 0);
}

// ===========================================================================
// Bit-exact doubles
// ===========================================================================

void TstV3LongFormat::doubleValues_roundTripBitExact()
{
    // float8send is the raw 8 bytes, so this is bit equality and not merely
    // numeric equality - it would catch a value that took a detour through a
    // lower-precision text form on the way into measurements.
    const qint64 mismatched = scalar(
        "SELECT count(*) FROM data_rows d "
        "JOIN metric_defs md ON md.kind = 'metric' AND md.key = 'tpm' "
        "JOIN measurements m ON m.sample_id = d.sample_id "
        "                   AND m.sort_order = d.sort_order "
        "                   AND m.metric_id = md.id "
        "WHERE float8send(d.tpm) <> float8send(m.value_num)");
    QVERIFY2(mismatched == 0,
             qPrintable(QStringLiteral("%1 tpm value(s) are not bit-identical after migration")
                            .arg(mismatched)));

    // And the specific pathological one: 0.1 + 0.2, whose shortest round-trip
    // form needs all 17 significant digits.
    const QString migrated = textScalar(QStringLiteral(
        "SELECT m.value_num::text FROM measurements m "
        "JOIN metric_defs md ON md.id = m.metric_id "
        "WHERE md.kind = 'metric' AND md.key = 'tpm' "
        "  AND m.sample_id = %1 AND m.sort_order = 0").arg(m_trickySampleId));
    QCOMPARE(migrated, QString::fromLatin1(kTrickyDoubleText));

    QCOMPARE(scalar(QStringLiteral(
        "SELECT (float8send(value_num) = float8send(%1::float8))::int "
        "FROM measurements m JOIN metric_defs md ON md.id = m.metric_id "
        "WHERE md.kind = 'metric' AND md.key = 'tpm' "
        "  AND m.sample_id = %2 AND m.sort_order = 0")
        .arg(QString::fromLatin1(kTrickyDoubleText)).arg(m_trickySampleId)), 1LL);

    // The seed itself must have reached Postgres intact, or the check above
    // would be comparing two copies of the same corruption.
    QCOMPARE(textScalar(QStringLiteral("SELECT tpm::text FROM data_rows "
                                       "WHERE sample_id = %1 AND sort_order = 0")
                            .arg(m_trickySampleId)),
             QString::fromLatin1(kTrickyDoubleText));
    QCOMPARE(sqlNum(kTrickyDouble), QString::fromLatin1(kTrickyDoubleText));
}

// ===========================================================================
// updated_by carry-through
// ===========================================================================

void TstV3LongFormat::migratedRows_carrySourceUpdatedBy()
{
    // The migration carries updated_by from the source row and lets the trigger
    // own updated_at / version. data_rows_v surfaces the most recently updated
    // member's updated_by, which for a freshly migrated row is the source value.
    QVERIFY2(scalar("SELECT count(DISTINCT updated_by) FROM data_rows") > 1,
             "seed defect: one updated_by value everywhere makes this check vacuous");

    QCOMPARE(scalar(
        "SELECT count(*) FROM data_rows d "
        "JOIN data_rows_v v ON v.sample_id = d.sample_id AND v.sort_order = d.sort_order "
        "WHERE v.updated_by IS DISTINCT FROM d.updated_by"), 0LL);

    QCOMPARE(scalar(
        "SELECT count(*) FROM samples s "
        "JOIN sample_headers sh ON sh.sample_id = s.id "
        "WHERE sh.updated_by IS DISTINCT FROM s.updated_by"), 0LL);
}

// ===========================================================================
// Idempotence
// ===========================================================================

void TstV3LongFormat::migration_isIdempotent()
{
    const qint64 measBefore = scalar("SELECT count(*) FROM measurements");
    const qint64 hdrBefore  = scalar("SELECT count(*) FROM sample_headers");
    QVERIFY(measBefore > 0 && hdrBefore > 0);

    QSqlQuery q(db());
    DB_EXEC(q, QStringLiteral("SELECT * FROM dve_migrate_to_long_format()"));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toLongLong(), 0LL);   // measurements_inserted
    QCOMPARE(q.value(1).toLongLong(), 0LL);   // sample_headers_inserted

    QCOMPARE(scalar("SELECT count(*) FROM measurements"),   measBefore);
    QCOMPARE(scalar("SELECT count(*) FROM sample_headers"), hdrBefore);

    // ON CONFLICT DO NOTHING must not have fired an UPDATE, so nothing may have
    // churned a version. min = max = 1 is the whole statement.
    QCOMPARE(scalar("SELECT min(version) FROM measurements"),   1LL);
    QCOMPARE(scalar("SELECT max(version) FROM measurements"),   1LL);
    QCOMPARE(scalar("SELECT min(version) FROM sample_headers"), 1LL);
    QCOMPARE(scalar("SELECT max(version) FROM sample_headers"), 1LL);

    // Parity must still hold after the second run.
    checkParity("data_rows", "data_rows_v",
                "v.sample_id = d.sample_id AND v.sort_order = d.sort_order", kMetricCols);
    if (QTest::currentTestFailed()) return;
    checkParity("samples", "samples_v", "v.id = d.id", kHeaderCols);
}

// ===========================================================================
// H22 / D9 - the measurements lifecycle tripwire
// ===========================================================================

void TstV3LongFormat::orphanTripwire_noMeasurementWithoutItsDataRow()
{
    // ┌── READ THIS IF YOU JUST TURNED THIS TEST RED ─────────────────────────┐
    // │ You broke the decision-D9 contract. Hazard H22. Details below.        │
    // └───────────────────────────────────────────────────────────────────────┘
    //
    // A measurement's identity is (sample_id, metric_id, sort_order). The
    // (sample_id, sort_order) half of that NAMES A data_rows ROW - but there is
    // no referential link to data_rows at all. measurements' only lifecycle FK
    // is sample_id -> samples(id) ON DELETE CASCADE, which fires when a SAMPLE
    // is deleted and does exactly nothing when a data ROW is.
    //
    // So the orphan prune in persistFileCore phase C
    // (src/database/DatabaseOps.cpp:759-784) - a bare
    // `DELETE FROM data_rows WHERE id IN (...)` - leaves the deleted row's
    // measurements behind. Delete row 5 of 10 and save: the survivors renumber
    // their sort_order, the stranded measurements re-bind to the WRONG rows on
    // the next load, and sort_order 9 becomes a data_rows_v group with no
    // data_rows partner. Silent data corruption, not a crash.
    //
    // D9 decided the PRUNE owns this, NOT a foreign key. The tempting fix -
    // `measurements.data_row_id BIGINT REFERENCES data_rows(id) ON DELETE
    // CASCADE` - cannot survive 3d, where data_rows is renamed aside and a view
    // takes its name: an FK onto a view is impossible, so it would have to be
    // dropped at exactly the moment the data matters most.
    //
    // The contract is therefore a CODE contract: when 3c's save path removes a
    // data_rows row it must delete that row's measurements by
    // (sample_id, sort_order), in the SAME transaction, driven by the same
    // surviving-row set the prune already computes.
    //
    // This assertion IS that contract, written down. It passes trivially in 3b -
    // nothing writes measurements outside the migration and nothing deletes data
    // rows - and it is here before any code depends on it precisely so that 3c
    // cannot break it quietly. The companion slot below proves it is capable of
    // failing.
    const qint64 orphans = scalar(
        "SELECT count(*) FROM measurements m WHERE NOT EXISTS ("
        "  SELECT 1 FROM data_rows d "
        "   WHERE d.sample_id = m.sample_id AND d.sort_order = m.sort_order)");
    QVERIFY2(orphans == 0,
             qPrintable(QStringLiteral(
                 "%1 measurement row(s) sit at a (sample_id, sort_order) that data_rows does "
                 "not have (hazard H22, decision D9). A data_rows delete did not take its "
                 "measurements with it; on the next load they will re-bind to the wrong rows "
                 "once the survivors renumber. The prune owns this deletion - see the banner "
                 "on this test.").arg(orphans)));
}

void TstV3LongFormat::orphanTripwire_dataRowDeleteDoesNotCascadeToMeasurements()
{
    // ── The FK posture that makes H22 real, pinned (decision D9) ─────────────
    // There is deliberately NO foreign key from measurements to data_rows. If
    // you are reading this because you just added one: it cannot survive 3d (see
    // the banner on the slot above), which is why D9 put the fix in the prune
    // instead. Change D9 before changing this.
    QCOMPARE(scalar(
        "SELECT count(*) FROM pg_constraint c "
        "JOIN pg_class t ON t.oid = c.conrelid "
        "JOIN pg_class r ON r.oid = c.confrelid "
        "WHERE c.contype = 'f' AND t.relname = 'measurements' AND r.relname = 'data_rows'"), 0LL);

    // WHY THERE IS NO sample_headers ANALOGUE OF THE TRIPWIRE ABOVE:
    // a sample_header's identity is (sample_id, field_id) - there is no
    // sort_order dimension, so the only row it can be orphaned from is its
    // `samples` row, and THAT link is a real FK with ON DELETE CASCADE that
    // Postgres enforces. "No sample_headers row at a sample_id absent from
    // samples" would be unfalsifiable by construction, which is the class of
    // assertion this suite is trying not to write. So pin the CASCADE itself -
    // that is the part that can actually drift, and the whole reason the
    // measurements tripwire needs a hand-written partner and this one does not.
    for (const QString& t : QStringList{ "measurements", "sample_headers" }) {
        QVERIFY2(scalar(QStringLiteral(
            "SELECT count(*) FROM pg_constraint c "
            "JOIN pg_class tb ON tb.oid = c.conrelid "
            "JOIN pg_class r ON r.oid = c.confrelid "
            "WHERE c.contype = 'f' AND tb.relname = '%1' AND r.relname = 'samples' "
            "  AND c.confdeltype = 'c'").arg(t)) == 1,
            qPrintable(QStringLiteral(
                "%1 has no ON DELETE CASCADE foreign key to samples. Deleting a sample would "
                "now strand its rows, and the H22 orphan class would widen from data_rows to "
                "samples as well.").arg(t)));
    }

    // ── The demonstration ────────────────────────────────────────────────────
    // Reproduce EXACTLY what the orphan prune does and watch measurements
    // survive it. This proves the tripwire above is not vacuous, and it is the
    // executable statement of H22 itself. Everything runs inside a transaction
    // that is always rolled back, so the migrated dataset is untouched and slot
    // order stays irrelevant.
    QSqlDatabase d = db();
    QVERIFY2(d.transaction(), qPrintable(d.lastError().text()));
    const auto rollback = qScopeGuard([&d] { d.rollback(); });

    const qint64 victimId = scalar(QStringLiteral(
        "SELECT id FROM data_rows WHERE sample_id = %1 ORDER BY sort_order LIMIT 1")
        .arg(m_regimeSampleId));
    QVERIFY(victimId > 0);

    const qint64 measBefore = scalar("SELECT count(*) FROM measurements");
    QVERIFY(measBefore > 0);

    QSqlQuery q(d);
    // Same shape as pruneOrphans() in src/database/DatabaseOps.cpp:759-784.
    DB_EXEC(q, QStringLiteral("DELETE FROM data_rows WHERE id IN (%1)").arg(victimId));

    // Nothing cascaded: the measurements are all still there.
    QCOMPARE(scalar("SELECT count(*) FROM measurements"), measBefore);

    const qint64 stranded = scalar(
        "SELECT count(*) FROM measurements m WHERE NOT EXISTS ("
        "  SELECT 1 FROM data_rows d "
        "   WHERE d.sample_id = m.sample_id AND d.sort_order = m.sort_order)");
    QVERIFY2(stranded > 0,
             "deleting a data_rows row left NO stranded measurements, so the orphan tripwire "
             "above is currently vacuous. Either the seed stopped covering this sample, or "
             "something now cleans measurements up on a data_rows delete. If a schema-level "
             "lifecycle link was added, re-read decision D9 - it cannot survive 3d, and the "
             "tripwire is the contract that replaces it.");
}

// ===========================================================================
// v3 Phase 3c Task 1 - MetricDefCache
// ===========================================================================

void TstV3LongFormat::metricDefCache_resolvesSeededKeyWithoutInserting()
{
    QSqlDatabase d = db();
    const qint64 before = scalar("SELECT count(*) FROM metric_defs");

    DVE::MetricDefCache cache(d, QStringLiteral("tst-writer"));
    QVERIFY2(cache.load(), "MetricDefCache::load() failed against the test container");
    QVERIFY(cache.isAvailable());

    const qint64 metricId = scalar(
        "SELECT id FROM metric_defs WHERE kind = 'metric' AND key = 'tpm'");
    QVERIFY(metricId > 0);
    QCOMPARE(cache.ensureMetric(QStringLiteral("metric"), QStringLiteral("tpm"),
                                QVariant(1.0)), metricId);
    QCOMPARE(cache.valueType(metricId), QStringLiteral("number"));

    // `resistance` exists under BOTH kinds - the registry keeps metrics and
    // header fields in separate namespaces (spec 5 rule 1 / D11), which is why
    // metric_defs is UNIQUE (kind, key) and never UNIQUE (key). A cache keyed on
    // `key` alone would hand the header id to a metric write.
    const qint64 hdrId = scalar(
        "SELECT id FROM metric_defs WHERE kind = 'header' AND key = 'resistance'");
    const qint64 mtrId = scalar(
        "SELECT id FROM metric_defs WHERE kind = 'metric' AND key = 'resistance'");
    QVERIFY(hdrId > 0 && mtrId > 0 && hdrId != mtrId);
    QCOMPARE(cache.ensureMetric(QStringLiteral("header"),
                                QStringLiteral("resistance"), QVariant(1.0)), hdrId);
    QCOMPARE(cache.ensureMetric(QStringLiteral("metric"),
                                QStringLiteral("resistance"), QVariant(1.0)), mtrId);

    // Nothing was written: resolving a known key is a pure map read.
    QCOMPARE(scalar("SELECT count(*) FROM metric_defs"), before);
}

void TstV3LongFormat::metricDefCache_registersUnknownKeyExactlyOnce()
{
    QSqlDatabase d = db();
    QVERIFY2(d.transaction(), qPrintable(d.lastError().text()));
    const auto rollback = qScopeGuard([&d] { d.rollback(); });

    const qint64 before = scalar("SELECT count(*) FROM metric_defs");

    // `stale` loads BEFORE the key exists, so its later ensureMetric takes the
    // COLD path - INSERT ... ON CONFLICT DO NOTHING with nothing returned,
    // followed by a re-select. That is the concurrent-writer convergence case.
    DVE::MetricDefCache stale(d, QStringLiteral("other-writer"));
    QVERIFY(stale.load());

    DVE::MetricDefCache cache(d, QStringLiteral("tst-writer"));
    QVERIFY(cache.load());

    const qint64 id = cache.ensureMetric(QStringLiteral("metric"),
                                         QStringLiteral("coil_temp"), QVariant(210.5));
    QVERIFY2(id > 0, "an unknown key was not registered");
    QCOMPARE(scalar("SELECT count(*) FROM metric_defs"), before + 1);

    const QString row = textScalar(QStringLiteral(
        "SELECT kind || '|' || key || '|' || display_name || '|' || value_type "
        "|| '|' || coalesce(role, '<null>') || '|' || coalesce(unit, '<null>') "
        "|| '|' || coalesce(tags::text, '<null>') || '|' || updated_by "
        "|| '|' || version FROM metric_defs WHERE id = %1").arg(id));
    // display_name IS the key and role is 'measured', deliberately: the
    // human-facing name lives in the workbook manifest, which is not plumbed
    // down to the database layer, and custom columns are not displayed until
    // Phase 4. Enriching it is a Phase 4 task.
    QCOMPARE(row, QStringLiteral(
        "metric|coil_temp|coil_temp|number|measured|<null>|<null>|tst-writer|1"));
    QCOMPARE(cache.valueType(id), QStringLiteral("number"));

    // Same key again, through the same cache (warm) and through the cache that
    // was loaded before the row existed (cold, ON CONFLICT). Neither inserts.
    QCOMPARE(cache.ensureMetric(QStringLiteral("metric"),
                                QStringLiteral("coil_temp"), QVariant(210.5)), id);
    QCOMPARE(stale.ensureMetric(QStringLiteral("metric"),
                                QStringLiteral("coil_temp"), QVariant(210.5)), id);
    QCOMPARE(scalar("SELECT count(*) FROM metric_defs"), before + 1);

    // The same key under the other kind is a DIFFERENT metric and gets its own
    // row - see the namespace note in the slot above.
    const qint64 hdrId = cache.ensureMetric(QStringLiteral("header"),
                                            QStringLiteral("coil_temp"),
                                            QVariant(QStringLiteral("hot")));
    QVERIFY(hdrId > 0 && hdrId != id);
    QCOMPARE(scalar("SELECT count(*) FROM metric_defs"), before + 2);
    QCOMPARE(cache.valueType(hdrId), QStringLiteral("text"));
}

void TstV3LongFormat::metricDefCache_neverRewritesAnExistingRow()
{
    // ensureSchema owns every registry-sourced field of metric_defs. If
    // ensureMetric ever did ON CONFLICT DO UPDATE, the 35 migration-seeded rows
    // would be rewritten on first save and
    // tst_databasemanager::v3LongFormat_metricDefsSeedIsNotRewritten - which
    // demands they still carry version 1 and updated_by='migration' - would turn
    // red. metric_defs is NOT in bump_version's no-op suppression list, so any
    // write at all is visible as version 2.
    QSqlDatabase d = db();
    QVERIFY2(d.transaction(), qPrintable(d.lastError().text()));
    const auto rollback = qScopeGuard([&d] { d.rollback(); });

    // Load BEFORE the row exists, so the resolve below takes the cold path -
    // the only path that issues SQL, and therefore the only one that could
    // possibly rewrite anything.
    DVE::MetricDefCache stale(d, QStringLiteral("tst-writer"));
    QVERIFY(stale.load());

    QSqlQuery q(d);
    DB_EXEC(q, QStringLiteral(
        "INSERT INTO metric_defs (kind, key, display_name, value_type, unit, role, updated_by) "
        "VALUES ('metric', 'coil_temp', 'Coil Temp (C)', 'number', 'C', 'derived', 'someone-else')"));
    const qint64 id = scalar(
        "SELECT id FROM metric_defs WHERE kind = 'metric' AND key = 'coil_temp'");
    QVERIFY(id > 0);

    // Resolve with a hint that DISAGREES with the stored value_type.
    QCOMPARE(stale.ensureMetric(QStringLiteral("metric"), QStringLiteral("coil_temp"),
                                QVariant(QStringLiteral("text-ish"))), id);

    QCOMPARE(textScalar(QStringLiteral(
                 "SELECT display_name || '|' || value_type || '|' || coalesce(unit,'') "
                 "|| '|' || coalesce(role,'') || '|' || updated_by || '|' || version "
                 "FROM metric_defs WHERE id = %1").arg(id)),
             QStringLiteral("Coil Temp (C)|number|C|derived|someone-else|1"));

    // ... and the cache adopts the STORED type, not its own guess, so the write
    // path routes value_num/value_text by what the database actually says.
    QCOMPARE(stale.valueType(id), QStringLiteral("number"));
}

void TstV3LongFormat::metricDefCache_infersValueTypeInEnvelopeOrder()
{
    using DVE::MetricDefCache;
    // The discrimination ORDER mirrors the recovery JSON envelope's
    // extraValueFromJson (src/pipeline/ReportDataJson.cpp:124-140) so the two
    // serializations of DataRow::extra agree about what a value IS.
    QCOMPARE(MetricDefCache::inferValueType(QVariant(QByteArray("\x89PNG"))),
             QStringLiteral("image"));
    QCOMPARE(MetricDefCache::inferValueType(QVariant(QVariantList{ 1.0, 2.5 })),
             QStringLiteral("numberlist"));
    QCOMPARE(MetricDefCache::inferValueType(QVariant(true)),  QStringLiteral("bool"));
    QCOMPARE(MetricDefCache::inferValueType(QVariant(3)),     QStringLiteral("number"));
    QCOMPARE(MetricDefCache::inferValueType(QVariant(qint64(3))), QStringLiteral("number"));
    QCOMPARE(MetricDefCache::inferValueType(QVariant(3.5)),   QStringLiteral("number"));
    QCOMPARE(MetricDefCache::inferValueType(QVariant(QStringLiteral("x"))),
             QStringLiteral("text"));
    // A QStringList is NOT a numberlist - the envelope's {"a": [...]} case is
    // QVariantList only, and a QStringList falls to its `default:` string form.
    QCOMPARE(MetricDefCache::inferValueType(QVariant(QStringList{ "a", "b" })),
             QStringLiteral("text"));
    // Anything else coerces to text, exactly as the envelope's default does.
    QCOMPARE(MetricDefCache::inferValueType(
                 QVariant(QDateTime::fromSecsSinceEpoch(0, QTimeZone::UTC))),
             QStringLiteral("text"));
}

void TstV3LongFormat::metricDefCache_encodeDecodeRoundTripsEveryType()
{
    using DVE::MetricDefCache;
    // encodeValue is the writer's storage contract and decodeValue is the
    // reader's; they live in one file so the two cannot drift. Exactly ONE of
    // value_num / value_text is ever populated - never both, and (sparse rule
    // D2) never neither.
    const auto roundTrip = [](const QString& type, const QVariant& in) {
        QVariant num, text;
        MetricDefCache::encodeValue(type, in, num, text);
        const bool exactlyOne = num.isNull() != text.isNull();
        if (!exactlyOne) return QVariant(QStringLiteral("<both or neither populated>"));
        return MetricDefCache::decodeValue(type, num, text);
    };

    const QVariant d = roundTrip(QStringLiteral("number"), QVariant(210.5));
    QCOMPARE(d.typeId(), int(QMetaType::Double));
    QCOMPARE(d.toDouble(), 210.5);

    // A genuine zero is a real measurement (index D2), not "empty".
    QCOMPARE(roundTrip(QStringLiteral("number"), QVariant(0.0)).toDouble(), 0.0);

    const QVariant b = roundTrip(QStringLiteral("bool"), QVariant(true));
    QCOMPARE(b.typeId(), int(QMetaType::Bool));
    QCOMPARE(b.toBool(), true);
    QCOMPARE(roundTrip(QStringLiteral("bool"), QVariant(false)).toBool(), false);

    const QVariant s = roundTrip(QStringLiteral("text"), QVariant(QStringLiteral("hot")));
    QCOMPARE(s.typeId(), int(QMetaType::QString));
    QCOMPARE(s.toString(), QStringLiteral("hot"));

    const QVariant img = roundTrip(QStringLiteral("image"), QVariant(QByteArray("\x89PNG\r\n")));
    QCOMPARE(img.typeId(), int(QMetaType::QByteArray));
    QCOMPARE(img.toByteArray(), QByteArray("\x89PNG\r\n"));

    const QVariant lst = roundTrip(QStringLiteral("numberlist"),
                                   QVariant(QVariantList{ 1.0, 2.5, 0.0 }));
    QCOMPARE(lst.typeId(), int(QMetaType::QVariantList));
    QCOMPARE(lst.toList().size(), 3);
    QCOMPARE(lst.toList().at(1).toDouble(), 2.5);

    // The RESOLVED type wins over the value's own C++ type, because the
    // database's value_type is what the reader will route on. A numeric string
    // under a 'number' key still lands in value_num...
    QVariant num, text;
    MetricDefCache::encodeValue(QStringLiteral("number"),
                                QVariant(QStringLiteral("12.5")), num, text);
    QVERIFY(!num.isNull() && text.isNull());
    QCOMPARE(num.toDouble(), 12.5);
    // ... and a value that genuinely cannot be a number falls back to
    // value_text rather than being silently coerced to 0.
    MetricDefCache::encodeValue(QStringLiteral("number"),
                                QVariant(QStringLiteral("hot")), num, text);
    QVERIFY2(num.isNull() && !text.isNull(),
             "a non-numeric value under a 'number' key was coerced instead of "
             "falling back to value_text");
    QCOMPARE(text.toString(), QStringLiteral("hot"));

    // ---- a 'bool' key holding a WORD -------------------------------------
    // THE trap. QVariant::toBool() on a string is true for everything except
    // "", "0" and "false", so "no" / "n" / "N" / "No" all encoded as TRUE.
    // `did_burn` / `did_clog` / `did_leak` are ValueType::Bool in the compiled
    // registry, SchemaDrivenReader inserts header cells RAW, and did_burn is
    // not a LegacyAdapter fixedHeaderKey - so a "Did this burn?" cell reading
    // "no" reached this branch as a QString and every negative answer inverted.
    // After one save the original text was gone: stable at the wrong value.
    QCOMPARE(roundTrip(QStringLiteral("bool"), QVariant(QStringLiteral("no"))),
             QVariant(false));
    // The accepted spellings are LegacyAdapter::normaliseBCL's, which is where
    // the project's Y/N vocabulary is already defined.
    for (const QString& yes : { QStringLiteral("Y"), QStringLiteral("y"),
                                QStringLiteral("YES"), QStringLiteral(" Yes ") })
        QCOMPARE(roundTrip(QStringLiteral("bool"), QVariant(yes)), QVariant(true));
    for (const QString& no : { QStringLiteral("N"), QStringLiteral("n"),
                               QStringLiteral("NO"), QStringLiteral(" No ") })
        QCOMPARE(roundTrip(QStringLiteral("bool"), QVariant(no)), QVariant(false));
    // Numerics still coerce - 0 is false, anything else true.
    QCOMPARE(roundTrip(QStringLiteral("bool"), QVariant(0)), QVariant(false));
    QCOMPARE(roundTrip(QStringLiteral("bool"), QVariant(1)), QVariant(true));
    QCOMPARE(roundTrip(QStringLiteral("bool"), QVariant(QStringLiteral("0"))),
             QVariant(false));
    // A word the vocabulary does not recognise is NOT guessed at: it keeps its
    // text form, exactly the fallback 'number' takes for "hot".
    QCOMPARE(roundTrip(QStringLiteral("bool"), QVariant(QStringLiteral("maybe"))),
             QVariant(QStringLiteral("maybe")));

    // ---- a 'numberlist' key holding something that is not a list ---------
    // QVariant::toList() returns an EMPTY list for a non-sequence, so the value
    // was destroyed and stored as the literal "[]".
    QCOMPARE(roundTrip(QStringLiteral("numberlist"), QVariant(QStringLiteral("N/A"))),
             QVariant(QStringLiteral("N/A")));
    // A list whose elements are not all numbers keeps them, rather than
    // flattening each one to 0.0 through QVariant::toDouble().
    const QVariant mixedList = roundTrip(
        QStringLiteral("numberlist"),
        QVariant(QVariantList{ 1.0, QStringLiteral("N/A") }));
    QCOMPARE(mixedList.toList().size(), 2);
    QCOMPARE(mixedList.toList().at(0).toDouble(), 1.0);
    QCOMPARE(mixedList.toList().at(1).toString(), QStringLiteral("N/A"));

    // ---- an 'image' key that could not hold bytes ------------------------
    // Only a QByteArray can honestly occupy the base64 slot. Anything else
    // falls back to text - and must not be base64-DECODED into garbage on the
    // way out.
    QCOMPARE(roundTrip(QStringLiteral("image"), QVariant(QStringLiteral("no photo"))),
             QVariant(QStringLiteral("no photo")));

    // ---- 'mixed' follows 'number' ----------------------------------------
    // The live registry has metric|voltage as `mixed`, so a numeric per-row
    // voltage must land in value_num and come back a double, not a QString.
    const QVariant mx = roundTrip(QStringLiteral("mixed"), QVariant(3.75));
    QCOMPARE(mx.typeId(), int(QMetaType::Double));
    QCOMPARE(mx.toDouble(), 3.75);
    QCOMPARE(roundTrip(QStringLiteral("mixed"), QVariant(QStringLiteral("3.75"))).toDouble(),
             3.75);
    QCOMPARE(roundTrip(QStringLiteral("mixed"), QVariant(QStringLiteral("n/a"))),
             QVariant(QStringLiteral("n/a")));

    // ---- a 'text' key holding a container --------------------------------
    // Whatever encodeValue writes, decodeValue must hand exactly that back:
    // under a `text` key the STORED STRING is the value. A list additionally
    // must not be flattened to nothing - QVariant::toString() is EMPTY for a
    // QVariantList, which would destroy it outright rather than demote it.
    MetricDefCache::encodeValue(QStringLiteral("text"),
                                QVariant(QVariantList{ 1.0, 2.5 }), num, text);
    QVERIFY(num.isNull() && !text.isNull());
    QVERIFY2(!text.toString().isEmpty(),
             "a list under a 'text' key was flattened to an empty string");
    QCOMPARE(MetricDefCache::decodeValue(QStringLiteral("text"), num, text),
             QVariant(text.toString()));

    MetricDefCache::encodeValue(QStringLiteral("text"),
                                QVariant(QByteArray("\x89PNG\r\n")), num, text);
    QVERIFY(num.isNull() && !text.isNull());
    QCOMPARE(MetricDefCache::decodeValue(QStringLiteral("text"), num, text),
             QVariant(text.toString()));
}

// ===========================================================================
// Pre-flight aborts
// ===========================================================================

// Runs `setup`, then dve_migrate_to_long_format(), inside a transaction that is
// ALWAYS rolled back - so the deliberate corruption never escapes the test and
// the order of the suite's slots does not matter. The migration call gets its
// own SAVEPOINT: a RAISE inside it poisons the transaction, and rolling back to
// the savepoint is what makes the connection usable again for the follow-up
// counts.
void TstV3LongFormat::expectAbort(const QStringList& setup, const QStringList& mustContain)
{
    QSqlDatabase d = db();
    QVERIFY2(d.transaction(), qPrintable(d.lastError().text()));
    const auto rollback = qScopeGuard([&d] { d.rollback(); });

    QSqlQuery q(d);
    for (const QString& s : setup)
        DB_EXEC(q, s);

    DB_EXEC(q, QStringLiteral("SAVEPOINT sp_migrate"));

    const bool ok = q.exec(QStringLiteral("SELECT * FROM dve_migrate_to_long_format()"));
    const QString err = q.lastError().text();
    QVERIFY2(!ok, "dve_migrate_to_long_format() was expected to abort but it succeeded");

    for (const QString& needle : mustContain) {
        QVERIFY2(err.contains(needle),
                 qPrintable(QStringLiteral("abort message does not mention \"%1\".\nActual: %2")
                                .arg(needle, err)));
    }

    DB_EXEC(q, QStringLiteral("ROLLBACK TO SAVEPOINT sp_migrate"));

    // These two counts confirm the FIXTURE was clean - that clearForAbort()
    // really did empty both long tables before the run - and nothing more.
    //
    // They deliberately do NOT claim to test all-or-nothing insertion, and they
    // could not: a plpgsql function is a single statement to its caller, so the
    // RAISE discarded everything the function did before the ROLLBACK TO
    // SAVEPOINT above ever ran. Given the QVERIFY2(!ok) already established,
    // `count(*) = 0` here is a tautology on the migration's side. Partial
    // insertion is impossible by plpgsql exception semantics, not by anything
    // this migration does or could get wrong.
    QCOMPARE(scalar("SELECT count(*) FROM measurements"),   0LL);
    QCOMPARE(scalar("SELECT count(*) FROM sample_headers"), 0LL);
}

namespace {

// One statement that builds a minimal file/test/sample and hangs `rowsClause`
// off the new sample. `rowsClause` is an INSERT ... SELECT tail that can read
// the sample id from the `s` CTE.
QString tinyFixtureSql(const QString& rowsClause)
{
    return QStringLiteral(
        "WITH f AS (INSERT INTO files (file_path, file_name, loaded_at) "
        "           VALUES ('/tmp/tst_v3longformat_abort.xlsx', 'abort.xlsx', '2026-07-31') "
        "           RETURNING id), "
        "     t AS (INSERT INTO tests (file_id, sheet_name) SELECT id, 'S' FROM f RETURNING id), "
        "     s AS (INSERT INTO samples (test_id, sample_name) SELECT id, 'A' FROM t RETURNING id) ")
        + rowsClause;
}

// Wipe the long tables and every wide row, so "the abort inserted nothing" is a
// statement about this fixture alone.
QStringList clearForAbort()
{
    return QStringList{
        QStringLiteral("DELETE FROM measurements"),
        QStringLiteral("DELETE FROM sample_headers"),
        QStringLiteral("DELETE FROM files")
    };
}

} // namespace

void TstV3LongFormat::preflight_abortsOnNullSortOrder()
{
    // H17: data_rows.sort_order is `INTEGER DEFAULT 0` with no NOT NULL, and
    // MigrationTool copied legacy values verbatim. measurements.sort_order is
    // NOT NULL and identity-bearing, so COALESCE(...,0) would fabricate an
    // ordering that could collide with a real row 0.
    QStringList setup = clearForAbort();
    setup << tinyFixtureSql(
        QStringLiteral("INSERT INTO data_rows (sample_id, sort_order, tpm) "
                       "SELECT id, NULL, 1.0 FROM s"));
    expectAbort(setup, QStringList{ "NULL sort_order" });
}

void TstV3LongFormat::preflight_abortsOnDuplicateSortOrder()
{
    // data_rows has NO unique constraint on (sample_id, sort_order). Under
    // measurements.UNIQUE(sample_id, metric_id, sort_order) the migration's
    // ON CONFLICT DO NOTHING would silently collapse a duplicate and lose a
    // whole data row - this abort is the poka-yoke that makes running against
    // production safe.
    QStringList setup = clearForAbort();
    setup << tinyFixtureSql(
        QStringLiteral("INSERT INTO data_rows (sample_id, sort_order, tpm) "
                       "SELECT s.id, v.o, v.t FROM s, (VALUES (0, 1.0), (0, 2.0)) AS v(o, t)"));
    expectAbort(setup, QStringList{ "duplicate (sample_id, sort_order)",
                                    "sort_order=0", "appears 2 times" });
}

void TstV3LongFormat::preflight_abortsOnAllNullDataRow()
{
    // H18: under the sparse rule a row whose 13 value columns are all NULL
    // produces no long rows at all, has no group, and DISAPPEARS from
    // data_rows_v. The long format cannot express "this row exists but holds
    // nothing", so the migration must refuse rather than drop it.
    //
    // The columns must be spelled out: every numeric one carries DEFAULT 0.0,
    // so simply omitting them would produce a row full of genuine zeros.
    QStringList nulls;
    nulls.reserve(kMetricCols.size());
    for (int i = 0; i < kMetricCols.size(); ++i) nulls << QStringLiteral("NULL");

    QStringList setup = clearForAbort();
    setup << tinyFixtureSql(
        QStringLiteral("INSERT INTO data_rows (sample_id, sort_order, %1) "
                       "SELECT id, 0, %2 FROM s")
            .arg(kMetricCols.join(QLatin1String(", ")), nulls.join(QLatin1String(", "))));
    expectAbort(setup, QStringList{ "ALL 13 value columns NULL", "sort_order=0" });
}

void TstV3LongFormat::migration_abortsOnMissingMetricDefSeed()
{
    // A metric_defs key the migration expects but cannot find must RAISE by
    // name, never silently skip the column - a skipped column is a whole metric
    // missing from the migrated database with no error anywhere.
    QStringList setup = clearForAbort();
    setup << tinyFixtureSql(
        QStringLiteral("INSERT INTO data_rows (sample_id, sort_order, tpm, oil_consumed) "
                       "SELECT id, 0, 1.0, 2.0 FROM s"));
    setup << QStringLiteral("DELETE FROM metric_defs WHERE kind = 'metric' AND key = 'oil_consumed'");
    expectAbort(setup, QStringList{ "row for key oil_consumed" });
}

// ===========================================================================
// v3 Phase 3c Task 5 - THE END-TO-END GATE
//
// Everything above this line tests one layer. These three slots test the
// SENTENCE the whole phase was written for, through the production chain, with
// nothing mocked:
//
//   tests/data/manifest_demo.xlsx  (a very-hidden _dve_schema manifest declaring
//                                   a custom "Coil Temp (C)" column)
//     -> ExcelReader (openpyxl subprocess)
//     -> DataProcessor -> SchemaResolver -> SchemaDrivenReader -> LegacyAdapter
//     -> DataRow::extra["coil_temp"]                       (Phase 2c, shipped)
//     -> DatabaseOps::persistFileCore -> measurements      (3c Task 2)
//     -> DatabaseManager::loadFile                         (3c Task 3)
//     -> OfflineSnapshot regen + read                      (3c Task 4)
//
// Before 3c the fourth arrow did not exist and the fifth returned nothing: a
// custom column has NEVER survived the database.
// ===========================================================================

// Parses the demo workbook once and caches it. A MISSING FIXTURE IS A HARD
// FAILURE - a gate that quietly skips when its input is absent is not a gate. A
// missing python + openpyxl is a SKIP, matching tst_v3inference /
// tst_v3roundtrip / tst_dataprocessor: those are environment facts about the
// machine, not statements about this code.
void TstV3LongFormat::ensureDemoWorkbook()
{
    const QString path = testDataFile(QStringLiteral("manifest_demo.xlsx"));
    QVERIFY2(QFileInfo::exists(path),
             qPrintable(QStringLiteral(
                 "%1 is MISSING - run tests/generate_fixtures.py. The v3 Phase 3c "
                 "end-to-end gate has no input without it and must not be allowed "
                 "to report success.").arg(path)));

    if (m_demoState != DemoState::NotTried) return;   // already parsed, or already known bad

    DVE::DataProcessor dp;
    DVE::FileResult fr = dp.processFile(path);
    if (fr.filePath.isEmpty()) {
        m_demoState = DemoState::Unavailable;
        m_demoError = QStringLiteral(
            "the Excel pipeline is unavailable (bundled/system python + openpyxl "
            "not found, or the subprocess timed out): %1").arg(dp.lastError());
        return;
    }

    // Not an assertion about 3c, an assertion about the FIXTURE: if the workbook
    // stopped being parsed through its manifest, everything below would still be
    // measuring something - just not the thing the phase exists for.
    QVERIFY2(dp.lastFileHadManifest(),
             "manifest_demo.xlsx was parsed WITHOUT its _dve_schema manifest - the "
             "gate would be testing the standard positional path, not the custom "
             "column that Phase 3c is about");

    m_demo      = fr;
    m_demoState = DemoState::Ready;
}

#define REQUIRE_DEMO_WORKBOOK()                                                   \
    do {                                                                          \
        ensureDemoWorkbook();                                                     \
        if (QTest::currentTestFailed()) return;                                   \
        if (m_demoState != DemoState::Ready) QSKIP(qPrintable(m_demoError));      \
    } while (false)

// The gate's referee, in ONE place so all four phases make exactly the same
// claim and a phase cannot quietly check less than its neighbours.
//
// `phase` names the link under test and is in every message: "coil_temp is
// missing" is useless without knowing whether the parse, the write, the read or
// the snapshot dropped it.
//
// NON-VACUITY IS ENFORCED, NOT HOPED FOR. Sheet, sample and row counts are
// asserted against the fixture's pinned shape first, and the number of values
// actually compared is asserted to be the full kDemoSamples * kDemoRowsPerSample
// at the end. An empty tree, a tree with no rows, or rows carrying empty `extra`
// maps all fail an assertion here rather than sailing through a loop that never
// runs.
void TstV3LongFormat::verifyDemoCoilTemp(const DVE::FileResult& fr, const char* phase)
{
    const QString ph = QString::fromLatin1(phase);

    QVERIFY2(fr.sheets.size() == 1,
             qPrintable(QStringLiteral("PHASE %1: expected exactly 1 sheet, got %2")
                            .arg(ph).arg(fr.sheets.size())));
    const DVE::SheetResult& sh = fr.sheets.at(0);
    QVERIFY2(sh.sheetName == QLatin1String(kDemoSheetName),
             qPrintable(QStringLiteral("PHASE %1: sheet is \"%2\", expected \"%3\"")
                            .arg(ph, sh.sheetName, QString::fromLatin1(kDemoSheetName))));
    QVERIFY2(sh.samples.size() == kDemoSamples,
             qPrintable(QStringLiteral("PHASE %1: sheet \"%2\" has %3 sample(s), expected %4")
                            .arg(ph, sh.sheetName).arg(sh.samples.size()).arg(kDemoSamples)));

    int found = 0;
    for (int s = 0; s < sh.samples.size(); ++s) {
        const DVE::SampleResult& sr = sh.samples.at(s);
        QVERIFY2(sr.rows.size() == kDemoRowsPerSample,
                 qPrintable(QStringLiteral("PHASE %1: sample %2 has %3 data row(s), expected %4")
                                .arg(ph).arg(s).arg(sr.rows.size()).arg(kDemoRowsPerSample)));

        // Every header key the manifest declares maps onto a fixed
        // SampleResult member (LegacyAdapter::fixedHeaderKeys), so this workbook
        // has NO sample-level open metrics and sample_headers stays empty for it.
        // Stated rather than assumed: if the fixture ever grows one, this fires
        // and the gate gets extended - and e2eInventedKeys() with it - instead of
        // silently covering only half the feature. sample_headers itself is
        // covered by tst_saveintegrity_e2e and tst_databasemanager.
        QVERIFY2(sr.extra.isEmpty(),
                 qPrintable(QStringLiteral(
                     "PHASE %1: sample %2 carries %3 sample-level open metric(s) (%4). The "
                     "fixture changed; extend this gate and e2eInventedKeys() together.")
                     .arg(ph).arg(s).arg(sr.extra.size())
                     .arg(QStringList(sr.extra.keys()).join(QStringLiteral(", ")))));

        for (int r = 0; r < sr.rows.size(); ++r) {
            const DVE::DataRow& dr = sr.rows.at(r);
            const QVariant v = dr.extra.value(QLatin1String(kCoilTempKey));

            QVERIFY2(v.isValid(),
                     qPrintable(QStringLiteral(
                         "PHASE %1: sample %2 row %3 has NO `coil_temp` - the custom column "
                         "was lost. The row carries %4 open metric(s): [%5].")
                         .arg(ph).arg(s).arg(r).arg(dr.extra.size())
                         .arg(QStringList(dr.extra.keys()).join(QStringLiteral(", ")))));

            // metric_defs says `coil_temp` is a `number`, so the storage contract
            // (MetricDefCache::decodeValue) owes a double on the way back - not a
            // string that happens to compare equal after toDouble(). The parse
            // side is a double too: ExcelReader lowers every JSON number to
            // QVariant(double), so ONE assertion covers every phase.
            QVERIFY2(v.typeId() == QMetaType::Double,
                     qPrintable(QStringLiteral(
                         "PHASE %1: sample %2 row %3 `coil_temp` came back as %4, expected "
                         "double - the value survived but its TYPE did not.")
                         .arg(ph).arg(s).arg(r)
                         .arg(QString::fromLatin1(v.typeName() ? v.typeName() : "<null>"))));

            const double expected = kDemoCoilBase[s] + r;
            QVERIFY2(fuzzyEqual(v.toDouble(), expected, 1e-9),
                     qPrintable(QStringLiteral(
                         "PHASE %1: sample %2 row %3 `coil_temp` = %4, expected %5")
                         .arg(ph).arg(s).arg(r)
                         .arg(v.toDouble(), 0, 'g', 17).arg(expected)));
            ++found;

            // The 13 standard columns are a stated NON-GOAL of 3c: they keep
            // flowing through the wide path byte-for-byte. Spot-check one numeric
            // (positional), one derived (recomputed) and one text column so a
            // green gate cannot mean "the extras survived and the wide path
            // broke". `puffs` is the sharpest of the three - the manifest puts it
            // in PHYSICAL slot 1, not the standard slot 0.
            QVERIFY2(fuzzyEqual(dr.puffs, (r + 1) * 10.0, 1e-9),
                     qPrintable(QStringLiteral("PHASE %1: sample %2 row %3 puffs = %4, "
                                               "expected %5 (standard wide path)")
                                    .arg(ph).arg(s).arg(r).arg(dr.puffs).arg((r + 1) * 10.0)));
            QVERIFY2(fuzzyEqual(dr.tpm, kDemoTpm, 1e-4),
                     qPrintable(QStringLiteral("PHASE %1: sample %2 row %3 tpm = %4, "
                                               "expected %5 (standard wide path)")
                                    .arg(ph).arg(s).arg(r).arg(dr.tpm).arg(kDemoTpm)));
            QVERIFY2(dr.puffingRegime == QLatin1String(kDemoRegime),
                     qPrintable(QStringLiteral("PHASE %1: sample %2 row %3 puffingRegime = "
                                               "\"%4\", expected \"%5\" (standard wide path)")
                                    .arg(ph).arg(s).arg(r)
                                    .arg(dr.puffingRegime,
                                         QString::fromLatin1(kDemoRegime))));
        }
    }

    QVERIFY2(found == kDemoSamples * kDemoRowsPerSample,
             qPrintable(QStringLiteral(
                 "PHASE %1: only %2 of the expected %3 `coil_temp` value(s) were compared - "
                 "this assertion is what stops the gate passing vacuously")
                 .arg(ph).arg(found).arg(kDemoSamples * kDemoRowsPerSample)));
}

// Removes everything a gate slot commits: the file row (children ride the
// CASCADE out) and the vocabulary key the save invents (hazard H24). Called
// BEFORE and AFTER every gate slot, so a single-slot run and an aborted run both
// leave the shared container exactly as they found it.
void TstV3LongFormat::clearE2eResidue()
{
    QSqlQuery q(db());
    q.prepare(QStringLiteral("DELETE FROM files WHERE file_path = ?"));
    q.addBindValue(QString::fromLatin1(kE2eFilePath));
    q.exec();
    // Files first, always: measurements.metric_id references metric_defs, so the
    // delete below fails on an FK if the rows are still there.
    for (const QString& k : e2eInventedKeys()) {
        q.prepare(QStringLiteral("DELETE FROM metric_defs WHERE kind = ? AND key = ?"));
        q.addBindValue(k.section(QLatin1Char('|'), 0, 0));
        q.addBindValue(k.section(QLatin1Char('|'), 1, 1));
        q.exec();
    }
}

// Every (kind|key) in metric_defs, sorted. Diffed across a save so the gate can
// assert the save invented EXACTLY the keys e2eInventedKeys() promises to clean
// up, rather than trusting that list to stay accurate.
QStringList TstV3LongFormat::metricDefKeys()
{
    QStringList out;
    QSqlQuery q(db());
    if (!q.exec(QStringLiteral("SELECT kind || '|' || key FROM metric_defs ORDER BY 1"))) {
        qWarning().noquote() << "metricDefKeys() failed:" << q.lastError().text();
        return out;
    }
    while (q.next()) out << q.value(0).toString();
    return out;
}

// The `coil_temp` measurements actually sitting in Postgres for one file, read
// out of band through this suite's own connection. Independent of the read path
// under test, so "the save wrote them" and "the load found them" are two
// separate claims and a bug in either is attributable.
qint64 TstV3LongFormat::storedCoilTempCount(qint64 fileId)
{
    return scalar(QStringLiteral(
        "SELECT count(*) FROM measurements m "
        "JOIN metric_defs md ON md.id = m.metric_id "
        "JOIN samples s ON s.id = m.sample_id "
        "JOIN tests   t ON t.id = s.test_id "
        "WHERE t.file_id = %1 AND md.kind = 'metric' AND md.key = %2")
        .arg(fileId).arg(sqlTextOrNull(QLatin1String(kCoilTempKey))));
}

// Hazard H22 / decision D9, scoped to one file: no measurement may sit at a
// (sample_id, sort_order) that data_rows does not have. The whole-table twin of
// this lives on orphanTripwire_noMeasurementWithoutItsDataRow above; this is the
// one the LIVE save path can break.
qint64 TstV3LongFormat::strandedMeasurements(qint64 fileId)
{
    return scalar(QStringLiteral(
        "SELECT count(*) FROM measurements m "
        "JOIN samples s ON s.id = m.sample_id "
        "JOIN tests   t ON t.id = s.test_id "
        "WHERE t.file_id = %1 AND NOT EXISTS ("
        "  SELECT 1 FROM data_rows d "
        "   WHERE d.sample_id = m.sample_id AND d.sort_order = m.sort_order)")
        .arg(fileId));
}

// One scalar out of a snapshot .sqlite file, through a throwaway connection.
// Reads the MIRROR directly rather than through OfflineSnapshot's API, so
// "the regen copied the rows" is provable without trusting the offline reader.
qint64 TstV3LongFormat::sqliteScalar(const QString& file, const QString& sql)
{
    const QString cname = QStringLiteral("tst_v3lf_snap_probe");
    qint64 out = -1;
    {
        QSqlDatabase sq = QSqlDatabase::addDatabase(QStringLiteral("QSQLITE"), cname);
        sq.setDatabaseName(file);
        if (sq.open()) {
            QSqlQuery q(sq);
            if (q.exec(sql) && q.next()) out = q.value(0).toLongLong();
            else qWarning().noquote() << "sqliteScalar() failed:" << q.lastError().text();
            sq.close();
        } else {
            qWarning().noquote() << "sqliteScalar() could not open" << file
                                 << ":" << sq.lastError().text();
        }
    }
    QSqlDatabase::removeDatabase(cname);
    return out;
}

// ---------------------------------------------------------------------------
// GATE 1 - process, save, DISCARD, reload. The headline claim.
// ---------------------------------------------------------------------------
void TstV3LongFormat::e2eGate_manifestDemoCustomColumnSurvivesPostgres()
{
    REQUIRE_DEMO_WORKBOOK();

    clearE2eResidue();
    const auto tidy = qScopeGuard([this] { clearE2eResidue(); });

    // ── PHASE 1: the parse ────────────────────────────────────────────────
    // Checked FIRST and named, so a regression in Phase 2's manifest path fails
    // here saying "parse" instead of looking like a database bug two steps on.
    DVE::FileResult fr = m_demo;
    fr.filePath = QLatin1String(kE2eFilePath);
    verifyDemoCoilTemp(fr, "1/parse (ExcelReader -> DataProcessor -> SchemaResolver "
                           "-> LegacyAdapter)");
    if (QTest::currentTestFailed()) return;

    DVE::IdentityManager identity;
    DVE::DatabaseManager dbm;
    QVERIFY2(dbm.open(pgDbConfig(), &identity), qPrintable(dbm.lastError()));
    const auto closeDbm = qScopeGuard([&dbm] { dbm.close(); });

    // Captured AFTER open(), because open() runs ensureSchema(), which upserts
    // the whole compiled MetricRegistry. On a freshly provisioned container that
    // is dozens of rows and none of them are ours; the diff below has to measure
    // what the SAVE invented, not what the connect did.
    const QStringList vocabBefore = metricDefKeys();
    QVERIFY2(!vocabBefore.isEmpty(), "metric_defs is empty - the container is not provisioned");

    // ── PHASE 2: the save ─────────────────────────────────────────────────
    QCOMPARE(dbm.tryWriteFile(fr), DVE::WriteResult::Success);
    QVERIFY2(fr.id > 0, "tryWriteFile reported Success but wrote back no file id");
    const int fileId = static_cast<int>(fr.id);

    // Out of band, so this is a claim about the DATABASE and not about the
    // reader: the 12 values are really in `measurements`.
    QCOMPARE(storedCoilTempCount(fr.id),
             static_cast<qint64>(kDemoSamples * kDemoRowsPerSample));

    // `coil_temp` is in NEITHER the compiled registry nor the migration seed, so
    // the save had to register it. That is the half of Task 1 a custom column
    // depends on, and this is the only place the real pipeline proves it.
    QStringList invented = metricDefKeys();
    for (const QString& k : vocabBefore) invented.removeOne(k);
    QCOMPARE(invented, e2eInventedKeys());
    QCOMPARE(textScalar(QStringLiteral(
                 "SELECT value_type FROM metric_defs WHERE kind='metric' AND key=%1")
                 .arg(sqlTextOrNull(QLatin1String(kCoilTempKey)))),
             QStringLiteral("number"));

    // H22 / D9: nothing stranded at a (sample_id, sort_order) data_rows lacks.
    QCOMPARE(strandedMeasurements(fr.id), 0LL);

    // ── PHASE 3: DISCARD the in-memory result, then reload ────────────────
    // `fr` is cleared rather than merely ignored: everything asserted from here
    // came out of Postgres, and there is no live struct left for it to have come
    // out of instead.
    fr = DVE::FileResult();
    const DVE::FileResult reloaded = dbm.loadFile(fileId);
    QVERIFY2(!reloaded.filePath.isEmpty(),
             qPrintable(QStringLiteral("loadFile(%1) returned nothing: %2")
                            .arg(fileId).arg(dbm.lastError())));

    verifyDemoCoilTemp(reloaded, "3/reload from Postgres");
}

// ---------------------------------------------------------------------------
// GATE 2 - the H1 prune regression guard: reload, save, reload.
// ---------------------------------------------------------------------------
void TstV3LongFormat::e2eGate_manifestDemoSurvivesASecondRoundTrip()
{
    // ┌── READ THIS IF YOU JUST TURNED THIS TEST RED ─────────────────────────┐
    // │ Hazard H1. The orphan prune is eating open metrics again.             │
    // └───────────────────────────────────────────────────────────────────────┘
    //
    // persistFileCore phase C captures a pre-image of every child the file owns
    // and DELETEs whatever the in-memory model did not reproduce. That is safe
    // only while the model can reproduce everything - and before 3c,
    // DatabaseManager::loadFile could not read extras at all, so a
    // load-from-database followed by a save handed the prune a model with no
    // extras in it and they were deleted. Phase-3 audit finding 2; the gap was
    // documented in code at src/MainWindow.cpp:1181-1186 and named Phase 3 as
    // the fix.
    //
    // ONE round trip cannot see that: the first save writes from the freshly
    // PARSED tree, which still has the extras. It takes a reload followed by a
    // save to arm it, which is exactly the sequence below. If the read half ever
    // regresses, gate 1 goes red on its reload and THIS slot goes red on its
    // count - and that pair of failures is the signature.
    REQUIRE_DEMO_WORKBOOK();

    clearE2eResidue();
    const auto tidy = qScopeGuard([this] { clearE2eResidue(); });

    DVE::FileResult fr = m_demo;
    fr.filePath = QLatin1String(kE2eFilePath);

    DVE::IdentityManager identity;
    DVE::DatabaseManager dbm;
    QVERIFY2(dbm.open(pgDbConfig(), &identity), qPrintable(dbm.lastError()));
    const auto closeDbm = qScopeGuard([&dbm] { dbm.close(); });

    QCOMPARE(dbm.tryWriteFile(fr), DVE::WriteResult::Success);
    const int    fileId   = static_cast<int>(fr.id);
    const qint64 expected = kDemoSamples * kDemoRowsPerSample;
    QCOMPARE(storedCoilTempCount(fr.id), expected);

    // Round trip 1: reload, and save the LOADED tree back. The loaded tree
    // carries server ids and versions, so this takes the UPDATE-plus-prune path
    // - the one that used to delete the extras.
    DVE::FileResult loaded = dbm.loadFile(fileId);
    QVERIFY2(!loaded.filePath.isEmpty(), qPrintable(dbm.lastError()));
    verifyDemoCoilTemp(loaded, "trip 2 / reload #1");
    if (QTest::currentTestFailed()) return;

    QCOMPARE(dbm.tryWriteFile(loaded), DVE::WriteResult::Success);
    const qint64 afterResave = storedCoilTempCount(fr.id);
    QVERIFY2(afterResave == expected,
             qPrintable(QStringLiteral(
                 "HAZARD H1: saving a tree that was LOADED from the database left %1 of %2 "
                 "`coil_temp` measurement(s). The orphan prune deleted what the reloaded "
                 "model failed to reproduce - see the banner on this test.")
                 .arg(afterResave).arg(expected)));
    QCOMPARE(strandedMeasurements(fr.id), 0LL);

    // Round trip 2: the values must still be there, still correct, still typed.
    const DVE::FileResult again = dbm.loadFile(fileId);
    QVERIFY2(!again.filePath.isEmpty(), qPrintable(dbm.lastError()));
    verifyDemoCoilTemp(again, "trip 2 / reload #2");
    if (QTest::currentTestFailed()) return;

    // And the re-save did not duplicate the vocabulary key either: ensureMetric
    // resolves the existing row rather than inserting a second one.
    QCOMPARE(scalar(QStringLiteral(
                 "SELECT count(*) FROM metric_defs WHERE kind='metric' AND key=%1")
                 .arg(sqlTextOrNull(QLatin1String(kCoilTempKey)))), 1LL);
}

// ---------------------------------------------------------------------------
// GATE 3 - the same round trip through the offline snapshot.
// ---------------------------------------------------------------------------
void TstV3LongFormat::e2eGate_manifestDemoCustomColumnSurvivesOfflineSnapshot()
{
    // Going offline must not erase a custom column. Before Task 4 the snapshot
    // mirrored neither `measurements` nor `sample_headers`, so a user who lost
    // the NAS got the file back with every custom value silently gone - the same
    // failure as never persisting it, arriving one layer later.
    REQUIRE_DEMO_WORKBOOK();

    clearE2eResidue();
    const auto tidy = qScopeGuard([this] { clearE2eResidue(); });

    // Populate Postgres exactly as gate 1 does.
    DVE::FileResult fr = m_demo;
    fr.filePath = QLatin1String(kE2eFilePath);
    {
        DVE::IdentityManager identity;
        DVE::DatabaseManager dbm;
        QVERIFY2(dbm.open(pgDbConfig(), &identity), qPrintable(dbm.lastError()));
        const auto closeDbm = qScopeGuard([&dbm] { dbm.close(); });
        QCOMPARE(dbm.tryWriteFile(fr), DVE::WriteResult::Success);
        QCOMPARE(storedCoilTempCount(fr.id),
                 static_cast<qint64>(kDemoSamples * kDemoRowsPerSample));
    }

    // Regenerate a snapshot from the populated database, into a temp directory
    // so the developer's real %LOCALAPPDATA% snapshot is never touched.
    QTemporaryDir snapDir;
    QVERIFY2(snapDir.isValid(), qPrintable(snapDir.errorString()));

    DVE::OfflineSnapshot snap;
    snap.setOverrideDirForTesting(snapDir.path());
    {
        DVE::PostgresConnection pg;
        QVERIFY2(pg.open(pgDbConfig()), qPrintable(pg.lastError()));
        const auto closePg = qScopeGuard([&pg] { pg.close(); });
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
    }

    // Independent of the offline reader: the MIRROR really carries the rows.
    // checkColumnArity (3a, release-active) would have aborted the regen on a
    // copy-block arity drift, so a snapshot that exists at all is already
    // structurally sound - this is the value-level half of that.
    QCOMPARE(sqliteScalar(snap.path(), QStringLiteral(
                 "SELECT count(*) FROM measurements m "
                 "JOIN metric_defs md ON md.id = m.metric_id "
                 "WHERE md.kind = 'metric' AND md.key = 'coil_temp'")),
             static_cast<qint64>(kDemoSamples * kDemoRowsPerSample));

    // Open it READ-ONLY, the way an offline app boot does, and read the file back.
    QVERIFY2(snap.openReadOnly(), qPrintable(snap.lastError()));
    const auto closeSnap = qScopeGuard([&snap] { snap.close(); });

    const DVE::FileResult offline = snap.loadFileByPath(QLatin1String(kE2eFilePath));
    QVERIFY2(!offline.filePath.isEmpty(),
             qPrintable(QStringLiteral("the snapshot returned nothing for %1: %2")
                            .arg(QString::fromLatin1(kE2eFilePath), snap.lastError())));

    verifyDemoCoilTemp(offline, "offline snapshot (regen -> read-only SQLite)");
}

QTEST_MAIN(TstV3LongFormat)
#include "tst_v3longformat.moc"

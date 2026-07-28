#include <QtTest/QtTest>

#include <QJsonArray>
#include <QJsonObject>
#include <QList>
#include <QRectF>
#include <QStringList>
#include <QVariant>

#include "pipeline/ReportData.h"
#include "pipeline/ReportDataJson.h"

using namespace DVE;

// Round-trip test for the FileResult <-> JSON serializer (Plan C recovery blob).
// Builds a fully-populated FileResult, serializes it, parses it back, and
// asserts every PERSISTED field at every level survives unchanged. Fields the
// serializer intentionally omits (tpmTrend / puffCounts / images /
// dbDataIncomplete) are asserted to come back empty/default, which documents
// the contract: those are re-derived on restore, never carried in the blob.
//
// In addition to the round-trip, this file carries a DB-FIELD-PARITY TRIPWIRE
// (fieldParityKeySets) -- the second half of the guard the Plan C design doc
// requires (lines 50 & 83: "a round-trip unit test PLUS a field-parity check").
// It pins the EXACT set of JSON keys the serializer emits at each of the four
// levels (file / sheet / sample / row) against a hard-coded contract list.
class TstReportDataJson : public QObject
{
    Q_OBJECT

private:
    // ---- Builders for a richly-populated tree -----------------------------

    static DataRow makeRow(double base)
    {
        DataRow r;
        r.puffs           = base + 1.0;
        r.beforeWeight    = base + 2.5;
        r.afterWeight     = base + 1.5;
        r.drawPressure    = base + 0.25;
        r.resistance      = base + 0.75;
        r.puffingRegime   = QStringLiteral("CRM %1").arg(base);
        r.smell           = QStringLiteral("smell-%1").arg(base);
        r.clog            = QStringLiteral("N");
        r.notes           = QStringLiteral("note for row %1").arg(base);
        r.tpm             = base + 3.14;
        r.tpmPowerDensity = base + 4.2;
        r.variationTPM    = base + 0.05;
        r.oilConsumed     = base + 6.0;
        r.id              = static_cast<qint64>(900000000000LL + static_cast<qint64>(base));
        r.version         = static_cast<int>(base) + 2;
        return r;
    }

    static SampleResult makeSample()
    {
        SampleResult s;
        s.sampleName          = QStringLiteral("Sample Alpha");
        s.sampleID            = QStringLiteral("SA-001");
        s.date                = QStringLiteral("2026-06-04");
        s.tester              = QStringLiteral("CB");
        s.media               = QStringLiteral("Distillate X");
        s.viscosity           = 1234.5;
        s.resistance          = 1.2;
        s.voltage             = 3.7;
        s.power               = 11.4;
        s.heatingTechnology   = QStringLiteral("Ceramic");
        s.puffingRegime       = QStringLiteral("CRM");
        s.initialOilMass      = 500.5;
        s.averageTPM          = 5.55;
        s.stdDevTPM           = 0.44;
        s.averagePowerDensity = 2.22;
        s.efficiencyPercent   = 88.8;
        s.totalOilConsumed    = 123.4;
        s.totalPuffs          = 250;
        s.normalizedTPM       = 1.11;
        s.burnStatus          = QStringLiteral("N");
        s.clogStatus          = QStringLiteral("Y");
        s.leakStatus          = QStringLiteral("N");

        s.rows.append(makeRow(10.0));
        s.rows.append(makeRow(20.0));

        s.extra.insert(QStringLiteral("customMetric"), QVariant(42.0));
        s.extra.insert(QStringLiteral("label"), QVariant(QStringLiteral("hello")));
        s.extra.insert(QStringLiteral("flag"), QVariant(true));

        s.imagePaths   << QStringLiteral("C:/imgs/a.png") << QStringLiteral("C:/imgs/b.png");
        s.imageLayouts << QRectF(0.5, 1.5, 2.5, 3.5) << QRectF(0.0, 0.0, 0.0, 0.0);
        s.imageCrops   << QRectF(0.1, 0.2, 0.3, 0.4) << QRectF(0.0, 0.0, 1.0, 1.0);
        s.imageIds     << static_cast<qint64>(700000000001LL) << static_cast<qint64>(-1);
        s.imageVersions << 5 << 0;

        s.id      = static_cast<qint64>(800000000099LL);
        s.version = 7;
        return s;
    }

    static SheetResult makeNormalSheet()
    {
        SheetResult sh;
        sh.sheetName        = QStringLiteral("Test 1");
        sh.templateVersion  = QStringLiteral("new");
        sh.hasPerRowRegime  = true;
        sh.fromInferredSchema = true;   // non-default -> round-trip exercises it
        sh.columnHeaders    << QStringLiteral("Puffs") << QStringLiteral("Before")
                            << QStringLiteral("After");
        sh.overallAvgTPM    = 6.66;
        sh.overallStdDevTPM = 0.33;
        sh.isRawTable       = false;
        sh.samples.append(makeSample());

        // OMITTED fields populated on purpose, to prove they DON'T survive.
        sh.tpmTrend   << 1.0 << 2.0 << 3.0;
        sh.puffCounts << 50.0 << 100.0;
        sh.images.append(QByteArray("\x89PNG-fake-bytes", 15));
        sh.dbDataIncomplete = true;

        sh.id      = static_cast<qint64>(600000000050LL);
        sh.version = 4;
        return sh;
    }

    static SheetResult makeRawSheet()
    {
        SheetResult sh;
        sh.sheetName       = QStringLiteral("Test SOP's");
        sh.templateVersion = QStringLiteral("new");
        sh.isRawTable      = true;
        sh.rawHeaders      << QStringLiteral("Step") << QStringLiteral("Instruction");
        QStringList row1;  row1 << QStringLiteral("1") << QStringLiteral("Do the thing");
        QStringList row2;  row2 << QStringLiteral("2") << QStringLiteral("Do the next thing");
        sh.rawRows << row1 << row2;
        sh.id      = static_cast<qint64>(600000000051LL);
        sh.version = 1;
        return sh;
    }

    static FileResult makeFile()
    {
        FileResult f;
        f.filePath        = QStringLiteral("C:/data/test.xlsx");
        f.fileName        = QStringLiteral("test.xlsx");
        f.templateVersion = QStringLiteral("new");
        f.sheetNames      << QStringLiteral("Test 1") << QStringLiteral("Test SOP's");
        f.sheets.append(makeNormalSheet());
        f.sheets.append(makeRawSheet());
        f.id      = static_cast<qint64>(500000000123LL);
        f.version = 9;
        return f;
    }

    // ---- Level-by-level comparison helpers --------------------------------

    static void compareRow(const DataRow& a, const DataRow& b)
    {
        QCOMPARE(a.puffs,           b.puffs);
        QCOMPARE(a.beforeWeight,    b.beforeWeight);
        QCOMPARE(a.afterWeight,     b.afterWeight);
        QCOMPARE(a.drawPressure,    b.drawPressure);
        QCOMPARE(a.resistance,      b.resistance);
        QCOMPARE(a.puffingRegime,   b.puffingRegime);
        QCOMPARE(a.smell,           b.smell);
        QCOMPARE(a.clog,            b.clog);
        QCOMPARE(a.notes,           b.notes);
        QCOMPARE(a.tpm,             b.tpm);
        QCOMPARE(a.tpmPowerDensity, b.tpmPowerDensity);
        QCOMPARE(a.variationTPM,    b.variationTPM);
        QCOMPARE(a.oilConsumed,     b.oilConsumed);
        QCOMPARE(a.extra,           b.extra);
        QCOMPARE(a.id,              b.id);
        QCOMPARE(a.version,         b.version);
    }

    static void compareSample(const SampleResult& a, const SampleResult& b)
    {
        QCOMPARE(a.sampleName,          b.sampleName);
        QCOMPARE(a.sampleID,            b.sampleID);
        QCOMPARE(a.date,                b.date);
        QCOMPARE(a.tester,              b.tester);
        QCOMPARE(a.media,               b.media);
        QCOMPARE(a.viscosity,           b.viscosity);
        QCOMPARE(a.resistance,          b.resistance);
        QCOMPARE(a.voltage,             b.voltage);
        QCOMPARE(a.power,               b.power);
        QCOMPARE(a.heatingTechnology,   b.heatingTechnology);
        QCOMPARE(a.puffingRegime,       b.puffingRegime);
        QCOMPARE(a.initialOilMass,      b.initialOilMass);
        QCOMPARE(a.averageTPM,          b.averageTPM);
        QCOMPARE(a.stdDevTPM,           b.stdDevTPM);
        QCOMPARE(a.averagePowerDensity, b.averagePowerDensity);
        QCOMPARE(a.efficiencyPercent,   b.efficiencyPercent);
        QCOMPARE(a.totalOilConsumed,    b.totalOilConsumed);
        QCOMPARE(a.totalPuffs,          b.totalPuffs);
        QCOMPARE(a.normalizedTPM,       b.normalizedTPM);
        QCOMPARE(a.burnStatus,          b.burnStatus);
        QCOMPARE(a.clogStatus,          b.clogStatus);
        QCOMPARE(a.leakStatus,          b.leakStatus);

        QCOMPARE(a.rows.size(), b.rows.size());
        for (int i = 0; i < a.rows.size(); ++i)
            compareRow(a.rows[i], b.rows[i]);

        QCOMPARE(a.extra, b.extra);

        QCOMPARE(a.imagePaths,    b.imagePaths);
        QCOMPARE(a.imageLayouts,  b.imageLayouts);
        QCOMPARE(a.imageCrops,    b.imageCrops);
        QCOMPARE(a.imageIds,      b.imageIds);
        QCOMPARE(a.imageVersions, b.imageVersions);

        QCOMPARE(a.startColumn, b.startColumn);

        QCOMPARE(a.id,      b.id);
        QCOMPARE(a.version, b.version);
    }

    static void compareSheet(const SheetResult& a, const SheetResult& b)
    {
        QCOMPARE(a.sheetName,          b.sheetName);
        QCOMPARE(a.templateVersion,    b.templateVersion);
        QCOMPARE(a.hasPerRowRegime,    b.hasPerRowRegime);
        QCOMPARE(a.fromInferredSchema, b.fromInferredSchema);
        QCOMPARE(a.columnHeaders,      b.columnHeaders);
        QCOMPARE(a.overallAvgTPM,    b.overallAvgTPM);
        QCOMPARE(a.overallStdDevTPM, b.overallStdDevTPM);
        QCOMPARE(a.isRawTable,       b.isRawTable);
        QCOMPARE(a.rawHeaders,       b.rawHeaders);
        QCOMPARE(a.rawRows,          b.rawRows);
        QCOMPARE(a.blockCols,        b.blockCols);
        QCOMPARE(a.dataStartRow,     b.dataStartRow);
        QCOMPARE(a.columnKeys,       b.columnKeys);
        QCOMPARE(a.headerCells,      b.headerCells);
        QCOMPARE(a.id,               b.id);
        QCOMPARE(a.version,          b.version);

        QCOMPARE(a.samples.size(), b.samples.size());
        for (int i = 0; i < a.samples.size(); ++i)
            compareSample(a.samples[i], b.samples[i]);
    }

    // ---- Field-parity helper ----------------------------------------------

    // Sort a JSON object's keys for a deterministic, order-independent compare.
    static QStringList sortedKeys(const QJsonObject& o)
    {
        QStringList keys = o.keys();   // QJsonObject::keys() is already sorted,
        keys.sort();                   // but sort defensively so the contract
        return keys;                   // list below is the single source of truth.
    }

private slots:
    void roundTripPreservesEveryPersistedField()
    {
        const FileResult original = makeFile();

        const QJsonObject json = fileResultToJson(original);
        const FileResult restored = fileResultFromJson(json);

        // ---- File level ----
        QCOMPARE(restored.filePath,        original.filePath);
        QCOMPARE(restored.fileName,        original.fileName);
        QCOMPARE(restored.templateVersion, original.templateVersion);
        QCOMPARE(restored.sheetNames,      original.sheetNames);
        QCOMPARE(restored.id,              original.id);
        QCOMPARE(restored.version,         original.version);

        // ---- Sheets (recurses into samples + rows) ----
        QCOMPARE(restored.sheets.size(), original.sheets.size());
        for (int i = 0; i < original.sheets.size(); ++i)
            compareSheet(original.sheets[i], restored.sheets[i]);
    }

    // DataRow::extra carries open per-row values from inferred (non-standard)
    // schemas (TPM v3 smoke-fix batch). Unlike SampleResult::extra (plain
    // fromVariantMap/toVariantMap), it uses a typed envelope so numbers,
    // strings, and QByteArray round-trip with their exact QVariant type
    // preserved - this asserts both the values AND the typeId for each of
    // the four envelope cases.
    void rowExtraRoundTripsLosslessTypedEnvelope()
    {
        FileResult original = makeFile();
        QVERIFY(!original.sheets.isEmpty());
        QVERIFY(!original.sheets[0].samples.isEmpty());
        QVERIFY(!original.sheets[0].samples[0].rows.isEmpty());

        QMap<QString, QVariant> extra;
        extra.insert(QStringLiteral("pv1"), QVariant(3.14159));
        extra.insert(QStringLiteral("chronology_index"),
                     QVariant(static_cast<qint64>(4200000000LL)));
        extra.insert(QStringLiteral("chronology"), QVariant(QStringLiteral("Day 3 AM")));
        // Non-UTF8 bytes on purpose: proves the envelope round-trips raw
        // bytes, not just text that happens to fit in a QByteArray.
        extra.insert(QStringLiteral("raw_thumb"), QVariant(QByteArray("\x00\xff\x10", 3)));
        original.sheets[0].samples[0].rows[0].extra = extra;

        const FileResult restored = fileResultFromJson(fileResultToJson(original));
        const QMap<QString, QVariant>& got = restored.sheets[0].samples[0].rows[0].extra;

        QCOMPARE(got.size(), extra.size());

        QVERIFY(got.contains(QStringLiteral("pv1")));
        QCOMPARE(got.value(QStringLiteral("pv1")).typeId(), static_cast<int>(QMetaType::Double));
        QCOMPARE(got.value(QStringLiteral("pv1")).toDouble(), 3.14159);

        QVERIFY(got.contains(QStringLiteral("chronology_index")));
        QCOMPARE(got.value(QStringLiteral("chronology_index")).typeId(),
                 static_cast<int>(QMetaType::LongLong));
        QCOMPARE(got.value(QStringLiteral("chronology_index")).toLongLong(),
                 static_cast<qint64>(4200000000LL));

        QVERIFY(got.contains(QStringLiteral("chronology")));
        QCOMPARE(got.value(QStringLiteral("chronology")).typeId(),
                 static_cast<int>(QMetaType::QString));
        QCOMPARE(got.value(QStringLiteral("chronology")).toString(),
                 QStringLiteral("Day 3 AM"));

        QVERIFY(got.contains(QStringLiteral("raw_thumb")));
        QCOMPARE(got.value(QStringLiteral("raw_thumb")).typeId(),
                 static_cast<int>(QMetaType::QByteArray));
        QCOMPARE(got.value(QStringLiteral("raw_thumb")).toByteArray(),
                 QByteArray("\x00\xff\x10", 3));

        // Whole-map equality too (exercises compareRow's own QCOMPARE(a.extra, b.extra)).
        QCOMPARE(got, extra);
    }

    // {"a"} envelope (TPM v3 Phase 2a): draw_pressure_per_puff carries its
    // assembled PV1..PV5 element list as a QVariantList of doubles - it must
    // round-trip as a list with its QVariant type preserved, not coerce to a
    // string under the "s" fallback.
    void extraEnvelopeRoundTripsNumberList()
    {
        FileResult original = makeFile();
        QVERIFY(!original.sheets.isEmpty());
        QVERIFY(!original.sheets[0].samples.isEmpty());
        QVERIFY(!original.sheets[0].samples[0].rows.isEmpty());

        original.sheets[0].samples[0].rows[0].extra.insert(
            QStringLiteral("draw_pressure_per_puff"),
            QVariantList{1.1, 2.2, 3.3, 4.4, 5.5});

        const FileResult restored = fileResultFromJson(fileResultToJson(original));
        const QVariant v = restored.sheets[0].samples[0].rows[0].extra.value(
            QStringLiteral("draw_pressure_per_puff"));
        QCOMPARE(v.typeId(), static_cast<int>(QMetaType::QVariantList));
        const QVariantList l = v.toList();
        QCOMPARE(l.size(), 5);
        QCOMPARE(l[2].toDouble(), 3.3);
    }

    // Existing recovery snapshots (written before DataRow::extra existed) must
    // stay byte-stable: an extra-less row must serialize with NO "extra" key
    // at all, not an empty object.
    void rowExtraOmittedFromJsonWhenEmpty()
    {
        const FileResult original = makeFile();
        const QJsonObject fileObj = fileResultToJson(original);

        const QJsonArray sheetsArr = fileObj.value("sheets").toArray();
        QVERIFY(!sheetsArr.isEmpty());
        const QJsonArray samplesArr = sheetsArr.at(0).toObject().value("samples").toArray();
        QVERIFY(!samplesArr.isEmpty());
        const QJsonArray rowsArr = samplesArr.at(0).toObject().value("rows").toArray();
        QVERIFY(!rowsArr.isEmpty());

        for (const QJsonValue& rv : rowsArr)
            QVERIFY(!rv.toObject().contains(QStringLiteral("extra")));
    }

    // Write provenance (TPM v3 Phase 2b): blockCols / dataStartRow /
    // columnKeys / headerCells on sheets and startColumn on samples ride the
    // recovery blob so a recovered file keeps derived write-back addresses.
    void writeProvenanceRoundTrips()
    {
        FileResult f = makeFile();
        SheetResult& sh = f.sheets[0];
        sh.blockCols = 13;
        sh.dataStartRow = 5;
        sh.columnKeys = QStringList{QStringLiteral("puffs"), QStringLiteral("before_weight")};
        sh.headerCells.insert(QStringLiteral("media"), QPoint(2, 2));
        sh.samples[0].startColumn = 13;

        const FileResult back = fileResultFromJson(fileResultToJson(f));
        QCOMPARE(back.sheets[0].blockCols, 13);
        QCOMPARE(back.sheets[0].dataStartRow, 5);
        QCOMPARE(back.sheets[0].columnKeys, sh.columnKeys);
        QCOMPARE(back.sheets[0].headerCells.value(QStringLiteral("media")), QPoint(2, 2));
        QCOMPARE(back.sheets[0].samples[0].startColumn, 13);
    }

    void preProvenanceSnapshotsDecodeToNoProvenance()
    {
        // A JSON produced BEFORE 2b (no provenance keys) must decode to the
        // "unknown" defaults so write-back falls back to legacy behavior -
        // and a provenance-less tree must serialize with NO provenance keys
        // at all (not zeros/empties), keeping pre-2b snapshots byte-stable.
        // Same absent-key convention as the row "extra" key.
        const FileResult f = makeFile();                // defaults: no provenance
        const QJsonObject j = fileResultToJson(f);

        const QJsonObject sheetObj = j.value("sheets").toArray().at(0).toObject();
        QVERIFY(!sheetObj.contains(QStringLiteral("block_cols")));
        QVERIFY(!sheetObj.contains(QStringLiteral("data_start_row")));
        QVERIFY(!sheetObj.contains(QStringLiteral("column_keys")));
        QVERIFY(!sheetObj.contains(QStringLiteral("header_cells")));
        const QJsonObject sampleObj = sheetObj.value("samples").toArray().at(0).toObject();
        QVERIFY(!sampleObj.contains(QStringLiteral("start_column")));

        const FileResult back = fileResultFromJson(j);
        QCOMPARE(back.sheets[0].blockCols, 0);
        QCOMPARE(back.sheets[0].samples[0].startColumn, -1);
    }

    void omittedFieldsAreNotCarried()
    {
        // The normal sheet was built with non-empty tpmTrend / puffCounts /
        // images and dbDataIncomplete=true. These are re-derived on restore,
        // so the blob must NOT carry them: after round-trip they are
        // empty/default. This locks in the coverage contract.
        const FileResult original = makeFile();
        const FileResult restored = fileResultFromJson(fileResultToJson(original));

        QVERIFY(!original.sheets[0].tpmTrend.isEmpty());      // sanity: builder populated it
        QVERIFY(!original.sheets[0].images.isEmpty());        // sanity
        QVERIFY(original.sheets[0].dbDataIncomplete);         // sanity

        const SheetResult& sh = restored.sheets[0];
        QVERIFY(sh.tpmTrend.isEmpty());
        QVERIFY(sh.puffCounts.isEmpty());
        QVERIFY(sh.images.isEmpty());
        QCOMPARE(sh.dbDataIncomplete, false);
    }

    // == DB-FIELD-PARITY TRIPWIRE ============================================
    //
    // The Plan C design doc (lines 50 & 83) requires the ReportDataJson field
    // set to mirror what the DB persists, guarded by "a round-trip unit test
    // PLUS a field-parity check." roundTripPreservesEveryPersistedField() is
    // the round-trip half. THIS is the field-parity half: it pins the EXACT set
    // of JSON keys fileResultToJson() emits at each of the four tree levels
    // against a hard-coded contract.
    //
    // ------------------------------------------------------------------------
    //  IF YOU ADD A PERSISTED FIELD to ReportData.h AND its DB mapping, you MUST
    //  ALSO:
    //    1. add it to BOTH fileResultToJson() and fileResultFromJson()
    //       (src/pipeline/ReportDataJson.cpp), and
    //    2. add its JSON key to the matching expected-key list BELOW, and
    //    3. populate it in the makeFile() builder so the round-trip test above
    //       actually exercises it.
    //  This test is the tripwire that FORCES step 2: it fails loudly if a key is
    //  ADDED to the serializer without updating the contract here, OR if a key
    //  is DROPPED from the serializer. Either direction = drift = red.
    //
    //  Keys intentionally NOT serialized (re-derived/transient) -- do NOT add
    //  them here: sheet.tpmTrend, sheet.puffCounts, sheet.images,
    //  sheet.dbDataIncomplete. See omittedFieldsAreNotCarried().
    //
    //  Keys CONDITIONALLY serialized (absent from makeFile()'s default tree,
    //  so absent from the lists below - same convention as row "extra"):
    //  row.extra (only when non-empty), the Phase-2b write-provenance keys
    //  sheet.block_cols / sheet.data_start_row / sheet.column_keys /
    //  sheet.header_cells (only when blockCols > 0) and sample.start_column
    //  (only when >= 0). See writeProvenanceRoundTrips() /
    //  preProvenanceSnapshotsDecodeToNoProvenance().
    // ========================================================================
    void fieldParityKeySets()
    {
        // The contract: the exact JSON keys expected at each tree level. Kept
        // in sorted order to mirror sortedKeys() so the QCOMPARE diff reads
        // cleanly on failure.

        const QStringList kFileKeys = QStringList{
            "file_name",
            "file_path",
            "id",
            "sheet_names",
            "sheets",
            "template_version",
            "version",
        };

        const QStringList kSheetKeys = QStringList{
            "column_headers",
            "from_inferred_schema",
            "has_per_row_regime",
            "id",
            "is_raw_table",
            "overall_avg_tpm",
            "overall_std_dev_tpm",
            "raw_headers",
            "raw_rows",
            "samples",
            "sheet_name",
            "template_version",
            "version",
        };

        const QStringList kSampleKeys = QStringList{
            "average_power_density",
            "average_tpm",
            "burn_status",
            "clog_status",
            "date",
            "efficiency_percent",
            "extra",
            "heating_technology",
            "id",
            "image_crops",
            "image_ids",
            "image_layouts",
            "image_paths",
            "image_versions",
            "initial_oil_mass",
            "leak_status",
            "media",
            "normalized_tpm",
            "power",
            "puffing_regime",
            "resistance",
            "rows",
            "sample_id",
            "sample_name",
            "std_dev_tpm",
            "tester",
            "total_oil_consumed",
            "total_puffs",
            "version",
            "viscosity",
            "voltage",
        };

        const QStringList kRowKeys = QStringList{
            "after_weight",
            "before_weight",
            "clog",
            "draw_pressure",
            "id",
            "notes",
            "oil_consumed",
            "puffing_regime",
            "puffs",
            "resistance",
            "smell",
            "tpm",
            "tpm_power_density",
            "variation_tpm",
            "version",
        };

        // Serialize a fully-populated tree (non-default values throughout) and
        // walk down to one object at each level. makeFile() guarantees >= 1
        // sheet, >= 1 sample, >= 1 row in the first (non-raw) sheet.
        const FileResult original = makeFile();
        const QJsonObject fileObj = fileResultToJson(original);

        // ---- File level ----
        QCOMPARE(sortedKeys(fileObj), kFileKeys);

        // ---- Sheet level (first sheet = the populated, non-raw one) ----
        const QJsonArray sheetsArr = fileObj.value("sheets").toArray();
        QVERIFY(!sheetsArr.isEmpty());
        const QJsonObject sheetObj = sheetsArr.at(0).toObject();
        QCOMPARE(sortedKeys(sheetObj), kSheetKeys);

        // ---- Sample level ----
        const QJsonArray samplesArr = sheetObj.value("samples").toArray();
        QVERIFY(!samplesArr.isEmpty());
        const QJsonObject sampleObj = samplesArr.at(0).toObject();
        QCOMPARE(sortedKeys(sampleObj), kSampleKeys);

        // ---- Row level ----
        const QJsonArray rowsArr = sampleObj.value("rows").toArray();
        QVERIFY(!rowsArr.isEmpty());
        const QJsonObject rowObj = rowsArr.at(0).toObject();
        QCOMPARE(sortedKeys(rowObj), kRowKeys);
    }
};

QTEST_MAIN(TstReportDataJson)
#include "tst_reportdatajson.moc"

#include <QtTest>
#include <QFileInfo>
#include <QJsonArray>
#include "CorpusUtils.h"
#include "JsonDiff.h"
#include "pipeline/DataProcessor.h"
#include "pipeline/ReportDataJson.h"

namespace {

// Phase 2b: the production path records write provenance (block_cols /
// data_start_row / column_keys / header_cells on sheets, start_column on
// samples) and ReportDataJson serializes it - but the legacy positional
// referee (processFileLegacy) deliberately does NOT populate the sheet-level
// provenance, so the two sides would now differ on those keys alone. The
// byte-identity gate's contract is "every LEGACY-ERA byte identical", so the
// diff is taken on the legacy-visible domain: strip the provenance keys from
// BOTH sides before comparing. Provenance correctness is gated elsewhere
// (tst_v3roundtrip + the provenance tests in tst_dataprocessor /
// tst_v3inference / tst_reportdatajson), not by legacy identity.
QJsonObject stripProvenance(QJsonObject obj);

QJsonValue stripProvenanceValue(const QJsonValue& v)
{
    if (v.isObject())
        return stripProvenance(v.toObject());
    if (v.isArray()) {
        QJsonArray out;
        const QJsonArray in = v.toArray();
        for (const QJsonValue& e : in)
            out.append(stripProvenanceValue(e));
        return out;
    }
    return v;
}

QJsonObject stripProvenance(QJsonObject obj)
{
    static const QStringList kProvenanceKeys{
        QStringLiteral("block_cols"), QStringLiteral("data_start_row"),
        QStringLiteral("column_keys"), QStringLiteral("header_cells"),
        QStringLiteral("start_column")};
    for (const QString& k : kProvenanceKeys)
        obj.remove(k);
    for (auto it = obj.begin(); it != obj.end(); ++it)
        it.value() = stripProvenanceValue(it.value());
    return obj;
}

} // namespace

// Phase 0 shadow harness: proves the REAL Excel pipeline (ExcelReader spawns a
// Python/openpyxl subprocess) is deterministic - the same workbook processed
// twice must serialize to byte-identical JSON via the canonical
// DVE::fileResultToJson contract. This is the baseline the Phase 2 round-trip
// harness (schema-driven reader vs legacy) will be compared against.
class TestV3Shadow : public QObject {
    Q_OBJECT
private slots:
    void parseIsDeterministic_data();
    void parseIsDeterministic();
    void productionMatchesLegacyParser_data();
    void productionMatchesLegacyParser();

private:
    // One data row per corpus fixture, tagged "<dir>/<file>".
    static void addCorpusRows();
};

void TestV3Shadow::addCorpusRows()
{
    QTest::addColumn<QString>("path");
    const QStringList files = DVE::testutil::corpusFiles();
    QVERIFY2(!files.isEmpty(), "no corpus files; tests/data fixtures missing");
    for (const QString& f : files) {
        const QFileInfo fi(f);
        const QString tag = QString("%1/%2").arg(fi.dir().dirName(), fi.fileName());
        QTest::newRow(qPrintable(tag)) << f;
    }
}

void TestV3Shadow::parseIsDeterministic_data()
{
    addCorpusRows();
}

void TestV3Shadow::parseIsDeterministic()
{
    QFETCH(QString, path);
    DVE::DataProcessor p1, p2;
    const DVE::FileResult a = p1.processFile(path);
    const DVE::FileResult b = p2.processFile(path);
    // Guard both calls independently: the Python subprocess has its own 60s
    // timeout (ExcelReader::runPythonReader) that a busy machine can blow past
    // on either invocation, not just the first. Treating that as a genuine
    // determinism failure would be a false positive - skip instead, same as
    // the "python not found" case.
    const bool aUnavailable = a.sheets.isEmpty() && p1.lastError().contains("python", Qt::CaseInsensitive);
    const bool bUnavailable = b.sheets.isEmpty() && p2.lastError().contains("python", Qt::CaseInsensitive);
    if (aUnavailable || bUnavailable)
        QSKIP("bundled/system python unavailable or subprocess timed out");
    const QStringList diff = DVE::testutil::diffJson(DVE::fileResultToJson(a),
                                                     DVE::fileResultToJson(b));
    QVERIFY2(diff.isEmpty(), qPrintable(diff.join('\n')));
}

// Phase 1 shadow gate: the PRODUCTION read path (DataProcessor::processFile,
// now schema-driven) must serialize to byte-identical JSON as the legacy
// positional referee (DataProcessor::processFileLegacy) for every corpus
// fixture. This is the correctness gate for the strangler - any divergence is a
// bug in the production path (never the legacy referee or the serializer).
void TestV3Shadow::productionMatchesLegacyParser_data()
{
    addCorpusRows();
}

void TestV3Shadow::productionMatchesLegacyParser()
{
    QFETCH(QString, path);
    DVE::DataProcessor pLegacy, pProd;
    const DVE::FileResult legacyR = pLegacy.processFileLegacy(path);
    const bool legacyUnavailable = legacyR.sheets.isEmpty() && pLegacy.lastError().contains("python", Qt::CaseInsensitive);
    const DVE::FileResult prodR = pProd.processFile(path);
    const bool prodUnavailable = prodR.sheets.isEmpty() && pProd.lastError().contains("python", Qt::CaseInsensitive);
    if (legacyUnavailable || prodUnavailable)
        QSKIP("bundled/system python unavailable or subprocess timed out");
    // Smoke-fix batch: a file whose block layout is NOT the standard 12-wide
    // shape (e.g. the 13-col S26 / 8-col UserSim files, and the in-repo 13-wide
    // format_a analog) now takes the header-driven inference path in production.
    // The legacy positional referee still mangles those the old 12-wide way, so
    // the two DELIBERATELY diverge - correctness there is gated by the inference
    // E2E value tests (tst_v3inference), not by legacy identity. The determinism
    // slot above still covers every file, inference-path ones included.
    if (pProd.lastFileUsedInference())
        QSKIP("inference path - correctness gated by the inference E2E tests, legacy is known-wrong here");
    // Diff on the legacy-visible domain (see stripProvenance above).
    const QStringList diff = DVE::testutil::diffJson(stripProvenance(DVE::fileResultToJson(legacyR)),
                                                     stripProvenance(DVE::fileResultToJson(prodR)));
    QVERIFY2(diff.isEmpty(), qPrintable(diff.join('\n')));
}

QTEST_MAIN(TestV3Shadow)
#include "tst_v3shadow.moc"

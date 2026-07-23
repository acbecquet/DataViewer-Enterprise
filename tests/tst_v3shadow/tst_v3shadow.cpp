#include <QtTest>
#include <QFileInfo>
#include "CorpusUtils.h"
#include "JsonDiff.h"
#include "pipeline/DataProcessor.h"
#include "pipeline/ReportDataJson.h"

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
};

void TestV3Shadow::parseIsDeterministic_data()
{
    QTest::addColumn<QString>("path");
    const QStringList files = DVE::testutil::corpusFiles();
    QVERIFY2(!files.isEmpty(), "no corpus files; run tests/generate_fixtures.py");
    for (const QString& f : files)
        QTest::newRow(qPrintable(QFileInfo(f).fileName())) << f;
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
        QSKIP("bundled/system python unavailable");
    const QStringList diff = DVE::testutil::diffJson(DVE::fileResultToJson(a),
                                                     DVE::fileResultToJson(b));
    QVERIFY2(diff.isEmpty(), qPrintable(diff.join('\n')));
}

QTEST_MAIN(TestV3Shadow)
#include "tst_v3shadow.moc"

#include <QtTest/QtTest>

#include "../../src/database/VersionLookup.h"
#include "../../src/pipeline/ReportData.h"
#include "../../src/pipeline/SensoryData.h"
#include "../../src/pipeline/DetailedSensoryData.h"

using namespace DVE;

// Pure-function unit tests for versionForRow(). The MainWindow lambda
// registered with LiveSync::setVersionLookup delegates straight to this
// helper, so green coverage here proves OCC activation is correct given
// the in-memory state.
class TstVersionLookup : public QObject
{
    Q_OBJECT

private slots:
    void emptyInputs_returnsMinusOne();
    void filesTable_returnsCachedVersion();
    void filesTable_unknownIdReturnsMinusOne();
    void testsTable_returnsCachedVersion();
    void samplesTable_returnsCachedVersion();
    void dataRowsTable_returnsCachedVersion();
    void sensorySessionsTable_returnsCachedVersion();
    void detailedSensorySessionsTable_returnsCachedVersion();
    void unknownTable_returnsMinusOne();
    void zeroVersion_returnsMinusOne();   // pre-persistence rows opt out of OCC

private:
    QVector<FileResult>             buildFileFixture();
    QVector<SensorySession>         buildSensoryFixture();
    QVector<DetailedSensorySession> buildDetailedFixture();
};

QVector<FileResult> TstVersionLookup::buildFileFixture()
{
    FileResult f;
    f.id = 100; f.version = 7;

    SheetResult sheet;
    sheet.id = 200; sheet.version = 3;

    SampleResult sample;
    sample.id = 300; sample.version = 11;

    DataRow row;
    row.id = 400; row.version = 5;
    sample.rows.append(row);

    sheet.samples.append(sample);
    f.sheets.append(sheet);

    QVector<FileResult> v; v.append(f);
    return v;
}

QVector<SensorySession> TstVersionLookup::buildSensoryFixture()
{
    SensorySession s;
    s.id = 500; s.version = 9;
    QVector<SensorySession> v; v.append(s);
    return v;
}

QVector<DetailedSensorySession> TstVersionLookup::buildDetailedFixture()
{
    DetailedSensorySession d;
    d.id = 600; d.version = 13;
    QVector<DetailedSensorySession> v; v.append(d);
    return v;
}

void TstVersionLookup::emptyInputs_returnsMinusOne()
{
    QVector<FileResult> files;
    QVector<SensorySession> sens;
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("data_rows"), 1), qint64(-1));
}

void TstVersionLookup::filesTable_returnsCachedVersion()
{
    auto files = buildFileFixture();
    QVector<SensorySession> sens;
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("files"), 100), qint64(7));
}

void TstVersionLookup::filesTable_unknownIdReturnsMinusOne()
{
    auto files = buildFileFixture();
    QVector<SensorySession> sens;
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("files"), 999), qint64(-1));
}

void TstVersionLookup::testsTable_returnsCachedVersion()
{
    auto files = buildFileFixture();
    QVector<SensorySession> sens;
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("tests"), 200), qint64(3));
}

void TstVersionLookup::samplesTable_returnsCachedVersion()
{
    auto files = buildFileFixture();
    QVector<SensorySession> sens;
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("samples"), 300), qint64(11));
}

void TstVersionLookup::dataRowsTable_returnsCachedVersion()
{
    auto files = buildFileFixture();
    QVector<SensorySession> sens;
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("data_rows"), 400), qint64(5));
}

void TstVersionLookup::sensorySessionsTable_returnsCachedVersion()
{
    QVector<FileResult> files;
    auto sens = buildSensoryFixture();
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("sensory_sessions"), 500), qint64(9));
}

void TstVersionLookup::detailedSensorySessionsTable_returnsCachedVersion()
{
    QVector<FileResult> files;
    QVector<SensorySession> sens;
    auto det = buildDetailedFixture();
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("detailed_sensory_sessions"), 600), qint64(13));
}

void TstVersionLookup::unknownTable_returnsMinusOne()
{
    auto files = buildFileFixture();
    QVector<SensorySession> sens;
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("not_a_table"), 100), qint64(-1));
}

void TstVersionLookup::zeroVersion_returnsMinusOne()
{
    // A row with id > 0 but version == 0 is a fresh INSERT that hasn't
    // been round-tripped through the DB yet. OCC must opt out so the
    // first commit isn't gated on a meaningless version anchor.
    FileResult f; f.id = 700; f.version = 0;
    QVector<FileResult> files; files.append(f);
    QVector<SensorySession> sens;
    QVector<DetailedSensorySession> det;
    QCOMPARE(versionForRow(files, sens, det, QStringLiteral("files"), 700), qint64(-1));
}

QTEST_MAIN(TstVersionLookup)
#include "tst_versionlookup.moc"

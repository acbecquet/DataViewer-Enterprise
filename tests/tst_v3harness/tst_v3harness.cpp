#include <QtTest>
#include <QTemporaryDir>
#include <QJsonObject>
#include <QJsonArray>
#include "CorpusUtils.h"
#include "JsonDiff.h"

class TestV3Harness : public QObject {
    Q_OBJECT
private slots:
    void corpusIncludesFixtures();
    void corpusIncludesEnvDir();
    void corpusSkipsUnsetEnvCleanly();
    void jsonDiffEqual();
    void jsonDiffReportsPath();
    void jsonDiffReportsMissingKey();
};

void TestV3Harness::corpusIncludesFixtures()
{
    qunsetenv("DVE_TEST_CORPUS_DIR");
    const QStringList files = DVE::testutil::corpusFiles();
    QVERIFY(!files.isEmpty());                       // tests/data/*.xlsx always present
    for (const QString& f : files)
        QVERIFY2(f.endsWith(".xlsx"), qPrintable(f));
}

void TestV3Harness::corpusIncludesEnvDir()
{
    QTemporaryDir dir;
    QVERIFY(dir.isValid());
    QFile f(dir.filePath("real_test.xlsx"));
    QVERIFY(f.open(QIODevice::WriteOnly));
    f.write("stub");
    f.close();
    qputenv("DVE_TEST_CORPUS_DIR", dir.path().toUtf8());
    const QStringList files = DVE::testutil::corpusFiles();
    QVERIFY(files.contains(QDir(dir.path()).filePath("real_test.xlsx")));
    qunsetenv("DVE_TEST_CORPUS_DIR");
}

void TestV3Harness::corpusSkipsUnsetEnvCleanly()
{
    qunsetenv("DVE_TEST_CORPUS_DIR");
    QVERIFY(DVE::testutil::corpusDirDescription().contains("fixtures only"));
}

void TestV3Harness::jsonDiffEqual()
{
    QJsonObject a{{"x", 1}, {"nested", QJsonObject{{"y", "z"}}}};
    QCOMPARE(DVE::testutil::diffJson(a, a), QStringList());
}

void TestV3Harness::jsonDiffReportsPath()
{
    QJsonObject a{{"s", QJsonArray{QJsonObject{{"tpm", 3.5}}}}};
    QJsonObject b{{"s", QJsonArray{QJsonObject{{"tpm", 3.6}}}}};
    const QStringList d = DVE::testutil::diffJson(a, b);
    QCOMPARE(d.size(), 1);
    QVERIFY2(d.first().contains("/s[0]/tpm"), qPrintable(d.first()));
}

void TestV3Harness::jsonDiffReportsMissingKey()
{
    QJsonObject a{{"x", 1}};
    QJsonObject b{{"x", 1}, {"extra", 2}};
    QCOMPARE(DVE::testutil::diffJson(a, b).size(), 1);
}

QTEST_MAIN(TestV3Harness)
#include "tst_v3harness.moc"

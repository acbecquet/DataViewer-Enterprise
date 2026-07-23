#include <QtTest>
#include <QTemporaryDir>
#include "CorpusUtils.h"

class TestV3Harness : public QObject {
    Q_OBJECT
private slots:
    void corpusIncludesFixtures();
    void corpusIncludesEnvDir();
    void corpusSkipsUnsetEnvCleanly();
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

QTEST_MAIN(TestV3Harness)
#include "tst_v3harness.moc"

#include <QtTest>
#include <QTemporaryFile>
#include <QFile>
#include "ZipWriter.h"
#include "TestHelpers.h"

class TestZipWriter : public QObject
{
    Q_OBJECT

private slots:
    void testEmptyArchive();
    void testSingleFile();
    void testMultipleFiles();
    void testStoreMode();
    void testRoundtrip();
    void testCrc32Consistency();
    void testLargeEntry();
};

void TestZipWriter::testEmptyArchive()
{
    ZipWriter zw;
    QByteArray data = zw.toByteArray();
    // A valid ZIP starts with "PK" (0x50 0x4B)
    QVERIFY(data.size() >= 2);
    QVERIFY(data.startsWith("PK"));
}

void TestZipWriter::testSingleFile()
{
    ZipWriter zw;
    zw.addFile("hello.txt", QByteArray("Hello, World!"));
    QByteArray data = zw.toByteArray();
    QVERIFY(data.startsWith("PK"));
    QVERIFY(data.size() > 4);
}

void TestZipWriter::testMultipleFiles()
{
    ZipWriter zw;
    QByteArray fileA("File A content");
    QByteArray fileB("File B content - slightly longer");
    QByteArray fileC("File C");

    zw.addFile("a.txt", fileA);
    zw.addFile("b.txt", fileB);
    zw.addFile("c.txt", fileC);

    QByteArray data = zw.toByteArray();
    QVERIFY(data.startsWith("PK"));
    // ZIP with 3 files should be larger than any single file's content
    QVERIFY(data.size() > fileA.size());
}

void TestZipWriter::testStoreMode()
{
    ZipWriter zw;
    QByteArray content("Store this uncompressed data exactly as-is");
    zw.addFile("stored.txt", content, false);  // compress=false

    QByteArray data = zw.toByteArray();
    QVERIFY(data.startsWith("PK"));
    // In store mode, the raw data should appear verbatim in the archive
    QVERIFY(data.contains(content));
}

void TestZipWriter::testRoundtrip()
{
    ZipWriter zw;
    zw.addFile("test.txt", QByteArray("roundtrip test data"));

    QTemporaryFile tmpFile;
    tmpFile.setAutoRemove(true);
    QVERIFY(tmpFile.open());
    QString path = tmpFile.fileName();
    tmpFile.close();

    bool ok = zw.save(path);
    QVERIFY(ok);

    QFile f(path);
    QVERIFY(f.exists());
    QVERIFY(f.open(QIODevice::ReadOnly));
    QByteArray data = f.readAll();
    f.close();

    // ZIP local file header magic: PK\x03\x04
    QVERIFY(data.size() >= 4);
    QCOMPARE(data[0], 'P');
    QCOMPARE(data[1], 'K');
    QCOMPARE(data[2], '\x03');
    QCOMPARE(data[3], '\x04');
}

void TestZipWriter::testCrc32Consistency()
{
    QByteArray content("identical content for CRC test");

    ZipWriter zw1;
    zw1.addFile("file.txt", content);
    QByteArray data1 = zw1.toByteArray();

    ZipWriter zw2;
    zw2.addFile("file.txt", content);
    QByteArray data2 = zw2.toByteArray();

    QCOMPARE(data1, data2);
}

void TestZipWriter::testLargeEntry()
{
    // 1 MB of repeated data — should compress well
    QByteArray large(1024 * 1024, 'A');

    ZipWriter zw;
    zw.addFile("large.bin", large, true);  // compress=true

    QByteArray data = zw.toByteArray();
    QVERIFY(data.startsWith("PK"));
    // Compressed size should be much smaller than 1 MB
    QVERIFY(data.size() < large.size());
    QVERIFY(data.size() > 0);
}

QTEST_APPLESS_MAIN(TestZipWriter)
#include "tst_zipwriter.moc"

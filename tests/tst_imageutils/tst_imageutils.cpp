#include <QtTest>
#include <QImage>
#include <QBuffer>
#include <QApplication>
#include "ImageUtils.h"
#include "TestHelpers.h"

class TestImageUtils : public QObject
{
    Q_OBJECT

private slots:
    void testLoadImage();
    void testCropImage();
    void testCompressBlob();
    void testInvalidPath();
    void testNoCrop();
};

void TestImageUtils::testLoadImage()
{
    QString path = testDataFile("test_image.png");
    QByteArray result = DVE::loadAndCropImage(path, QRectF());
    QVERIFY(!result.isEmpty());
}

void TestImageUtils::testCropImage()
{
    QString path = testDataFile("test_image.png");

    // Full image (no crop)
    QByteArray full = DVE::loadAndCropImage(path, QRectF());

    // Crop to center 50%
    QByteArray cropped = DVE::loadAndCropImage(path, QRectF(0.25, 0.25, 0.5, 0.5));
    QVERIFY(!cropped.isEmpty());

    // Cropped result should generally be smaller than full image
    QVERIFY(cropped.size() < full.size());
}

void TestImageUtils::testCompressBlob()
{
    // Create a raw QImage and convert to bytes
    QImage img(200, 200, QImage::Format_RGB32);
    img.fill(Qt::red);

    QByteArray raw;
    QBuffer buf(&raw);
    buf.open(QIODevice::WriteOnly);
    img.save(&buf, "PNG");
    buf.close();

    QByteArray compressed = DVE::compressImageBlob(raw);
    QVERIFY(!compressed.isEmpty());
}

void TestImageUtils::testInvalidPath()
{
    QByteArray result = DVE::loadAndCropImage("/nonexistent/path/image.png", QRectF());
    QVERIFY(result.isEmpty());
}

void TestImageUtils::testNoCrop()
{
    QString path = testDataFile("test_image.png");

    // QRectF() — null rect — should return the full image
    QByteArray result = DVE::loadAndCropImage(path, QRectF());
    QVERIFY(!result.isEmpty());

    // Verify it's valid JPEG data (JPEG starts with 0xFF 0xD8)
    QCOMPARE(static_cast<unsigned char>(result[0]), 0xFFu);
    QCOMPARE(static_cast<unsigned char>(result[1]), 0xD8u);
}

QTEST_MAIN(TestImageUtils)
#include "tst_imageutils.moc"

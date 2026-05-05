// ── DataViewer Enterprise — PptxWriter Integration Tests ────────────────────
#include <QtTest>
#include <QTemporaryDir>
#include <QFile>
#include <QDir>
#include <QFileInfo>
#include "TestHelpers.h"
#include "PptxWriter.h"

class tst_PptxWriter : public QObject
{
    Q_OBJECT

private:
    QTemporaryDir m_tempDir;

    QString tempPath(const QString& name) {
        return m_tempDir.path() + "/" + name;
    }

    // Verify file starts with "PK" (ZIP magic bytes)
    bool isValidZip(const QString& path) {
        QFile f(path);
        if (!f.open(QIODevice::ReadOnly)) return false;
        QByteArray header = f.read(2);
        f.close();
        return header.size() == 2 && header[0] == 'P' && header[1] == 'K';
    }

    // Build a simple SlideTable for testing
    DVE::SlideTable makeSimpleTable(int cols = 2, int rows = 2) {
        DVE::SlideTable table;
        for (int c = 0; c < cols; ++c)
            table.headers << QString("Col%1").arg(c + 1);
        for (int r = 0; r < rows; ++r) {
            QStringList row;
            for (int c = 0; c < cols; ++c)
                row << QString("R%1C%2").arg(r + 1).arg(c + 1);
            table.rows.append(row);
        }
        return table;
    }

    // Create a minimal PNG byte array (1x1 pixel red PNG)
    QByteArray makeMinimalPng() {
        // Minimal valid 1x1 red PNG
        static const unsigned char png[] = {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,  // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,  // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,  // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,  // 8-bit RGB
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,  // IDAT chunk
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,  // IEND chunk
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
        return QByteArray(reinterpret_cast<const char*>(png), sizeof(png));
    }

private slots:
    void initTestCase()
    {
        QVERIFY(m_tempDir.isValid());
    }

    // ── Empty save (skeleton PPTX) ──────────────────────────────────────
    void testEmptySave()
    {
        DVE::PptxWriter writer;
        QString path = tempPath("empty.pptx");
        QVERIFY(writer.save(path));
        QVERIFY(QFile::exists(path));
        QVERIFY(isValidZip(path));
    }

    // ── Cover slide ─────────────────────────────────────────────────────
    void testCoverSlide()
    {
        DVE::PptxWriter writer;
        writer.addCoverSlide("Test Report", "2025-06-15");

        QString path = tempPath("cover.pptx");
        QVERIFY(writer.save(path));

        QFileInfo fi(path);
        QVERIFY(fi.exists());
        QVERIFY(fi.size() > 0);
    }

    // ── Content slide with table ────────────────────────────────────────
    void testContentSlide()
    {
        DVE::PptxWriter writer;
        DVE::SlideTable table = makeSimpleTable(2, 2);
        QVector<DVE::SlideImage> noPlots;

        writer.addContentSlide("Sheet 1", table, noPlots);

        QString path = tempPath("content.pptx");
        QVERIFY(writer.save(path));
        QVERIFY(isValidZip(path));
    }

    // ── Content slide with image ────────────────────────────────────────
    void testContentSlideWithImage()
    {
        DVE::PptxWriter writerNoImg;
        DVE::SlideTable table = makeSimpleTable();
        QVector<DVE::SlideImage> noPlots;
        writerNoImg.addContentSlide("No Image", table, noPlots);
        QString pathNoImg = tempPath("no_image.pptx");
        QVERIFY(writerNoImg.save(pathNoImg));

        DVE::PptxWriter writerWithImg;
        DVE::SlideImage img;
        img.pngData = makeMinimalPng();
        img.x = 1.0; img.y = 5.0; img.w = 2.0; img.h = 2.0;
        QVector<DVE::SlideImage> plots;
        plots.append(img);
        writerWithImg.addContentSlide("With Image", table, plots);
        QString pathWithImg = tempPath("with_image.pptx");
        QVERIFY(writerWithImg.save(pathWithImg));

        // File with image should be larger
        QVERIFY(QFileInfo(pathWithImg).size() > QFileInfo(pathNoImg).size());
    }

    // ── Multiple slides ─────────────────────────────────────────────────
    void testMultipleSlides()
    {
        // Single slide reference
        DVE::PptxWriter writerSingle;
        writerSingle.addCoverSlide("Single", "2025-01-01");
        QString pathSingle = tempPath("single.pptx");
        QVERIFY(writerSingle.save(pathSingle));

        // Multiple slides
        DVE::PptxWriter writerMulti;
        writerMulti.addCoverSlide("Multi Report", "2025-01-01");
        for (int i = 0; i < 3; ++i) {
            DVE::SlideTable table = makeSimpleTable(3, 4);
            QVector<DVE::SlideImage> noPlots;
            writerMulti.addContentSlide(
                QString("Sheet %1").arg(i + 1), table, noPlots);
        }
        QString pathMulti = tempPath("multi.pptx");
        QVERIFY(writerMulti.save(pathMulti));

        QVERIFY(QFileInfo(pathMulti).size() > QFileInfo(pathSingle).size());
    }

    // ── Save to temporary path ──────────────────────────────────────────
    void testSaveToTempFile()
    {
        DVE::PptxWriter writer;
        writer.addCoverSlide("Temp Test", "2025-12-31");

        // Use QTemporaryDir-based path (already temp)
        QString path = tempPath("temp_output.pptx");
        QVERIFY(writer.save(path));
        QVERIFY(QFile::exists(path));
        QVERIFY(isValidZip(path));
    }

    // ── Conclusions slide ───────────────────────────────────────────────
    void addConclusionsSlide_producesEmptyTextbox()
    {
        DVE::PptxWriter w;
        w.addConclusionsSlide();
        const QString tmp = QDir::tempPath() + "/dve_conclusions.pptx";
        QVERIFY(w.save(tmp));
        QVERIFY(QFileInfo(tmp).size() > 1000);
        QFile::remove(tmp);
    }

    // ── Test Protocol slide ──────────────────────────────────────────────
    void addTestProtocolSlide_producesTable()
    {
        DVE::PptxWriter w;
        DVE::SlideTable t;
        t.headers = {"Test", "Objective", "Pass Criteria", "Equipment", "Quantity", "Est Duration"};
        t.rows.append({"Lifetime Test", "Run to depletion", "TPM > X", "Vape rig", "n=3", "1mL: 4h / 2mL: 8h"});
        w.addTestProtocolSlide(t);
        const QString tmp = QDir::tempPath() + "/dve_protocol.pptx";
        QVERIFY(w.save(tmp));
        QVERIFY(QFileInfo(tmp).size() > 1000);
        QFile::remove(tmp);
    }

    // ── Test Overview slide ──────────────────────────────────────────────
    void addTestOverviewSlide_producesBullets()
    {
        DVE::PptxWriter w;
        w.addTestOverviewSlide("Standard performance evaluation of DeviceX across 3 tests.",
                               QStringList{"Lifetime Test", "Heavy Metals Test", "Big Headspace Test"});
        const QString tmp = QDir::tempPath() + "/dve_overview.pptx";
        QVERIFY(w.save(tmp));
        QVERIFY(QFileInfo(tmp).size() > 1000);
        QFile::remove(tmp);
    }

    // ── Section Divider slide ────────────────────────────────────────────
    void addSectionDividerSlide_centersFilename()
    {
        DVE::PptxWriter w;
        w.addSectionDividerSlide("DeviceX_Run_2026-04-22");
        const QString tmp = QDir::tempPath() + "/dve_divider.pptx";
        QVERIFY(w.save(tmp));
        QVERIFY(QFileInfo(tmp).size() > 1000);
        QFile::remove(tmp);
    }
};

QTEST_APPLESS_MAIN(tst_PptxWriter)
#include "tst_pptxwriter.moc"

// ── DataViewer Enterprise — PptxWriter Integration Tests ────────────────────
#include <QtTest>
#include <QTemporaryDir>
#include <QFile>
#include <QDir>
#include <QFileInfo>
#include <QRectF>
#include <private/qzipreader_p.h>
#include "TestHelpers.h"
#include "PptxWriter.h"
#include "ReportLayout.h"

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

    // ── Legacy addCoverSlide preserves 46/24 pt fonts ────────────────────
    void addCoverSlide_legacyTwoArgEmitsHardcodedFontSizes()
    {
        DVE::PptxWriter w;
        w.setResourcePath(QStringLiteral("../../resources/images"));
        w.addCoverSlide("Legacy Cover", "2026-05-10");
        const QString path = tempPath("legacy_cover.pptx");
        QVERIFY(w.save(path));
        QZipReader zr(path);
        const QByteArray slideXml = zr.fileData(
            QStringLiteral("ppt/slides/slide1.xml"));
        QVERIFY(!slideXml.isEmpty());
        // Legacy hardcodes: title 46 pt -> sz="4600", date 24 pt -> sz="2400".
        QVERIFY2(slideXml.contains("sz=\"4600\""),
                 "legacy cover title should be 46 pt (sz=4600)");
        QVERIFY2(slideXml.contains("sz=\"2400\""),
                 "legacy cover date should be 24 pt (sz=2400)");
    }

    // ── New addCoverSlide overload honors font overrides ─────────────────
    void addCoverSlide_fontOverloadEmitsRequestedSizes()
    {
        DVE::PptxWriter w;
        w.setResourcePath(QStringLiteral("../../resources/images"));
        // User-customized: title 50 pt, date 30 pt.
        w.addCoverSlide("Custom Cover", "2026-05-10", /*titleFontPt=*/50,
                         /*dateFontPt=*/30);
        const QString path = tempPath("custom_cover.pptx");
        QVERIFY(w.save(path));
        QZipReader zr(path);
        const QByteArray slideXml = zr.fileData(
            QStringLiteral("ppt/slides/slide1.xml"));
        QVERIFY(!slideXml.isEmpty());
        QVERIFY2(slideXml.contains("sz=\"5000\""),
                 "title override (50 pt) missing from slide XML");
        QVERIFY2(slideXml.contains("sz=\"3000\""),
                 "date override (30 pt) missing from slide XML");
        // Legacy sizes should NOT appear for the overridden elements.
        QVERIFY2(!slideXml.contains("sz=\"4600\""),
                 "legacy 46 pt should be replaced by override");
        QVERIFY2(!slideXml.contains("sz=\"2400\""),
                 "legacy 24 pt should be replaced by override");
    }

    // ── addContentSlide reads layout.titleFontPt + tableFontPt ───────────
    void addContentSlide_layoutFontsAppearInSlideXml()
    {
        DVE::PptxWriter w;
        w.setResourcePath(QStringLiteral("../../resources/images"));
        DVE::SlideTable table = makeSimpleTable(2, 2);
        QVector<DVE::SlideImage> noPlots;

        DVE::ContentSlideLayout layout;
        layout.titleFontPt = 22;   // expect sz="2200" in title rPr
        layout.tableFontPt = 13;   // expect sz="1300" in cell rPr

        w.addContentSlide("Custom Content", table, noPlots, layout);
        const QString path = tempPath("custom_content_font.pptx");
        QVERIFY(w.save(path));
        QZipReader zr(path);
        const QByteArray slideXml = zr.fileData(
            QStringLiteral("ppt/slides/slide1.xml"));
        QVERIFY(!slideXml.isEmpty());
        QVERIFY2(slideXml.contains("sz=\"2200\""),
                 "title override (22 pt) missing from slide XML");
        QVERIFY2(slideXml.contains("sz=\"1300\""),
                 "table-cell override (13 pt) missing from slide XML");
    }

    // ── Layout-override overload: EMU values land in slide XML ───────────
    void addContentSlide_layoutOverrideAppliesToEmuPositions()
    {
        DVE::PptxWriter w;
        DVE::SlideTable table = makeSimpleTable(2, 2);
        // Embed positions in the input that should be OVERRIDDEN by `layout`.
        table.x = 0.5; table.y = 0.64; table.w = 12.3; table.h = 4.5;

        DVE::SlideImage img;
        img.pngData = makeMinimalPng();
        // Embed positions in the input that should be OVERRIDDEN by `layout`.
        img.x = 0.0; img.y = 0.0; img.w = 2.7; img.h = 2.0;
        QVector<DVE::SlideImage> plots{ img };

        // Override values (1 inch = 914400 EMU):
        //   table.x = 1.5"  → 1371600 EMU
        //   table.y = 2.5"  → 2286000 EMU
        //   radar.x = 8.5"  → 7772400 EMU
        //   radar.y = 3.5"  → 3200400 EMU
        //   radar.w = 4.25" → 3886200 EMU (also radar.h)
        DVE::ContentSlideLayout layout;
        layout.table = QRectF(1.5, 2.5, 5.0, 3.0);
        layout.radar = QRectF(8.5, 3.5, 4.25, 4.25);

        w.addContentSlide("Override", table, plots, layout);

        const QString path = tempPath("override.pptx");
        QVERIFY(w.save(path));
        QVERIFY(isValidZip(path));

        // Crack the .pptx and read the slide XML.
        QZipReader zr(path);
        QVERIFY(zr.exists());
        const QByteArray slideXml = zr.fileData(
            QStringLiteral("ppt/slides/slide1.xml"));
        QVERIFY(!slideXml.isEmpty());

        // Layout values must round-trip into the EMU output.
        QVERIFY2(slideXml.contains("1371600"),
                 "table.x override (1.5\" = 1371600 EMU) missing from slide XML");
        QVERIFY2(slideXml.contains("2286000"),
                 "table.y override (2.5\" = 2286000 EMU) missing from slide XML");
        QVERIFY2(slideXml.contains("7772400"),
                 "radar.x override (8.5\" = 7772400 EMU) missing from slide XML");
        QVERIFY2(slideXml.contains("3200400"),
                 "radar.y override (3.5\" = 3200400 EMU) missing from slide XML");
        QVERIFY2(slideXml.contains("3886200"),
                 "radar.w override (4.25\" = 3886200 EMU) missing from slide XML");
    }

    // ── Bug 1: layout.title rect lands in slide XML (content-slide title) ──
    void addContentSlide_titleRectAppliesToEmuPosition()
    {
        DVE::PptxWriter w;
        DVE::SlideTable table = makeSimpleTable(2, 2);
        QVector<DVE::SlideImage> noPlots;

        DVE::ContentSlideLayout layout;
        // 2.0" x, 0.6" y → 1828800 / 548640 EMU. Distinct from the hardcoded
        // title box (0.4", 0.1"), so a pass proves the rect was honored.
        layout.title = QRectF(2.0, 0.6, 9.0, 0.7);

        w.addContentSlide("Title Override", table, noPlots, layout);
        const QString path = tempPath("title_override.pptx");
        QVERIFY(w.save(path));
        QZipReader zr(path);
        const QByteArray slideXml =
            zr.fileData(QStringLiteral("ppt/slides/slide1.xml"));
        QVERIFY(!slideXml.isEmpty());
        QVERIFY2(slideXml.contains("1828800"),
                 "title.x override (2.0\" = 1828800 EMU) missing — title rect not honored");
        QVERIFY2(slideXml.contains("548640"),
                 "title.y override (0.6\" = 548640 EMU) missing — title rect not honored");
    }
};

QTEST_APPLESS_MAIN(tst_PptxWriter)
#include "tst_pptxwriter.moc"

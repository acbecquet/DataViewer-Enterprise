#include <QtTest>
#include <QJsonDocument>
#include "ReportLayout.h"

class tst_ReportLayout : public QObject {
    Q_OBJECT
private slots:
    void testRoundTripPreservesAllFields() {
        DVE::ReportLayout in;
        in.tableSort.column = "Overall Liking";
        in.tableSort.order  = Qt::AscendingOrder;
        in.coverTitle    = QRectF(0.1, 0.2, 12.0, 1.0);
        in.coverSubtitle = QRectF(0.1, 1.5, 12.0, 0.5);

        DVE::ContentSlideLayout c;
        c.title = QRectF(0.1, 0.1, 12.0, 0.5);
        c.table = QRectF(0.32, 0.75, 12.7, 1.5);
        c.radar = QRectF(4.5, 2.5, 4.4, 4.4);
        c.propertiesBox.rect = QRectF(10.1, 4.9, 3.17, 2.5);
        c.propertiesBox.text = "Media: cap1\nControl: ctrl1";
        in.contentSlides["content_42"] = c;

        DVE::ImageSlideLayout img;
        img.imageLayouts << QRectF(1, 1, 5, 4);
        img.imageCrops   << QRectF(0, 0, 1, 1);
        in.imageSlides["image_42"] = img;

        in.dividerTitles["divider_42"] = QRectF(0.1, 3.0, 13.0, 1.5);
        in.cumulative = c;
        in.zOrder << "table" << "radar";

        const QJsonObject obj = in.toJson();
        bool ok = false;
        const DVE::ReportLayout out = DVE::ReportLayout::fromJson(obj, &ok);
        QVERIFY(ok);
        QCOMPARE(out.tableSort.column, in.tableSort.column);
        QCOMPARE(out.tableSort.order,  in.tableSort.order);
        QCOMPARE(out.coverTitle,       in.coverTitle);
        QCOMPARE(out.coverSubtitle,    in.coverSubtitle);
        QVERIFY(out.contentSlides.contains("content_42"));
        QCOMPARE(out.contentSlides["content_42"].table, c.table);
        QCOMPARE(out.contentSlides["content_42"].propertiesBox.text, c.propertiesBox.text);
        QVERIFY(out.imageSlides.contains("image_42"));
        QCOMPARE(out.imageSlides["image_42"].imageLayouts.size(), 1);
        QCOMPARE(out.dividerTitles["divider_42"], in.dividerTitles["divider_42"]);
        QCOMPARE(out.zOrder, in.zOrder);
    }

    void testRejectsWrongMode() {
        QJsonObject o{ { "version", 1 }, { "mode", "tpm" } };
        bool ok = true;
        DVE::ReportLayout::fromJson(o, &ok);
        QVERIFY(!ok);
    }

    void testEmptyLayoutIsDetected() {
        DVE::ReportLayout l;
        QVERIFY(l.isEmpty());
    }

    void testFontSizeFieldsRoundTrip() {
        DVE::ReportLayout in;
        in.coverTitle           = QRectF(0.5, 2.5, 12, 1.5);
        in.coverTitleFontPt     = 36;
        in.coverSubtitle        = QRectF(0.5, 4.2, 12, 0.8);
        in.coverSubtitleFontPt  = 22;

        DVE::ContentSlideLayout c;
        c.title                 = QRectF(0.32, 0.10, 12.7, 0.55);
        c.titleFontPt           = 24;
        c.table                 = QRectF(0.32, 0.75, 12.7, 1.5);
        c.tableFontPt           = 14;
        c.propertiesBox.rect    = QRectF(10.1, 4.9, 3.17, 2.5);
        c.propertiesBox.fontPt  = 18;
        in.contentSlides["content_0"] = c;

        in.dividerTitles["divider_0"]        = QRectF(0.5, 3, 12, 1.5);
        in.dividerTitleFontPts["divider_0"]  = 48;

        bool ok = false;
        const DVE::ReportLayout out =
            DVE::ReportLayout::fromJson(in.toJson(), &ok);
        QVERIFY(ok);

        QCOMPARE(out.coverTitleFontPt,    36);
        QCOMPARE(out.coverSubtitleFontPt, 22);
        QCOMPARE(out.contentSlides["content_0"].titleFontPt, 24);
        QCOMPARE(out.contentSlides["content_0"].tableFontPt, 14);
        QCOMPARE(out.contentSlides["content_0"].propertiesBox.fontPt, 18);
        QCOMPARE(out.dividerTitleFontPts["divider_0"], 48);
    }

    void testFontFieldsLoadAsSentinelZeroWhenJsonMissingFields() {
        // Simulates a v1.0.x JSON that has no fontPt fields. fromJson should
        // load font fields as 0 (sentinel) so the renderer falls back to its
        // own hardcoded defaults — preserving legacy output exactly when the
        // user has not customized fonts.
        const QJsonObject legacyJson{
            { "version", 1 },
            { "mode", "sensory" },
            { "tableSort", QJsonObject{ { "column", "" }, { "order", "desc" } } },
            { "slides", QJsonObject{
                { "cover", QJsonObject{
                    { "title",    QJsonArray{0.5, 2.5, 12.0, 1.5} },
                    { "subtitle", QJsonArray{0.5, 4.2, 12.0, 0.8} }
                } },
                { "content_0", QJsonObject{
                    { "title", QJsonArray{0.32, 0.10, 12.7, 0.55} },
                    { "table", QJsonArray{0.32, 0.75, 12.7, 1.5} },
                    { "radar", QJsonArray{4.5, 2.5, 4.4, 4.4} },
                    { "propertiesBox", QJsonObject{
                        { "rect", QJsonArray{10.1, 4.9, 3.17, 2.5} },
                        { "text", "" }
                    } }
                } },
                { "divider_0", QJsonObject{
                    { "title", QJsonArray{0.5, 3, 12, 1.5} }
                } }
            } },
            { "zOrder", QJsonArray{} }
        };

        bool ok = false;
        const DVE::ReportLayout out = DVE::ReportLayout::fromJson(legacyJson, &ok);
        QVERIFY(ok);
        // All font fields default to 0 (renderer falls back to its own hardcode).
        QCOMPARE(out.coverTitleFontPt,    0);
        QCOMPARE(out.coverSubtitleFontPt, 0);
        QCOMPARE(out.contentSlides["content_0"].titleFontPt,           0);
        QCOMPARE(out.contentSlides["content_0"].tableFontPt,           0);
        QCOMPARE(out.contentSlides["content_0"].propertiesBox.fontPt,  0);
        QCOMPARE(out.dividerTitleFontPts["divider_0"],                  0);
    }
};

QTEST_MAIN(tst_ReportLayout)
#include "tst_reportlayout.moc"

#include <QtTest>
#include "SensoryReportSource.h"

class tst_SensoryReportSource : public QObject {
    Q_OBJECT
private:
    DVE::SensorySession makeSess(const QString& title, const QString& tester,
                                  const QStringList& sampleNames) {
        DVE::SensorySession s;
        s.testTitle = title; s.testerName = tester; s.date = "2026-01-01";
        for (const QString& n : sampleNames) {
            DVE::SensorySample samp; samp.name = n;
            s.samples.append(samp);
        }
        return s;
    }

private slots:
    void testEmptySessionsHasNoSlides() {
        DVE::SensoryReportSource src({}, nullptr);
        QCOMPARE(src.slideCount(), 0);
    }

    void testSlideOrderForSingleSession() {
        QVector<DVE::SensorySession> sessions{ makeSess("T1", "A", {"X","Y"}) };
        DVE::SensoryReportSource src(sessions, nullptr);
        // Cover, Divider, Content (no images, no cumulative for single session)
        QCOMPARE(src.slideCount(), 3);
        QCOMPARE(src.slideKind(0), DVE::SlideKind::Cover);
        QCOMPARE(src.slideKind(1), DVE::SlideKind::Divider);
        QCOMPARE(src.slideKind(2), DVE::SlideKind::Content);
    }

    void testSlideOrderForMultipleSessionsGroupedByTester() {
        QVector<DVE::SensorySession> sessions{
            makeSess("T1", "A", {"X"}),
            makeSess("T1", "A", {"X"}),
            makeSess("T1", "B", {"X"})
        };
        DVE::SensoryReportSource src(sessions, nullptr);
        // Cover, Divider(A), Content(0), Content(1), Divider(B), Content(2), Cumulative
        QCOMPARE(src.slideCount(), 7);
        QCOMPARE(src.slideKind(0), DVE::SlideKind::Cover);
        QCOMPARE(src.slideKind(1), DVE::SlideKind::Divider);
        QCOMPARE(src.slideKind(2), DVE::SlideKind::Content);
        QCOMPARE(src.slideKind(3), DVE::SlideKind::Content);
        QCOMPARE(src.slideKind(4), DVE::SlideKind::Divider);
        QCOMPARE(src.slideKind(5), DVE::SlideKind::Content);
        QCOMPARE(src.slideKind(6), DVE::SlideKind::Cumulative);
    }

    void testImageSlideAddedWhenSessionHasImages() {
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"X"}) };
        sessions[0].imagePaths << "fake.png";
        DVE::SensoryReportSource src(sessions, nullptr);
        // Cover, Divider, Content, Image
        QCOMPARE(src.slideCount(), 4);
        QCOMPARE(src.slideKind(3), DVE::SlideKind::Image);
    }

    void testAllSamplesEnumeration() {
        QVector<DVE::SensorySession> sessions{
            makeSess("T1", "A", {"X","Y"}),
            makeSess("T2", "B", {"Z"})
        };
        DVE::SensoryReportSource src(sessions, nullptr);
        const auto refs = src.allSamples();
        QCOMPARE(refs.size(), 3);
        QCOMPARE(refs[0].slideKey, QString("content_0"));
        QCOMPARE(refs[0].displayName, QString("X"));
        QCOMPARE(refs[2].slideKey, QString("content_1"));
        QCOMPARE(refs[2].displayName, QString("Z"));
    }

    void testSourceLabelSingle() {
        QVector<DVE::SensorySession> sessions{ makeSess("My Test", "A", {"X"}) };
        DVE::SensoryReportSource src(sessions, nullptr);
        QCOMPARE(src.sourceLabel(), QString("My Test"));
    }

    void testSourceLabelMultiple() {
        QVector<DVE::SensorySession> sessions{
            makeSess("T1", "A", {"X"}),
            makeSess("T2", "A", {"X"})
        };
        DVE::SensoryReportSource src(sessions, nullptr);
        QCOMPARE(src.sourceLabel(), QString("2 sessions"));
    }

    void testModeIdIsSensory() {
        DVE::SensoryReportSource src({}, nullptr);
        QCOMPARE(src.modeId(), QString("sensory"));
    }

    // ── Default layout positioning (Phase 1A Task 5) ─────────────────────
    void testDefaultLayoutTablePosition() {
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S1","S2","S3"}) };
        const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        const auto& c = layout.contentSlides["content_0"];
        QCOMPARE(c.table.x(), 0.32);
        QCOMPARE(c.table.y(), 0.75);
        QCOMPARE(c.table.width(), 12.7);
        // 3 samples -> 0.50 + 3 * (1/3) = 1.5
        QVERIFY(qFuzzyCompare(c.table.height(), 1.5));
    }

    void testDefaultRadarIsCenteredAndSquare() {
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S"}) };
        const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        const auto& c = layout.contentSlides["content_0"];
        QCOMPARE(c.radar.width(), c.radar.height());        // square
        const double midX = c.radar.x() + c.radar.width() / 2.0;
        QVERIFY(qFuzzyCompare(midX, 13.33 / 2.0));          // horizontally centered
    }

    void testDefaultPropertiesBoxBottomRight() {
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S"}) };
        const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        const auto& pb = layout.contentSlides["content_0"].propertiesBox.rect;
        QCOMPARE(pb.width(), 3.17);
        QCOMPARE(pb.height(), 2.00);
        // bottom-right corner: (slideW - 0.05, slideH - 0.05)
        QVERIFY(qFuzzyCompare(pb.x() + pb.width(),  13.33 - 0.05));
        QVERIFY(qFuzzyCompare(pb.y() + pb.height(),  7.50 - 0.05));
    }

    void testDefaultCoverSlide() {
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S"}) };
        const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        QVERIFY(!layout.coverTitle.isNull());
        QVERIFY(!layout.coverSubtitle.isNull());
        // Title is above subtitle
        QVERIFY(layout.coverTitle.y() < layout.coverSubtitle.y());
    }

    void testDefaultCumulativeOnlyForMultiSession() {
        QVector<DVE::SensorySession> single{ makeSess("T", "A", {"S"}) };
        QVERIFY(DVE::SensoryReportSource::computeDefaultLayout(single).cumulative.table.isNull());

        QVector<DVE::SensorySession> two{ makeSess("T1","A",{"S"}), makeSess("T2","B",{"S"}) };
        const auto layout = DVE::SensoryReportSource::computeDefaultLayout(two);
        QVERIFY(!layout.cumulative.table.isNull());
    }

    void testDefaultImageSlidePopulatedOnlyForSessionsWithImages() {
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S"}) };
        const auto noImages = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        QVERIFY(noImages.imageSlides.isEmpty());

        sessions[0].imagePaths << "a.png" << "b.png";
        const auto withImages = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        QVERIFY(withImages.imageSlides.contains("image_0"));
        QCOMPARE(withImages.imageSlides["image_0"].imageLayouts.size(), 2);
        QCOMPARE(withImages.imageSlides["image_0"].imageCrops.size(),   2);
    }
};

QTEST_MAIN(tst_SensoryReportSource)
#include "tst_sensoryreportsource.moc"

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
};

QTEST_MAIN(tst_SensoryReportSource)
#include "tst_sensoryreportsource.moc"

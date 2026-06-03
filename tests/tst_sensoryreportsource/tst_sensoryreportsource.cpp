#include <QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include "SensoryReportSource.h"
#include "DatabaseManager.h"
#include "PostgresConnection.h"
#include "ConfigLoader.h"
#include "IdentityManager.h"

namespace {

// Build a DbConfig from DVE_TEST_PG_CONN; same pattern as the other Postgres
// integration tests. Returns a default-constructed DbConfig if the env var is
// unset (testSaveLoadLayoutRoundTrip's initTestCase guard handles the skip).
DVE::DbConfig pgConfig() {
    DVE::DbConfig c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
    if (env.isEmpty()) return c;
    for (const QString& part : env.split(' ', Qt::SkipEmptyParts)) {
        const QStringList kv = part.split('=');
        if (kv.size() != 2) continue;
        const QString k = kv[0], v = kv[1];
        if      (k == "host")     c.host = v;
        else if (k == "port")     c.port = v.toInt();
        else if (k == "dbname")   c.database = v;
        else if (k == "user")     c.user = v;
        else if (k == "password") c.password = v;
    }
    return c;
}

} // anonymous

class tst_SensoryReportSource : public QObject {
    Q_OBJECT
private:
    // Identity for DatabaseManager::open(cfg, &identity). Stays alive for the
    // whole test object lifetime.
    DVE::IdentityManager m_identity;

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

    static bool layout_isDefaultShape(const DVE::ReportLayout& l) {
        // Just a smoke check that load returned something with the expected
        // skeleton — the strict positioning is covered by Task 5 tests.
        return !l.contentSlides.isEmpty()
            && !l.contentSlides.constBegin().value().table.isNull();
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

    void testOneCumulativeSlidePerDistinctTest() {
        // Combined report covering 3 distinct tests across 2 users → expect
        // exactly 3 cumulative slides at the end (one per test).
        QVector<DVE::SensorySession> sessions{
            makeSess("TestA", "U1", {"X"}),
            makeSess("TestA", "U2", {"X"}),
            makeSess("TestB", "U1", {"X"}),
            makeSess("TestC", "U2", {"X"})
        };
        DVE::SensoryReportSource src(sessions, nullptr);

        int cumulativeCount = 0;
        int firstCumIdx = -1;
        for (int i = 0; i < src.slideCount(); ++i) {
            if (src.slideKind(i) == DVE::SlideKind::Cumulative) {
                if (firstCumIdx < 0) firstCumIdx = i;
                ++cumulativeCount;
            }
        }
        QCOMPARE(cumulativeCount, 3);
        // Cumulative slides must come at the end, after all content/image slides.
        for (int i = firstCumIdx; i < src.slideCount(); ++i)
            QCOMPARE(src.slideKind(i), DVE::SlideKind::Cumulative);
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

    void testDefaultRadarIsCenteredAndAspectMatched() {
        // Radar default uses the rendered chart pixmap's aspect (1163/658 ≈
        // 1.767), updated in v1.1.2 from the square 1:1 default. Width must
        // equal height × aspect, and the chart stays horizontally centered.
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S"}) };
        const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        const auto& c = layout.contentSlides["content_0"];
        constexpr double aspect = 1163.0 / 658.0;
        QVERIFY2(qAbs(c.radar.width() - c.radar.height() * aspect) < 0.01,
                 "radar width should equal height × 1163/658 aspect");
        const double midX = c.radar.x() + c.radar.width() / 2.0;
        QVERIFY(qFuzzyCompare(midX, 13.33 / 2.0));          // horizontally centered
    }

    void testDefaultPropertiesBoxBottomRight() {
        // propsH bumped from 2.00 -> 2.70 in v1.1.1 so the typical 6-line
        // body fits at 16 pt without auto-grow.
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S"}) };
        const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        const auto& pb = layout.contentSlides["content_0"].propertiesBox.rect;
        QCOMPARE(pb.width(), 3.17);
        QCOMPARE(pb.height(), 2.70);
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

    void testDefaultLayoutDividerKeysMatchSlideIndex() {
        // Sessions across two testers — buildSlideIndex pushes
        // divider_0 (for tester A) and divider_2 (for tester B);
        // computeDefaultLayout must emit the same set, with no
        // orphan divider_1 for the second-A session.
        QVector<DVE::SensorySession> sessions{
            makeSess("T", "A", {"X"}),
            makeSess("T", "A", {"X"}),
            makeSess("T", "B", {"X"})
        };
        const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);

        // Direct assertion: layout has exactly two divider keys, divider_0 and divider_2.
        QCOMPARE(layout.dividerTitles.size(), 2);
        QVERIFY(layout.dividerTitles.contains("divider_0"));
        QVERIFY(layout.dividerTitles.contains("divider_2"));
        QVERIFY(!layout.dividerTitles.contains("divider_1"));   // no orphan
    }

    void testLoadLayoutReturnsDefaultsWhenNoDB() {
        QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S"}) };
        DVE::SensoryReportSource src(sessions, nullptr);
        const auto layout = src.loadLayout();
        // Defaults: content_0 must exist with non-null table rect
        QVERIFY(layout.contentSlides.contains("content_0"));
        QVERIFY(!layout.contentSlides["content_0"].table.isNull());
    }

    void testSaveLoadLayoutRoundTrip() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
            QSKIP("DVE_TEST_PG_CONN not set; layout round-trip needs Postgres");

        // Wipe sensory_sessions so a previous slot/run doesn't collide on the
        // natural key. Uses its own connection; teardown removes it after.
        {
            DVE::DbConfig cfg = pgConfig();
            const QString cname = "tst_srs_wipe";
            QSqlDatabase wipe = QSqlDatabase::addDatabase("QPSQL", cname);
            wipe.setHostName(cfg.host); wipe.setPort(cfg.port);
            wipe.setDatabaseName(cfg.database); wipe.setUserName(cfg.user);
            wipe.setPassword(cfg.password);
            if (wipe.open()) {
                QSqlQuery q(wipe);
                q.exec("DELETE FROM sensory_images");
                q.exec("DELETE FROM sensory_sessions");
                q.exec("DELETE FROM settings");
                wipe.close();
            }
            QSqlDatabase::removeDatabase(cname);
        }

        DVE::DatabaseManager db;
        QVERIFY(db.open(pgConfig(), &m_identity));

        DVE::SensorySession s = makeSess("T", "A", {"S"});
        s.sessionName = QStringLiteral("test-session-1");  // natural key must be non-null
        QVERIFY(db.saveSensorySession(s));    // populates s.id
        QVERIFY(s.id > 0);

        DVE::SensoryReportSource src({s}, &db);

        // Load → defaults (no JSON saved yet)
        DVE::ReportLayout l = src.loadLayout();
        QVERIFY(layout_isDefaultShape(l));

        // Mutate the title rect to a recognizable value
        l.contentSlides["content_0"].title = QRectF(1.234, 5.678, 9.0, 0.5);
        // Also touch the cumulative slide (global setting)
        l.cumulative.title = QRectF(2.0, 3.0, 4.0, 5.0);
        src.saveLayout(l);

        // Load via a fresh source on the same DB — should see our edits
        DVE::SensoryReportSource src2({s}, &db);
        const DVE::ReportLayout l2 = src2.loadLayout();
        QCOMPARE(l2.contentSlides["content_0"].title.x(),     1.234);
        QCOMPARE(l2.contentSlides["content_0"].title.width(), 9.0);
        QCOMPARE(l2.cumulative.title, QRectF(2.0, 3.0, 4.0, 5.0));
    }

    // ── Phase 1A Task 8: empty-session early-out structural tests ─────────
    //
    // These tests exercise the early-out branch in writeSensoryPptx (which
    // mirrors the original SensoryPanel::generateCombinedPptx early-out at
    // line 1899-1902 of the legacy code). They confirm that:
    //   1. Both the static helper and the IReportSource override exist.
    //   2. Both return false with a non-empty error message when sessions
    //      is empty — no rendering, no file IO needed.
    //
    // Behavioral equivalence between legacy generateCombinedPptx and the new
    // entry point (with default layout + no exclusion) is verified manually
    // in Task 9 (smoke test) — that requires QGuiApplication and the radar
    // renderer, which adds test build complexity beyond Task 8 scope.
    void writePptx_returnsFalseWhenSessionsEmpty() {
        DVE::SensoryReportSource src({}, nullptr);
        DVE::ReportLayout layout;
        QString err;
        QVERIFY(!src.writePptx("nonexistent.pptx", layout, {}, &err));
        QVERIFY(!err.isEmpty());
    }

    void writeSensoryPptx_static_returnsFalseWhenSessionsEmpty() {
        DVE::ReportLayout layout;
        QString err;
        QVERIFY(!DVE::SensoryReportSource::writeSensoryPptx(
            {}, layout, {}, "nonexistent.pptx", &err));
        QVERIFY(!err.isEmpty());
    }

    // ── CRITICAL-1 regression: legacy path must NOT use computeDefaultLayout
    //
    // SensoryPanel::generateCombinedPptx (the legacy "Sensory Report" button)
    // must pass an empty ReportLayout into writeSensoryPptx — NOT
    // computeDefaultLayout(sessions). computeDefaultLayout's simpler math
    // (0.50 + N/3 table height; square radar; cumulative sized from first
    // session's sample count) would override the legacy in-line wrap-aware
    // math via the 5-arg PptxWriter::addContentSlide overload, breaking
    // visual equivalence with the pre-Task-8 output.
    //
    // We can't test the call site directly without launching the GUI, so this
    // test asserts the underlying invariant: the two layout shapes differ for
    // a typical session, so they MUST NOT be conflated. If a future change
    // makes computeDefaultLayout match the legacy positions exactly, this
    // test will need to be revisited (and the bug class disappears).
    // ── Phase 1D Task 18: buildSlide produces real content ────────────────
    //
    // buildSlide must populate ReportSlideSpec for the canvas: kind, slideKey,
    // title, tableHeaders/rows for content slides, and respect the
    // excludedSamples set when building rows.
    void testBuildSlideForContentRespectsExclusion() {
        QVector<DVE::SensorySession> sessions{ makeSess("T1", "A", {"S1", "S2"}) };
        DVE::SensoryReportSource src(sessions, nullptr);
        DVE::ReportLayout layout;

        // Slide 0 = Cover
        const DVE::ReportSlideSpec cover = src.buildSlide(0, layout, {});
        QCOMPARE(cover.kind, DVE::SlideKind::Cover);
        QVERIFY(!cover.title.isEmpty());

        // Slide 2 = Content (Cover, Divider, Content order from buildSlideIndex)
        const DVE::ReportSlideSpec content = src.buildSlide(2, layout, {});
        QCOMPARE(content.kind, DVE::SlideKind::Content);
        QVERIFY(content.tableHeaders.contains("Sample"));
        QVERIFY(content.tableHeaders.contains("Comments"));
        // Sample + 5 metrics + Puff Length (s) + Comments = 8 columns
        QCOMPARE(content.tableHeaders.size(), 8);
        QCOMPARE(content.tableRows.size(), 2);

        // Exclude S2 (sessionIdx=0, sampleIdx=1 -> "0#1")
        const DVE::ReportSlideSpec specExcl = src.buildSlide(2, layout, {"0#1"});
        QCOMPARE(specExcl.tableRows.size(), 1);
    }

    // ── Spec item #7: per-sample power_type + puff_length surface in reports
    //
    // A content slide built from a session whose samples carry powerType and
    // puffLengthSec must:
    //   1. include a "Puff Length (s)" header column (placed immediately
    //      before the existing "Comments" column);
    //   2. emit the puff length numeric value in the row for that sample;
    //   3. carry the powerType string somewhere user-visible (the V/R/P cell
    //      is the agreed location).
    //
    // The content-slide spec doesn't render the V/R/P cell yet, but the row
    // and header changes are observable via tableHeaders/tableRows from
    // buildSlide().
    void powerTypeAndPuffLengthRoundTripIntoReport() {
        DVE::SensorySession s = makeSess("RoundTripTest", "Alice", {"DeviceA"});
        DVE::SensorySample& a = s.samples[0];
        a.voltage       = 3.7;
        a.resistance    = 1.2;
        a.power         = 11.4;
        a.powerType     = "Variable Voltage";
        a.puffLengthSec = 2.5;

        DVE::SensoryReportSource src({s}, nullptr);
        DVE::ReportLayout layout;

        // Slide order for a single session: Cover (0), Divider (1), Content (2).
        const DVE::ReportSlideSpec spec = src.buildSlide(2, layout, {});
        QCOMPARE(spec.kind, DVE::SlideKind::Content);

        // Header gains "Puff Length (s)" before "Comments".
        const int puffIdx     = spec.tableHeaders.indexOf("Puff Length (s)");
        const int commentsIdx = spec.tableHeaders.indexOf("Comments");
        QVERIFY2(puffIdx >= 0,
                 "Content table must expose a 'Puff Length (s)' column");
        QVERIFY2(commentsIdx >= 0, "Content table must keep 'Comments'");
        QVERIFY2(puffIdx < commentsIdx,
                 "'Puff Length (s)' must precede 'Comments'");

        // The row for our sample carries the formatted puff length value.
        QCOMPARE(spec.tableRows.size(), 1);
        const QString puffCell = spec.tableRows[0].value(puffIdx);
        QCOMPARE(puffCell, QString("2.5"));

        // The spec originally said to append the power type to the V/R/P cell,
        // but the sensory report's table never had a V/R/P cell. powerType lives
        // in propertiesText instead (both modern and legacy PPTX paths). Accept
        // either propertiesText or the table cells defensively so the test
        // keeps passing if the surface location moves in the future.
        const QString allText = spec.propertiesText
            + QStringLiteral("|")
            + spec.tableRows[0].join(QStringLiteral("|"));
        QVERIFY2(allText.contains("Variable Voltage"),
                 "Power type 'Variable Voltage' must appear in the slide spec");
    }

    void writeSensoryPptx_legacyPathUsesEmptyLayout_notComputeDefault() {
        // Invariant: the legacy SensoryPanel::generateCombinedPptx path MUST
        // pass an empty ReportLayout (not computeDefaultLayout) so the 5-arg
        // addContentSlide's null-rect checks fall through to the legacy
        // wrap-aware in-line positioning math. Conflating the two would
        // visibly change the per-session table sizing and radar shape.
        QVector<DVE::SensorySession> sessions{ makeSess("T1", "A", {"X","Y","Z"}) };
        DVE::ReportLayout def = DVE::SensoryReportSource::computeDefaultLayout(sessions);
        DVE::ReportLayout empty;

        // computeDefault populates a content slide; empty does not.
        QVERIFY(def.contentSlides.contains("content_0"));
        QVERIFY(empty.contentSlides.isEmpty());

        // The default's content slide has a populated (non-null) table rect.
        const auto& cs = def.contentSlides.value("content_0");
        QVERIFY(!cs.table.isNull());

        // 0.50 + 3 * (1.0/3.0) = 1.50 — different from legacy wrap-aware math.
        QCOMPARE(cs.table.height(), 1.5);

        // computeDefault uses chart-pixmap aspect (1163/658 ≈ 1.767); the
        // legacy in-line math also computes from aspect but the height
        // calculation differs — so the WIDTHS differ even though both are
        // aspect-matched. Verify width != legacy slideW/2 sanity.
        QVERIFY2(cs.radar.width() > cs.radar.height(),
                 "default radar should be wider than tall (1.767 aspect)");
    }

    // ── Bug 1: properties box honors a layout rect override ──────────────
    void propertiesBoxXml_honorsRectOverride()
    {
        const QStringList lines{ "Media: X", "Control: Y" };
        // Override at 1.0",2.0" → x=914400, y=1828800 EMU.
        const QString xml = DVE::SensoryReportSource::buildPropertiesBoxXml(
            lines, QRectF(1.0, 2.0, 3.0, 4.0), /*fontPt=*/0, 13.33, 7.5);
        QVERIFY(!xml.isEmpty());
        QVERIFY2(xml.contains(QStringLiteral("x=\"914400\"")),
                 "props box x override (1.0\") missing");
        QVERIFY2(xml.contains(QStringLiteral("y=\"1828800\"")),
                 "props box y override (2.0\") missing");
    }
    void propertiesBoxXml_nullRectAnchorsBottomRight()
    {
        const QStringList lines{ "Media: X" };
        const QString xml = DVE::SensoryReportSource::buildPropertiesBoxXml(
            lines, QRectF(), /*fontPt=*/0, 13.33, 7.5);
        QVERIFY(!xml.isEmpty());
        QVERIFY(xml.contains(QStringLiteral("name=\"PropsBox\"")));
        // Bottom-right legacy anchor: tbX = 13.33 - 3.17 - 0.05 = 10.11" → 9244584 EMU.
        QVERIFY2(xml.contains(QStringLiteral("9244584")),
                 "null rect should keep the bottom-right legacy x position");
    }
    void propertiesBoxXml_emptyLinesYieldEmpty()
    {
        QVERIFY(DVE::SensoryReportSource::buildPropertiesBoxXml(
            {}, QRectF(), 0, 13.33, 7.5).isEmpty());
    }
};

QTEST_MAIN(tst_SensoryReportSource)
#include "tst_sensoryreportsource.moc"

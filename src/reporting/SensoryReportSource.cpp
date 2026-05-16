#include "SensoryReportSource.h"
#include "database/DatabaseManager.h"
#include "PptxWriter.h"
#include "../ui/RadarChartWidget.h"
#include "../utils/ImageUtils.h"
#include <QApplication>
#include <QBuffer>
#include <QDate>
#include <QDir>
#include <QHash>
#include <QJsonArray>
#include <QJsonDocument>
#include <QMap>
#include <QPainter>
#include <QPixmap>
#include <algorithm>

namespace DVE {

namespace {

// Duplicated from src/ui/SensoryPanel.cpp (findResourcePath there).
// Keeps the refactor surgical and avoids touching the utils tree.
const QString& findResourcePath()
{
    // Cache the result -- filesystem probing is not free and the answer
    // never changes during a single process lifetime.
    static const QString cached = []() -> QString {
        const QStringList candidates = {
            QApplication::applicationDirPath() + "/resources",
            QApplication::applicationDirPath() + "/../resources",
            "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer/resources",
            "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/resources"
        };
        for (const QString& p : candidates) {
            if (QDir(p).exists()) return p;
        }
        return candidates.first();
    }();
    return cached;
}

} // namespace

SensoryReportSource::SensoryReportSource(QVector<SensorySession> sessions,
                                          DatabaseManager* db)
    : m_sessions(std::move(sessions)), m_db(db)
{
    buildSlideIndex();
}

void SensoryReportSource::buildSlideIndex()
{
    m_slides.clear();
    if (m_sessions.isEmpty()) return;

    // Cover slide
    m_slides.push_back({ SlideKind::Cover, -1, QStringLiteral("cover") });

    // Group sessions by tester (mirrors generateCombinedPptx grouping)
    QHash<QString, QVector<int>> byTester;
    QStringList testerOrder;
    for (int i = 0; i < m_sessions.size(); ++i) {
        const QString& t = m_sessions[i].testerName;
        if (!byTester.contains(t)) testerOrder.append(t);
        byTester[t].append(i);
    }
    for (const QString& tester : testerOrder) {
        const QVector<int>& idxs = byTester[tester];
        m_slides.push_back({ SlideKind::Divider, idxs.first(),
                              QStringLiteral("divider_%1").arg(idxs.first()) });
        for (int sIdx : idxs) {
            m_slides.push_back({ SlideKind::Content, sIdx,
                                  QStringLiteral("content_%1").arg(sIdx) });
            if (!m_sessions[sIdx].imagePaths.isEmpty())
                m_slides.push_back({ SlideKind::Image, sIdx,
                                      QStringLiteral("image_%1").arg(sIdx) });
        }
    }

    // Cumulative summary: one slide per distinct test (testTitle), so a
    // combined report covering N tests gets N cumulative slides at the end.
    // Gated on 2+ total sessions to preserve the prior "nothing to aggregate"
    // behavior for trivial reports. sessionIdx carries the first session of
    // each test group so buildSlide can recover the testTitle filter; the key
    // stays "cumulative" so layout overrides and persistence remain global.
    if (m_sessions.size() >= 2) {
        QHash<QString, int> firstIdxByTest;
        QStringList testOrder;
        for (int i = 0; i < m_sessions.size(); ++i) {
            const QString title = m_sessions[i].testTitle.isEmpty()
                ? QStringLiteral("Sensory Evaluation")
                : m_sessions[i].testTitle;
            if (!firstIdxByTest.contains(title)) {
                firstIdxByTest.insert(title, i);
                testOrder.append(title);
            }
        }
        for (const QString& title : testOrder) {
            m_slides.push_back({ SlideKind::Cumulative,
                                  firstIdxByTest.value(title),
                                  QStringLiteral("cumulative") });
        }
    }
}

QString SensoryReportSource::sourceLabel() const
{
    if (m_sessions.size() == 1)
        return m_sessions.first().testTitle;
    return QStringLiteral("%1 sessions").arg(m_sessions.size());
}

int SensoryReportSource::slideCount() const { return m_slides.size(); }
SlideKind SensoryReportSource::slideKind(int idx) const { return m_slides[idx].kind; }
QString   SensoryReportSource::slideKey(int idx) const  { return m_slides[idx].key; }

QVector<SampleRef> SensoryReportSource::allSamples() const
{
    QVector<SampleRef> out;
    for (int i = 0; i < m_sessions.size(); ++i) {
        const auto& sess = m_sessions[i];
        const QString slideKey = QStringLiteral("content_%1").arg(i);
        for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
            const auto& samp = sess.samples[sIdx];
            SampleRef r;
            r.slideKey = slideKey;
            r.sampleId = QStringLiteral("%1#%2").arg(i).arg(sIdx);
            r.displayName = samp.name.isEmpty()
                ? QStringLiteral("Sample %1").arg(sIdx + 1) : samp.name;
            // sessionLabel must uniquely identify the session in the checkbox
            // panel — when multiple testers rate the same product, they share
            // the testTitle, so we append the tester name to avoid collapsing
            // them into one group with duplicate-looking sample names.
            QString label = sess.testTitle.isEmpty()
                ? QStringLiteral("Session %1").arg(i + 1) : sess.testTitle;
            if (!sess.testerName.isEmpty())
                label += QStringLiteral(" - ") + sess.testerName;
            r.sessionLabel = label;
            out.append(r);
        }
    }
    return out;
}

// ── Shared content-builder helpers ───────────────────────────────────────
//
// These extract the data (headers, rows, radar QImage, properties text)
// that's identical between the canvas (buildSlide) and PPTX (writeSensoryPptx)
// paths. Layout/positioning stays inline in writeSensoryPptx to preserve
// pre-Task-8 behavioral equivalence.

void SensoryReportSource::buildContentTable(const SensorySession& sess,
                                              int sessionIdx,
                                              const QSet<QString>& excludedSamples,
                                              QStringList& outHeaders,
                                              QVector<QStringList>& outRows)
{
    outHeaders.clear();
    outRows.clear();
    outHeaders << QStringLiteral("Sample");
    for (const QString& m : kSensoryMetrics) outHeaders << m;
    outHeaders << QStringLiteral("Puff Length (s)");
    outHeaders << QStringLiteral("Comments");

    for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
        const QString sampleKey = QStringLiteral("%1#%2").arg(sessionIdx).arg(sIdx);
        if (excludedSamples.contains(sampleKey)) continue;
        const SensorySample& s = sess.samples[sIdx];
        QStringList row;
        row << (s.name.isEmpty() ? QStringLiteral("Sample") : s.name);
        for (const QString& m : kSensoryMetrics) {
            double v = s.scores.value(m, 5.0);
            row << ((v == qRound(v)) ? QString::number(qRound(v))
                                      : QString::number(v, 'f', 1));
        }
        row << QString::number(s.puffLengthSec, 'f', 1);
        row << s.comments;
        outRows.append(row);
    }
}

QImage SensoryReportSource::renderRadarImage(const SensorySession& sess,
                                               int sessionIdx,
                                               const QSet<QString>& excludedSamples)
{
    SensorySession filteredSess = sess;
    filteredSess.samples.clear();
    for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
        const QString sampleKey = QStringLiteral("%1#%2").arg(sessionIdx).arg(sIdx);
        if (excludedSamples.contains(sampleKey)) continue;
        filteredSess.samples.append(sess.samples[sIdx]);
    }

    RadarChartWidget tempChart;
    tempChart.setSessions({filteredSess});
    tempChart.setReportMode(true);
    tempChart.setReportCropTop(70);
    tempChart.setReportCropBottom(130);   // matches the crop amount below so
                                           // drawAxisLabels can clamp labels
                                           // into the visible region
    tempChart.resize(1163, 858);
    // Force white background so the chart sits on a clean canvas in both the
    // preview and the PPTX. Without this, the widget's palette can paint a
    // light-gray window-color fill that bleeds into the rendered pixmap.
    QPixmap pixmap(1163, 858);
    pixmap.fill(Qt::white);
    QPainter painter(&pixmap);
    tempChart.render(&painter);
    painter.end();

    int cropTop = 70;
    int cropBottom = 130;
    QPixmap cropped = pixmap.copy(0, cropTop, 1163, 858 - cropTop - cropBottom);
    return cropped.toImage();
}

QString SensoryReportSource::buildPropertiesText(const SensorySession& sess,
                                                   int sessionIdx,
                                                   const QSet<QString>& excludedSamples)
{
    QStringList propLines;
    if (!sess.media.isEmpty())
        propLines << "Media: " + sess.media;
    if (!sess.control.isEmpty())
        propLines << "Control: " + sess.control;
    propLines << QString("Blind: %1").arg(sess.isBlind ? "Yes" : "No");
    if (!sess.primaryDifferences.isEmpty())
        propLines << "Primary Difference(s): " + sess.primaryDifferences;

    // #7: surface per-sample power type when the session uses a single type
    // across all included samples. If samples disagree, list them.
    {
        QSet<QString> ptypes;
        for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
            const QString sampleKey = QStringLiteral("%1#%2").arg(sessionIdx).arg(sIdx);
            if (excludedSamples.contains(sampleKey)) continue;
            const QString pt = sess.samples[sIdx].powerType.trimmed();
            if (!pt.isEmpty()) ptypes.insert(pt);
        }
        if (ptypes.size() == 1) {
            propLines << "Power Type: " + *ptypes.constBegin();
        } else if (ptypes.size() > 1) {
            QStringList sorted(ptypes.begin(), ptypes.end());
            std::sort(sorted.begin(), sorted.end());
            propLines << "Power Type: " + sorted.join(", ");
        }
    }

    // Highest / Lowest Rated by Overall Liking — with exclusion applied
    SensorySession filteredSess = sess;
    filteredSess.samples.clear();
    for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
        const QString sampleKey = QStringLiteral("%1#%2").arg(sessionIdx).arg(sIdx);
        if (excludedSamples.contains(sampleKey)) continue;
        filteredSess.samples.append(sess.samples[sIdx]);
    }
    if (!filteredSess.samples.isEmpty()) {
        double maxScore = -1.0, minScore = 10.0;
        QStringList maxNames, minNames;
        for (const SensorySample& samp : filteredSess.samples) {
            double score = samp.scores.value("Overall Liking", -1.0);
            if (score < 0) continue;
            if (score > maxScore) { maxScore = score; maxNames.clear(); maxNames << samp.name; }
            else if (qAbs(score - maxScore) < 0.01) maxNames << samp.name;
            if (score < minScore) { minScore = score; minNames.clear(); minNames << samp.name; }
            else if (qAbs(score - minScore) < 0.01) minNames << samp.name;
        }
        if (!maxNames.isEmpty())
            propLines << "Highest Rated: " + maxNames.join(", ");
        if (!minNames.isEmpty())
            propLines << "Lowest Rated: " + minNames.join(", ");
    }
    return propLines.join('\n');
}

void SensoryReportSource::buildCumulativeData(const QVector<SensorySession>& sessions,
                                                const QSet<QString>& excludedSamples,
                                                QStringList& outHeaders,
                                                QVector<QStringList>& outRows,
                                                SensorySession& outCumSession,
                                                const QString& filterTestTitle)
{
    outHeaders.clear();
    outRows.clear();
    outCumSession = SensorySession();

    auto effectiveTitle = [](const SensorySession& s) {
        return s.testTitle.isEmpty()
            ? QStringLiteral("Sensory Evaluation") : s.testTitle;
    };

    struct DeviceAccum { QMap<QString, double> sums; int count = 0; };
    QMap<QString, DeviceAccum> deviceMap;
    QStringList deviceOrder;

    // Track the first matching session so outCumSession.testTitle reflects
    // the group when a filter is applied (not just sessions.first()).
    int firstMatchIdx = -1;

    for (int sessIdx = 0; sessIdx < sessions.size(); ++sessIdx) {
        const SensorySession& sess = sessions[sessIdx];
        if (!filterTestTitle.isEmpty() && effectiveTitle(sess) != filterTestTitle)
            continue;
        if (firstMatchIdx < 0) firstMatchIdx = sessIdx;
        for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
            const QString sampleKey = QStringLiteral("%1#%2").arg(sessIdx).arg(sIdx);
            if (excludedSamples.contains(sampleKey)) continue;
            const SensorySample& s = sess.samples[sIdx];
            QString key = s.name.isEmpty() ? QStringLiteral("Sample") : s.name;
            if (!deviceMap.contains(key)) deviceOrder.append(key);
            DeviceAccum& acc = deviceMap[key];
            for (const QString& m : kSensoryMetrics)
                acc.sums[m] += s.scores.value(m, 5.0);
            acc.count++;
        }
    }

    if (firstMatchIdx >= 0)
        outCumSession.testTitle = effectiveTitle(sessions[firstMatchIdx]);
    else if (!sessions.isEmpty())
        outCumSession.testTitle = effectiveTitle(sessions.first());

    outHeaders << QStringLiteral("Device");
    for (const QString& m : kSensoryMetrics) outHeaders << m;

    for (const QString& devName : deviceOrder) {
        const DeviceAccum& acc = deviceMap[devName];
        SensorySample avgSample;
        avgSample.name = devName;
        QStringList row;
        row << devName;
        for (const QString& m : kSensoryMetrics) {
            double avg = acc.sums.value(m, 0) / qMax(1, acc.count);
            avgSample.scores[m] = avg;
            row << QString::number(avg, 'f', 1);
        }
        outCumSession.samples.append(avgSample);
        outRows.append(row);
    }
}

QImage SensoryReportSource::renderCumulativeRadarImage(const SensorySession& cumSession)
{
    RadarChartWidget cumChart;
    cumChart.setSessions({cumSession});
    cumChart.setReportMode(true);
    cumChart.setReportCropTop(70);
    cumChart.setReportCropBottom(130);
    cumChart.resize(1163, 858);
    QPixmap cumPix(1163, 858);
    cumPix.fill(Qt::white);
    QPainter cumPainter(&cumPix);
    cumChart.render(&cumPainter);
    cumPainter.end();

    int cumCropTop = 70;
    int cumCropBottom = 130;
    QPixmap cumCropped = cumPix.copy(0, cumCropTop, 1163, 858 - cumCropTop - cumCropBottom);
    return cumCropped.toImage();
}

// ── buildSlide: data + resolved layout for the canvas ────────────────────
ReportSlideSpec SensoryReportSource::buildSlide(int idx, const ReportLayout& layoutIn,
                                                  const QSet<QString>& excludedSamples) const
{
    ReportSlideSpec spec;
    if (idx < 0 || idx >= m_slides.size()) return spec;
    const SlideEntry& entry = m_slides[idx];
    spec.kind = entry.kind;
    spec.slideKey = entry.key;

    // Use the supplied layout if it has data; fall back to defaults if empty.
    const ReportLayout effective = layoutIn.isEmpty()
        ? computeDefaultLayout(m_sessions) : layoutIn;

    switch (entry.kind) {
        case SlideKind::Cover: {
            spec.title = m_sessions.isEmpty()
                ? QStringLiteral("Sensory Evaluation")
                : (m_sessions.first().testTitle.isEmpty()
                    ? QStringLiteral("Sensory Evaluation")
                    : m_sessions.first().testTitle);
            spec.layout.title = effective.coverTitle;
            // Use propertiesText to carry the date string for the canvas.
            spec.propertiesText = m_sessions.isEmpty()
                ? QDate::currentDate().toString("MMMM d, yyyy")
                : (m_sessions.first().date.isEmpty()
                    ? QDate::currentDate().toString("MMMM d, yyyy")
                    : m_sessions.first().date);
            break;
        }
        case SlideKind::Divider: {
            spec.title = (entry.sessionIdx >= 0 && entry.sessionIdx < m_sessions.size())
                ? m_sessions[entry.sessionIdx].testerName
                : QString();
            spec.layout.title = effective.dividerTitles.value(entry.key);
            break;
        }
        case SlideKind::Content: {
            const SensorySession& sess = m_sessions[entry.sessionIdx];
            QString title = sess.testTitle.isEmpty()
                ? QStringLiteral("Sensory Evaluation") : sess.testTitle;
            if (!sess.testerName.isEmpty()) title += " - " + sess.testerName;
            spec.title = title;
            buildContentTable(sess, entry.sessionIdx, excludedSamples,
                              spec.tableHeaders, spec.tableRows);
            spec.radarPixmap = renderRadarImage(sess, entry.sessionIdx, excludedSamples);
            spec.propertiesText = buildPropertiesText(sess, entry.sessionIdx, excludedSamples);
            spec.layout = effective.contentSlides.value(entry.key);
            break;
        }
        case SlideKind::Image: {
            const SensorySession& sess = m_sessions[entry.sessionIdx];
            spec.imagePaths = sess.imagePaths;
            // Prefer layout overrides if present in the layout, else fall back
            // to session-stored layouts.
            const ImageSlideLayout imgLayout = effective.imageSlides.value(entry.key);
            spec.imageLayouts = imgLayout.imageLayouts.isEmpty()
                ? sess.imageLayouts : imgLayout.imageLayouts;
            spec.imageCrops = imgLayout.imageCrops.isEmpty()
                ? sess.imageCrops : imgLayout.imageCrops;
            break;
        }
        case SlideKind::Cumulative: {
            // entry.sessionIdx points at the first session of this test group
            // (see buildSlideIndex). Use its testTitle to scope aggregation.
            QString groupTitle = QStringLiteral("Sensory Evaluation");
            if (entry.sessionIdx >= 0 && entry.sessionIdx < m_sessions.size()
                && !m_sessions[entry.sessionIdx].testTitle.isEmpty()) {
                groupTitle = m_sessions[entry.sessionIdx].testTitle;
            } else if (!m_sessions.isEmpty()
                       && !m_sessions.first().testTitle.isEmpty()) {
                groupTitle = m_sessions.first().testTitle;
            }
            spec.title = QStringLiteral("Sensory - %1 - Cumulative").arg(groupTitle);
            SensorySession cumSess;
            buildCumulativeData(m_sessions, excludedSamples,
                                spec.tableHeaders, spec.tableRows, cumSess,
                                groupTitle);
            spec.radarPixmap = renderCumulativeRadarImage(cumSess);
            spec.layout = effective.cumulative;
            break;
        }
    }

    // Apply table sort if configured and the column exists.
    if (!effective.tableSort.column.isEmpty() && !spec.tableHeaders.isEmpty()) {
        const int sortCol = spec.tableHeaders.indexOf(effective.tableSort.column);
        if (sortCol >= 0) {
            const Qt::SortOrder order = effective.tableSort.order;
            std::sort(spec.tableRows.begin(), spec.tableRows.end(),
                      [sortCol, order](const QStringList& a, const QStringList& b) {
                const QString& av = a.value(sortCol);
                const QString& bv = b.value(sortCol);
                bool aOk = false, bOk = false;
                const double an = av.toDouble(&aOk);
                const double bn = bv.toDouble(&bOk);
                if (aOk && bOk)
                    return order == Qt::AscendingOrder ? an < bn : an > bn;
                return order == Qt::AscendingOrder ? av < bv : av > bv;
            });
        }
    }
    return spec;
}

// Delegates to the shared static helper using this source's sessions.
bool SensoryReportSource::writePptx(const QString& outPath, const ReportLayout& layout,
                                     const QSet<QString>& excludedSamples, QString* errorOut)
{
    return writeSensoryPptx(m_sessions, layout, excludedSamples, outPath, errorOut);
}

// ── Shared PPTX renderer (used by legacy + IReportSource paths) ──────────
//
// This is the body of the original SensoryPanel::generateCombinedPptx, lifted
// here so it can be invoked through both entry points. With
// layout = computeDefaultLayout(sessions) and excludedSamples = {}, the
// output is bit-for-bit identical to the legacy generateCombinedPptx output.
//
// Layout consultation is applied via the PptxWriter 5-arg addContentSlide
// overload (Task 7): the legacy positions are first computed into the
// SlideTable / SlideImage values, then the per-slide ContentSlideLayout is
// looked up in `layout.contentSlides` (or `layout.cumulative` for the
// cumulative slide). When a layout rect is null, addContentSlide falls
// through to the legacy positions baked into the SlideTable / SlideImage —
// preserving legacy behavior when called with computeDefaultLayout values.
bool SensoryReportSource::writeSensoryPptx(const QVector<SensorySession>& sessions,
                                            const ReportLayout& layout,
                                            const QSet<QString>& excludedSamples,
                                            const QString& outPath,
                                            QString* errorOut)
{
    auto setErr = [errorOut](const QString& msg) {
        if (errorOut) *errorOut = msg;
    };

    if (sessions.isEmpty()) {
        setErr(QStringLiteral("No sessions to include in the report."));
        return false;
    }

    PptxWriter pptx;
    pptx.setResourcePath(findResourcePath());

    QString coverTitle = sessions.first().testTitle;
    if (coverTitle.isEmpty()) coverTitle = QStringLiteral("Sensory Evaluation");
    QString coverDate = sessions.first().date.isEmpty()
        ? QDate::currentDate().toString("MMMM d, yyyy")
        : sessions.first().date;
    // Pass persisted font sizes (0 = use PptxWriter's legacy 46/24 pt defaults).
    pptx.addCoverSlide(coverTitle, coverDate,
                        layout.coverTitleFontPt, layout.coverSubtitleFontPt);

    QMap<QString, QVector<int>> groups;
    QStringList groupOrder;
    for (int i = 0; i < sessions.size(); ++i) {
        QString title = sessions[i].testTitle.isEmpty()
            ? QStringLiteral("Sensory Evaluation") : sessions[i].testTitle;
        if (!groups.contains(title)) groupOrder.append(title);
        groups[title].append(i);
    }

    bool multiGroup = groups.size() > 1;

    for (const QString& groupTitle : groupOrder) {
        if (multiGroup) {
            QString groupDate;
            int firstIdx = groups[groupTitle].first();
            groupDate = sessions[firstIdx].date.isEmpty()
                ? QDate::currentDate().toString("MMMM d, yyyy")
                : sessions[firstIdx].date;
            // Per-group cover acts as a section divider; pick up the
            // divider-title font customization (keyed by the first session
            // idx, matching buildSlideIndex's divider key convention).
            const QString dividerKey =
                QStringLiteral("divider_%1").arg(firstIdx);
            const int dividerFontPt =
                layout.dividerTitleFontPts.value(dividerKey, 0);
            pptx.addCoverSlide(groupTitle, groupDate,
                                dividerFontPt, /*dateFontPt=*/0);
        }

        for (int si : groups[groupTitle]) {
            const SensorySession& sess = sessions[si];

            // Resolve per-slide layout once; the propertiesBox extraXml below
            // and the addContentSlide call further down both consult it.
            const QString contentKey = QStringLiteral("content_%1").arg(si);
            const ContentSlideLayout slideLayout =
                layout.contentSlides.value(contentKey, ContentSlideLayout{});

            SlideTable rawTable;
            rawTable.headers << "Sample";
            for (const QString& m : kSensoryMetrics)
                rawTable.headers << m;
            rawTable.headers << "Puff Length (s)";
            rawTable.headers << "Comments";

            // Sample exclusion: skip samples whose key is in excludedSamples.
            // Key format mirrors SensoryReportSource::allSamples().
            for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
                const QString sampleKey = QStringLiteral("%1#%2").arg(si).arg(sIdx);
                if (excludedSamples.contains(sampleKey)) continue;
                const SensorySample& s = sess.samples[sIdx];
                QStringList row;
                row << (s.name.isEmpty() ? QString("Sample") : s.name);
                for (const QString& m : kSensoryMetrics) {
                    double v = s.scores.value(m, 5.0);
                    row << ((v == qRound(v)) ? QString::number(qRound(v)) : QString::number(v, 'f', 1));
                }
                row << QString::number(s.puffLengthSec, 'f', 1);
                row << s.comments;
                rawTable.rows.append(row);
            }

            rawTable.x = 0.32;
            rawTable.w = 12.7;
            rawTable.h = 0.95;
            // 8 columns: Sample, 5 metrics, Puff Length (s), Comments. Steal
            // from Comments (0.34 -> 0.27) and trim a metric (0.13 -> 0.12) to
            // make room for a 0.08-wide Puff Length column. Sum = 1.00.
            rawTable.colWidthFractions = {0.10, 0.11, 0.10, 0.11, 0.11, 0.12, 0.08, 0.27};

            // Build title first so we can estimate its line count
            QString title = "Sensory Evaluation";
            if (!sess.testTitle.isEmpty()) title = sess.testTitle;
            if (!sess.testerName.isEmpty()) title += " - " + sess.testerName;

            // Estimate title height: 32pt Montserrat bold in 11" box ≈ 38 chars/line
            int titleLines = qMax(1, (title.length() + 37) / 38);
            double titleBottom = 0.1 + titleLines * 0.45 + 0.10;
            rawTable.y = qMax(0.75, titleBottom);

            // Estimate table height accounting for text wrapping in cells.
            // Note: only included samples contribute to row-height accumulation.
            const int sampleCharsPerLine = 16;
            const int commentsCharsPerLine = 54;
            double tableH = 0.30;  // header row
            for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
                const QString sampleKey = QStringLiteral("%1#%2").arg(si).arg(sIdx);
                if (excludedSamples.contains(sampleKey)) continue;
                const SensorySample& s = sess.samples[sIdx];
                QString name = s.name.isEmpty() ? QString("Sample") : s.name;
                int nameLines = qMax(1, (name.length() + sampleCharsPerLine - 1) / sampleCharsPerLine);
                int commentLines = s.comments.isEmpty() ? 1 : qMax(1, (s.comments.length() + commentsCharsPerLine - 1) / commentsCharsPerLine);
                int extraLines = qMax(nameLines, commentLines) - 1;
                tableH += 0.22 + extraLines * 0.16;  // base row + extra per wrapped line
            }

            // Build a filtered session for the radar chart: skip excluded samples.
            SensorySession filteredSess = sess;
            filteredSess.samples.clear();
            for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
                const QString sampleKey = QStringLiteral("%1#%2").arg(si).arg(sIdx);
                if (excludedSamples.contains(sampleKey)) continue;
                filteredSess.samples.append(sess.samples[sIdx]);
            }

            RadarChartWidget tempChart;
            tempChart.setSessions({filteredSess});
            tempChart.setReportMode(true);
            tempChart.setReportCropTop(70);
            tempChart.setReportCropBottom(130);
            tempChart.resize(1163, 858);

            QPixmap pixmap(1163, 858);
            pixmap.fill(Qt::white);
            QPainter painter(&pixmap);
            tempChart.render(&painter);
            painter.end();

            // Crop whitespace — less top (preserve "Overall Liking"), more bottom
            int cropTop = 70;
            int cropBottom = 130;
            QPixmap cropped = pixmap.copy(0, cropTop, 1163, 858 - cropTop - cropBottom);

            QByteArray chartPng;
            {
                QBuffer buf(&chartPng);
                buf.open(QIODevice::WriteOnly);
                cropped.save(&buf, "PNG");
            }

            // Dynamic chart placement: fill space between table bottom and slide bottom
            const double slideH = 7.5;                          // standard 16:9 slide height
            const double slideW = 13.33;
            tableH += 0.10;                                      // padding buffer for OOXML row expansion
            double tableBottom = rawTable.y + tableH;
            double gapAbove = 0.10;                              // gap above chart
            double gapBelow = 0.10;                              // gap below chart
            double chartY = tableBottom + gapAbove;
            double chartH = slideH - gapBelow - chartY;
            double imgAspect = double(cropped.width()) / double(cropped.height());
            double chartW = chartH * imgAspect;
            if (chartW > slideW - 0.4) {                        // clamp if too wide
                chartW = slideW - 0.4;
                chartH = chartW / imgAspect;
                chartY = slideH - gapBelow - chartH;            // re-anchor to bottom
            }
            double chartX = (slideW - chartW) / 2.0;            // centered

            QVector<SlideImage> plots;
            SlideImage chartImg;
            chartImg.pngData = std::move(chartPng);
            chartImg.x = chartX;
            chartImg.y = chartY;
            chartImg.w = chartW;
            chartImg.h = chartH;
            plots.append(std::move(chartImg));

            // ── Properties textbox (right of chart) ────────────────────────
            // Gather property lines: Media, Control, Blind?, Primary Differences,
            // Highest Rated, Lowest Rated
            QStringList propLines;
            if (!sess.media.isEmpty())
                propLines << "Media: " + sess.media;
            if (!sess.control.isEmpty())
                propLines << "Control: " + sess.control;
            propLines << QString("Blind: %1").arg(sess.isBlind ? "Yes" : "No");
            if (!sess.primaryDifferences.isEmpty())
                propLines << "Primary Difference(s): " + sess.primaryDifferences;

            // #7: power type summary across included samples
            {
                QSet<QString> ptypes;
                for (const SensorySample& samp : filteredSess.samples) {
                    const QString pt = samp.powerType.trimmed();
                    if (!pt.isEmpty()) ptypes.insert(pt);
                }
                if (ptypes.size() == 1) {
                    propLines << "Power Type: " + *ptypes.constBegin();
                } else if (ptypes.size() > 1) {
                    QStringList sorted(ptypes.begin(), ptypes.end());
                    std::sort(sorted.begin(), sorted.end());
                    propLines << "Power Type: " + sorted.join(", ");
                }
            }

            // Highest / Lowest Rated by Overall Liking — filtered by exclusion
            if (!filteredSess.samples.isEmpty()) {
                double maxScore = -1.0, minScore = 10.0;
                QStringList maxNames, minNames;
                for (const SensorySample& samp : filteredSess.samples) {
                    double score = samp.scores.value("Overall Liking", -1.0);
                    if (score < 0) continue;
                    if (score > maxScore) { maxScore = score; maxNames.clear(); maxNames << samp.name; }
                    else if (qAbs(score - maxScore) < 0.01) maxNames << samp.name;
                    if (score < minScore) { minScore = score; minNames.clear(); minNames << samp.name; }
                    else if (qAbs(score - minScore) < 0.01) minNames << samp.name;
                }
                if (!maxNames.isEmpty())
                    propLines << "Highest Rated: " + maxNames.join(", ");
                if (!minNames.isEmpty())
                    propLines << "Lowest Rated: " + minNames.join(", ");
            }

            QString extraXml;
            if (!propLines.isEmpty()) {
                // Build multi-paragraph textbox XML with tight white fill.
                // Font size comes from layout when set (sentinel 0 = legacy 16 pt).
                const int propsFontPt =
                    slideLayout.propertiesBox.fontPt > 0
                        ? slideLayout.propertiesBox.fontPt : 16;
                const int propsSz100 = propsFontPt * 100;

                double tbW    = 3.17;
                // Char-per-line and per-line-height estimates scale linearly
                // with the font size relative to the legacy 16 pt baseline.
                const double scaleVsLegacy = propsFontPt / 16.0;
                const int charsPerLine = qMax(6,
                    int(18.0 / scaleVsLegacy));
                int wrappedLines = 0;
                for (const QString& line : propLines) {
                    int extraLines = qMax(0,
                        (line.length() - charsPerLine) / charsPerLine);
                    wrappedLines += extraLines;
                }
                double tbH = qMax(2.0,
                    (2.0 + wrappedLines * 0.20) * scaleVsLegacy);
                // Anchor bottom-right corner to 0.05" from slide edges
                double tbX    = slideW - tbW - 0.05;
                double tbY    = slideH - tbH - 0.05;

                QString paras;
                for (const QString& line : propLines) {
                    QString safe = line;
                    safe.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;");
                    paras += QStringLiteral(
                        R"(<a:p><a:pPr algn="l"/>)"
                        R"(<a:r><a:rPr lang="en-US" sz="%2" b="0" dirty="0">)"
                        R"(<a:solidFill><a:srgbClr val="333333"/></a:solidFill>)"
                        R"(<a:latin typeface="Calibri"/>)"
                        R"(</a:rPr><a:t>%1</a:t></a:r></a:p>)")
                        .arg(safe).arg(propsSz100);
                }

                auto toEmu = [](double in) { return QString::number(qRound64(in * 914400.0)); };
                extraXml = QStringLiteral(
                    R"(<p:sp><p:nvSpPr>)"
                    R"(<p:cNvPr id="90" name="PropsBox"/>)"
                    R"(<p:cNvSpPr txBox="1"/><p:nvPr/>)"
                    R"(</p:nvSpPr><p:spPr>)"
                    R"(<a:xfrm><a:off x="%1" y="%2"/><a:ext cx="%3" cy="%4"/></a:xfrm>)"
                    R"(<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>)"
                    R"(<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>)"
                    R"(<a:ln w="0"><a:noFill/></a:ln>)"
                    R"(</p:spPr><p:txBody>)"
                    R"(<a:bodyPr wrap="square" lIns="36000" tIns="18000" rIns="36000" bIns="18000" rtlCol="0"/>)"
                    R"(<a:lstStyle/>%5)"
                    R"(</p:txBody></p:sp>)")
                    .arg(toEmu(tbX), toEmu(tbY),
                         toEmu(tbW), toEmu(tbH))
                    .arg(paras);
            }

            // slideLayout was resolved at the top of this iteration so the
            // propertiesBox extraXml above and this call share the same view
            // of the layout entry.
            pptx.addContentSlide(title, rawTable, plots, slideLayout, extraXml);

            // ── Image slide for this session ────────────────────────────────
            // Image-slide layouts are NOT parameterized in this task (Phase
            // 1B/2 concern). Image slides keep their legacy positions.
            if (!sess.imagePaths.isEmpty()) {
                const bool hasLayouts = (sess.imageLayouts.size() == sess.imagePaths.size());
                if (hasLayouts) {
                    QVector<SlideImage> slideImages;
                    for (int i = 0; i < sess.imagePaths.size(); ++i) {
                        QRectF crop = (i < sess.imageCrops.size())
                            ? sess.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sess.imagePaths[i], crop);
                        if (data.isEmpty()) continue;
                        SlideImage simg;
                        simg.pngData = data;
                        const QRectF& r = sess.imageLayouts[i];
                        simg.x = r.x(); simg.y = r.y();
                        simg.w = r.width(); simg.h = r.height();
                        slideImages.append(simg);
                    }
                    if (!slideImages.isEmpty())
                        pptx.addImageSlide(title + " - Images", slideImages);
                } else {
                    QVector<QByteArray> imgBytes;
                    for (int i = 0; i < sess.imagePaths.size(); ++i) {
                        QRectF crop = (i < sess.imageCrops.size())
                            ? sess.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sess.imagePaths[i], crop);
                        if (!data.isEmpty()) imgBytes.append(std::move(data));
                    }
                    if (!imgBytes.isEmpty())
                        pptx.addImageSlide(title + " - Images", imgBytes);
                }
            }
        }
    }

    // ── Cumulative slides: one per test (testTitle group) ───────────────────
    // A combined report covering multiple distinct tests gets one cumulative
    // slide per test at the end. Per-test grouping reuses `groups`/`groupOrder`
    // built above for the content/divider section. All cumulative slides share
    // layout.cumulative — it's a single global layout entry by design.
    for (const QString& groupTitle : groupOrder) {
        const QVector<int>& groupIdx = groups[groupTitle];

        // Build a single session whose samples are per-device averages
        // across this test group only.
        struct DeviceAccum { QMap<QString, double> sums; int count = 0; };
        QMap<QString, DeviceAccum> deviceMap;
        QStringList deviceOrder;

        for (int sessIdx : groupIdx) {
            const SensorySession& sess = sessions[sessIdx];
            for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
                const QString sampleKey = QStringLiteral("%1#%2").arg(sessIdx).arg(sIdx);
                if (excludedSamples.contains(sampleKey)) continue;
                const SensorySample& s = sess.samples[sIdx];
                QString key = s.name.isEmpty() ? QStringLiteral("Sample") : s.name;
                if (!deviceMap.contains(key)) deviceOrder.append(key);
                DeviceAccum& acc = deviceMap[key];
                for (const QString& m : kSensoryMetrics)
                    acc.sums[m] += s.scores.value(m, 5.0);
                acc.count++;
            }
        }

        SensorySession cumSess;
        cumSess.testTitle = groupTitle;

        SlideTable cumTable;
        cumTable.headers << "Device";
        for (const QString& m : kSensoryMetrics)
            cumTable.headers << m;

        for (const QString& devName : deviceOrder) {
            const DeviceAccum& acc = deviceMap[devName];
            SensorySample avgSample;
            avgSample.name = devName;
            QStringList row;
            row << devName;
            for (const QString& m : kSensoryMetrics) {
                double avg = acc.sums.value(m, 0) / qMax(1, acc.count);
                avgSample.scores[m] = avg;
                row << QString::number(avg, 'f', 1);
            }
            cumSess.samples.append(avgSample);
            cumTable.rows.append(row);
        }

        cumTable.x = 0.32;
        cumTable.w = 12.7;
        cumTable.h = 0.95;
        cumTable.colWidthFractions = {0.22, 0.156, 0.156, 0.156, 0.156, 0.156};

        // Estimate cumulative title height (same logic as per-session)
        QString cumTitle = "Sensory - " + groupTitle + " - Cumulative";
        int cumTitleLines = qMax(1, (cumTitle.length() + 37) / 38);
        double cumTitleBottom = 0.1 + cumTitleLines * 0.45 + 0.10;
        cumTable.y = qMax(0.75, cumTitleBottom);

        // Estimate table height accounting for text wrapping
        // Device column: 22% of 12.7" = 2.794" → ~35 chars at 9pt
        const int deviceCharsPerLine = 35;
        double cumTableH = 0.30;  // header
        for (const QString& devName : deviceOrder) {
            int extraLines = qMax(1, (devName.length() + deviceCharsPerLine - 1) / deviceCharsPerLine) - 1;
            cumTableH += 0.22 + extraLines * 0.16;
        }
        cumTableH += 0.10;  // padding buffer for OOXML row expansion

        // Render cumulative radar chart (report mode, same as per-session)
        RadarChartWidget cumChart;
        cumChart.setSessions({cumSess});
        cumChart.setReportMode(true);
        cumChart.setReportCropTop(70);
        cumChart.setReportCropBottom(130);
        cumChart.resize(1163, 858);

        QPixmap cumPix(1163, 858);
        cumPix.fill(Qt::white);
        QPainter cumPainter(&cumPix);
        cumChart.render(&cumPainter);
        cumPainter.end();

        int cumCropTop = 70;
        int cumCropBottom = 130;
        QPixmap cumCropped = cumPix.copy(0, cumCropTop, 1163, 858 - cumCropTop - cumCropBottom);

        QByteArray cumPng;
        {
            QBuffer buf(&cumPng);
            buf.open(QIODevice::WriteOnly);
            cumCropped.save(&buf, "PNG");
        }

        // Dynamic chart placement (same logic as per-session slides)
        const double slideH = 7.5;
        const double slideW = 13.33;
        double cumTableBottom = cumTable.y + cumTableH;
        double cumGapAbove = 0.10;
        double cumGapBelow = 0.10;
        double cumChartY = cumTableBottom + cumGapAbove;
        double cumChartH = slideH - cumGapBelow - cumChartY;
        double cumImgAspect = double(cumCropped.width()) / double(cumCropped.height());
        double cumChartW = cumChartH * cumImgAspect;
        if (cumChartW > slideW - 0.4) {
            cumChartW = slideW - 0.4;
            cumChartH = cumChartW / cumImgAspect;
            cumChartY = slideH - cumGapBelow - cumChartH;
        }
        double cumChartX = (slideW - cumChartW) / 2.0;

        QVector<SlideImage> cumPlots;
        SlideImage cumImg;
        cumImg.pngData = std::move(cumPng);
        cumImg.x = cumChartX;
        cumImg.y = cumChartY;
        cumImg.w = cumChartW;
        cumImg.h = cumChartH;
        cumPlots.append(std::move(cumImg));

        // Cumulative slide: consult layout.cumulative for overrides; null rects
        // fall through to the legacy positions baked into cumTable/cumPlots.
        pptx.addContentSlide(cumTitle, cumTable, cumPlots, layout.cumulative);
    }

    if (!pptx.save(outPath)) {
        setErr(pptx.lastError());
        return false;
    }
    return true;
}

// ── Persistence ──────────────────────────────────────────────────────────
ReportLayout SensoryReportSource::loadLayout() const
{
    // Defaults are the safety net — if no DB or the JSON is missing/corrupt,
    // fall back to the math in computeDefaultLayout.
    if (!m_db || m_sessions.isEmpty())
        return computeDefaultLayout(m_sessions);

    // Per-session slide layouts are anchored to the FIRST session's id.
    // (The dialog edits a single coherent layout per source — the "session"
    //  it persists into is the first one. Rationale: in v1 a SensoryReportSource
    //  represents one logical report run; there's no need to fan layouts out
    //  to each individual session row.)
    const int anchorId = m_sessions.first().id;
    if (anchorId <= 0)
        return computeDefaultLayout(m_sessions);

    const QString json = m_db->loadSensoryLayout(anchorId);
    if (json.isEmpty())
        return computeDefaultLayout(m_sessions);

    bool ok = false;
    const QJsonObject obj = QJsonDocument::fromJson(json.toUtf8()).object();
    ReportLayout layout = ReportLayout::fromJson(obj, &ok);
    if (!ok)
        return computeDefaultLayout(m_sessions);

    // Cumulative is global — pull from settings table to override whatever
    // the per-session JSON had stored (if anything).
    const QString cumJson = m_db->loadCumulativeLayout();
    if (!cumJson.isEmpty()) {
        const QJsonObject co = QJsonDocument::fromJson(cumJson.toUtf8()).object();
        ContentSlideLayout cum;
        cum.title = rectFromJsonArray(co.value("title").toArray());
        cum.table = rectFromJsonArray(co.value("table").toArray());
        cum.radar = rectFromJsonArray(co.value("radar").toArray());
        const QJsonObject pb = co.value("propertiesBox").toObject();
        cum.propertiesBox.rect = rectFromJsonArray(pb.value("rect").toArray());
        cum.propertiesBox.text = pb.value("text").toString();
        layout.cumulative = cum;
    }

    return layout;
}

void SensoryReportSource::saveLayout(const ReportLayout& layout)
{
    if (!m_db || m_sessions.isEmpty()) return;
    const int anchorId = m_sessions.first().id;
    if (anchorId <= 0) return;          // unsaved session — caller must save first

    const QJsonDocument doc(layout.toJson());
    m_db->saveSensoryLayout(anchorId,
                             QString::fromUtf8(doc.toJson(QJsonDocument::Compact)));

    // Cumulative is global — write only the cumulative sub-object.
    // Phase 2: Excel custom-property round-trip is intentionally not done here.
    QJsonObject cum;
    cum["title"]         = rectToJsonArray(layout.cumulative.title);
    cum["table"]         = rectToJsonArray(layout.cumulative.table);
    cum["radar"]         = rectToJsonArray(layout.cumulative.radar);
    cum["propertiesBox"] = QJsonObject{
        { "rect", rectToJsonArray(layout.cumulative.propertiesBox.rect) },
        { "text", layout.cumulative.propertiesBox.text }
    };
    m_db->saveCumulativeLayout(QString::fromUtf8(
        QJsonDocument(cum).toJson(QJsonDocument::Compact)));
}
ReportLayout SensoryReportSource::computeDefaultLayout(const QVector<SensorySession>& sessions)
{
    constexpr double slideW = 13.33;
    constexpr double slideH = 7.5;
    constexpr double tableX = 0.32;
    constexpr double tableW = 12.7;
    constexpr double tableY = 0.75;
    constexpr double gap    = 0.10;
    constexpr double propsW = 3.17;
    constexpr double propsH = 2.70;     // tall enough for typical 6-line content
                                          // (Media/Control/Blind/Primary/Highest/Lowest)
                                          // without clipping; TextItem auto-grow handles
                                          // sessions with longer wrapped lines.

    auto contentForSession = [&](const SensorySession& s) {
        ContentSlideLayout c;

        // Title spans full slide width above the table
        c.title = QRectF(tableX, 0.10, tableW, 0.55);

        // Table height grows with sample count: header (~0.5") + each row ~0.33"
        const double tableH = 0.50 + s.samples.size() * (1.0 / 3.0);
        c.table = QRectF(tableX, tableY, tableW, tableH);

        // Radar centered horizontally below the table. Aspect matches the
        // cropped chart pixmap (1163 x 658 after RadarChartWidget renders at
        // 1163x858 and we crop top=70 / bottom=130). Width is height *
        // ~1.767 so the preview matches the production PPTX layout.
        constexpr double radarAspect = 1163.0 / 658.0;        // ~ 1.7675
        const double tableBottom = tableY + tableH + gap;
        const double availH      = slideH - gap - tableBottom;
        // Min 0.5" so the radar doesn't shrink to nothing on huge tables.
        double radarH = qMax(0.5, availH);
        double radarW = radarH * radarAspect;
        double radarY = tableBottom;
        // Width clamp: keep 0.2" margin on each side of the slide.
        if (radarW > slideW - 0.4) {
            radarW = slideW - 0.4;
            radarH = radarW / radarAspect;
            radarY = slideH - gap - radarH;     // re-anchor to bottom
        }
        const double radarX = (slideW - radarW) / 2.0;
        c.radar = QRectF(radarX, radarY, radarW, radarH);

        // Properties textbox anchored bottom-right
        c.propertiesBox.rect = QRectF(slideW - propsW - 0.05,
                                       slideH - propsH - 0.05,
                                       propsW, propsH);
        c.propertiesBox.text.clear();   // filled by buildSlide() at render time
        return c;
    };

    ReportLayout layout;

    // Cover slide: centered title, subtitle below
    layout.coverTitle    = QRectF(0.5, 2.5, slideW - 1.0, 1.5);
    layout.coverSubtitle = QRectF(0.5, 4.2, slideW - 1.0, 0.8);

    // Group sessions by tester, preserving first-seen order — must mirror
    // SensoryReportSource::buildSlideIndex so the layout dict's divider keys
    // line up with the slide-index list (no orphan entries).
    QHash<QString, QVector<int>> byTester;
    QStringList testerOrder;
    for (int i = 0; i < sessions.size(); ++i) {
        const QString& t = sessions[i].testerName;
        if (!byTester.contains(t)) testerOrder.append(t);
        byTester[t].append(i);
    }

    for (const QString& tester : testerOrder) {
        const QVector<int>& idxs = byTester[tester];
        // One divider per tester group, keyed on the group's first session id —
        // exact same key buildSlideIndex emits for the matching Divider slide.
        const QString dividerKey = QStringLiteral("divider_%1").arg(idxs.first());
        layout.dividerTitles[dividerKey] = QRectF(0.5, slideH/2 - 0.75,
                                                   slideW - 1.0, 1.5);

        for (int i : idxs) {
            const QString contentKey = QStringLiteral("content_%1").arg(i);
            layout.contentSlides[contentKey] = contentForSession(sessions[i]);

            if (!sessions[i].imagePaths.isEmpty()) {
                ImageSlideLayout img;
                const int n = sessions[i].imagePaths.size();
                const int cols = qMin(4, n);
                const int rows = (n + cols - 1) / cols;
                const double cellW = (slideW - 0.5) / cols;
                const double cellH = (slideH - 0.5) / qMax(1, rows);
                for (int k = 0; k < n; ++k) {
                    const int r = k / cols;
                    const int col = k % cols;
                    img.imageLayouts.append(QRectF(0.25 + col * cellW,
                                                    0.25 + r * cellH,
                                                    cellW - 0.1, cellH - 0.1));
                    img.imageCrops.append(QRectF(0, 0, 1, 1));
                }
                layout.imageSlides[QStringLiteral("image_%1").arg(i)] = img;
            }
        }
    }

    // Cumulative slide layout — same shape as content, only when 2+ sessions
    if (sessions.size() >= 2)
        layout.cumulative = contentForSession(sessions.first());

    return layout;
}

} // namespace DVE

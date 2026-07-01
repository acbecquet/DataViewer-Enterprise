#include "UiStress.h"

#include "../MainWindow.h"
#include "../widgets/ScrollHost.h"
#include "AppTheme.h"

#include <QApplication>
#include <QCoreApplication>
#include <QDir>
#include <QDockWidget>
#include <QEventLoop>
#include <QFile>
#include <QFont>
#include <QJsonArray>
#include <QRegularExpression>
#include <QJsonDocument>
#include <QJsonObject>
#include <QPixmap>
#include <QSize>
#include <QStringList>
#include <QStyle>
#include <QTextStream>
#include <QTimer>
#include <QWidget>

namespace DVE {

namespace {

// Spin the event loop for ms milliseconds without blocking signal/slot
// delivery (so the responsive debounce fires and child layouts re-flow).
void settle(int ms)
{
    QEventLoop loop;
    QTimer::singleShot(ms, &loop, &QEventLoop::quit);
    loop.exec();
}

// Multiply every `font-size: Npt` / `Npx` in a stylesheet by `scale`.
//
// The app pins text size via hard-coded `font-size: 9pt` rules in the AppTheme
// stylesheet (AppTheme.cpp), and a QSS font-size ALWAYS overrides
// QApplication::setFont(). So scaling only the application font is a visual
// no-op -- the rendered text (and therefore every font-derived sizeHint) never
// changes, which is why the earlier scale axis produced identical `need` at
// x1.0 and x2.0. To make the axis real WITHOUT editing AppTheme, the harness
// rewrites the live stylesheet's font sizes for each scale. This is the true
// in-process analogue of OS text-scaling for this app.
QString scaleStyleSheetFonts(const QString& css, qreal scale)
{
    QString out = css;
    // Matches "font-size:" + optional spaces + a number + pt|px unit.
    static const QRegularExpression re(
        QStringLiteral("font-size\\s*:\\s*([0-9]+(?:\\.[0-9]+)?)\\s*(pt|px)"),
        QRegularExpression::CaseInsensitiveOption);
    QString rebuilt;
    int last = 0;
    auto it = re.globalMatch(css);
    while (it.hasNext()) {
        const QRegularExpressionMatch m = it.next();
        rebuilt += css.mid(last, m.capturedStart() - last);
        const qreal px = m.captured(1).toDouble() * scale;
        rebuilt += QStringLiteral("font-size: %1%2")
                       .arg(QString::number(px, 'f', 2), m.captured(2));
        last = m.capturedEnd();
    }
    rebuilt += css.mid(last);
    return rebuilt.isEmpty() ? out : rebuilt;
}

// The closed-loop no-clip check, run inside the real DataViewer.exe (where a
// fully-constructed MainWindow with live DB/python context exists -- which is
// why this guarantee is verified HERE rather than in a headless Qt Test).
//
// Everything is measured on ONE widget: the region's ScrollHost. `need`, `vp`,
// and the scrollbar state all come from that single host, so there is no way to
// mispair a "need" taken from one widget against a viewport taken from another
// (an earlier version measured MainWindow::regionWidget() -- e.g. the whole
// QStackedWidget or the tab-bar-inclusive ribbon, or even a compact-collapsed
// panel that wasn't on screen -- against a different host's viewport, which
// produced false clips even though the UI fit or scrolled fine). The clip
// threshold is the content's MINIMUM size hint: content cannot shrink below it,
// so that -- not the larger preferred sizeHint -- is where a real clip begins.
//
// A region PASSES an axis when the content's minimum fits the viewport OR the
// host's scrollbar is active on that axis. Only axes the host is allowed to
// scroll are checked: a ScrollHost::wrap(..., Qt::Horizontal) row has vertical
// policy AlwaysOff -- it's a fixed-height band by design and can never grow a V
// scrollbar, so V-checking it would flag an overflow that isn't a clip. We skip
// any axis whose policy is Qt::ScrollBarAlwaysOff (symmetric for H and V).
//
// Collapsed docks / not-yet-built lazy panels are skipped (host null/hidden).
bool regionsFitOrScroll(const MainWindow& w, QStringList& failures)
{
    static const QStringList kRegions = {
        QStringLiteral("central"),
        QStringLiteral("navigator"),
        QStringLiteral("notes"),
        QStringLiteral("ribbonGroups"),
    };
    bool ok = true;
    for (const QString& key : kRegions) {
        ScrollHost* host = w.scrollHostFor(key);
        if (!host || !host->isVisible() || !host->widget())
            continue;   // collapsed dock / lazy panel: nothing on screen to clip
        const QWidget* content = host->widget();
        const QSize vp = host->viewport()->size();

        // Clip threshold = the content's minimum. Fall back per-axis to the
        // preferred sizeHint only where the minimum is degenerate (0), so a
        // widget that reports no minimum in one direction is still measured
        // against something meaningful rather than trivially "fitting".
        const QSize minHint = content->minimumSizeHint();
        const QSize prefHint = content->sizeHint();
        const int needW = minHint.width()  > 0 ? minHint.width()  : prefHint.width();
        const int needH = minHint.height() > 0 ? minHint.height() : prefHint.height();

        // Only assert the axes this host can actually scroll. AlwaysOff means
        // the region is a fixed extent in that direction by design.
        const bool checkH = host->horizontalScrollBarPolicy() != Qt::ScrollBarAlwaysOff;
        const bool checkV = host->verticalScrollBarPolicy()   != Qt::ScrollBarAlwaysOff;

        if (checkH && needW > vp.width() && !host->scrollbarActive(Qt::Horizontal)) {
            ok = false;
            failures << QStringLiteral("%1 H need=%2 vp=%3")
                            .arg(key).arg(needW).arg(vp.width());
        }
        if (checkV && needH > vp.height() && !host->scrollbarActive(Qt::Vertical)) {
            ok = false;
            failures << QStringLiteral("%1 V need=%2 vp=%3")
                            .arg(key).arg(needH).arg(vp.height());
        }
    }
    return ok;
}

} // anonymous namespace

int runUiStress(const QString& outDirArg)
{
    const QString outDir = outDirArg.isEmpty()
        ? (QDir::tempPath() + "/dve_ui_stress")
        : outDirArg;
    if (!QDir().mkpath(outDir)) {
        QTextStream(stderr) << "ui-stress: cannot create output dir "
                            << outDir << Qt::endl;
        return 1;
    }

    // The (window-size x text-scale) matrix: corner-snap quarter, half-split,
    // the window floor, and two extreme aspect ratios; scales mimic OS text-
    // scaling at 100/125/150/200%.
    const QList<QSize> sizes = {
        {1920, 1080}, {1280, 800}, {960, 540},
        {800, 600},   {480, 360},  {600, 1200}, {1600, 500},
    };
    const QList<qreal> scales = { 1.0, 1.25, 1.5, 2.0 };

    // Apply the theme ONCE, up front. AppTheme::apply() resets the application
    // font to AppTheme::fontDefault() (9pt), so it must NOT run inside the scale
    // loop -- doing so would clobber the scaled font and make the text-scale axis
    // a no-op. Capture the post-apply() base font as the reference point size.
    AppTheme::apply();
    const QFont baseFont = QApplication::font();
    const qreal basePt = (baseFont.pointSizeF() > 0)
        ? baseFont.pointSizeF()
        : qreal(AppTheme::fontDefault().pointSizeF());
    // Capture the applied theme stylesheet: text size is pinned here (hard-coded
    // `font-size: 9pt` rules), so the scale axis rewrites THIS per scale.
    const QString baseStyleSheet = qApp->styleSheet();

    MainWindow window;
    window.show();
    settle(150);   // a grab needs the widget realized

    QJsonArray cases;
    bool allOk = true;

    for (const qreal scale : scales) {
        // Make the scale axis REAL. AppTheme::apply() already ran once before the
        // loop; here we scale text size the only way that actually takes effect
        // in this app -- by rewriting the pinned `font-size` rules in the theme
        // stylesheet (a QSS font-size overrides QApplication::setFont(), so the
        // app font is scaled too but the stylesheet is what governs rendered
        // text). Reapplying the stylesheet makes Qt re-polish every widget and
        // recompute font-derived sizeHints top-down; the size loop below then
        // measures against the reflowed geometry. VERIFIED: `need` at x2.0 is
        // meaningfully larger than at x1.0 (previously they were identical --
        // that was the scale-axis bug).
        QFont f = baseFont;
        f.setPointSizeF(basePt * scale);
        QApplication::setFont(f);
        qApp->setStyleSheet(scaleStyleSheetFonts(baseStyleSheet, scale));
        QApplication::processEvents();
        settle(80);   // let the re-polish + queued relayout run before measuring

        for (const QSize& s : sizes) {
            const QString label = QStringLiteral("%1x%2_x%3")
                .arg(s.width()).arg(s.height())
                .arg(QString::number(scale, 'g', 3));

            window.resize(s);
            settle(120);                       // debounce + re-flow
            QCoreApplication::processEvents();

            const QString pngName = label + QStringLiteral(".png");
            const QString pngPath = outDir + "/" + pngName;

            const QPixmap shot = window.grab();
            const bool grabbed = !shot.isNull() && shot.save(pngPath, "PNG");

            // Closed-loop no-clip pass/fail for this case (the full-window
            // fits-or-scrolls guarantee, verified in-process). A case is `pass`
            // only if every visible region fits or scrolls AND the grab wrote.
            QStringList failures;
            const bool regionsOk = regionsFitOrScroll(window, failures);
            const bool pass = grabbed && regionsOk;
            if (!pass) allOk = false;

            QJsonArray failArr;
            for (const QString& f : failures)
                failArr.append(f);

            // VeryNarrow dock-collapse evidence (Task 11): below 760px both side
            // docks should be hidden. Recorded per-case by real object name so a
            // human (or a JSON assertion) can confirm the auto-collapse fired
            // without instantiating MainWindow inside a headless Qt Test.
            const QDockWidget* navDock   = window.findChild<QDockWidget*>(
                QStringLiteral("navigatorDock"));
            const QDockWidget* notesDock = window.findChild<QDockWidget*>(
                QStringLiteral("notesDock"));

            QJsonObject o;
            o["label"]      = label;
            o["width"]      = s.width();
            o["height"]     = s.height();
            o["scale"]      = scale;
            o["png"]        = pngName;
            o["grab_w"]     = shot.width();
            o["grab_h"]     = shot.height();
            o["grabbed"]    = grabbed;
            o["regions_ok"] = regionsOk;
            o["pass"]       = pass;          // closed-loop per-case verdict
            o["failures"]   = failArr;       // clipped regions, if any
            o["nav_visible"]   = navDock   ? navDock->isVisible()   : false;
            o["notes_visible"] = notesDock ? notesDock->isVisible() : false;
            cases.append(o);

            QTextStream(stderr) << "ui-stress: " << label
                << (pass ? " PASS" : " FAIL")
                << (failures.isEmpty() ? QString()
                                       : QStringLiteral(" [%1]").arg(failures.join(QStringLiteral("; "))))
                << Qt::endl;
        }
    }

    // Restore the un-scaled base font + stylesheet (the scale loop rewrote both).
    QApplication::setFont(baseFont);
    qApp->setStyleSheet(baseStyleSheet);

    QJsonObject root;
    root["version"]    = QApplication::applicationVersion();
    root["out_dir"]    = QDir(outDir).absolutePath();
    root["all_ok"]     = allOk;   // true only if EVERY case grabbed AND no region clipped
    root["case_count"] = cases.size();
    root["cases"]      = cases;

    const QByteArray json = QJsonDocument(root).toJson(QJsonDocument::Indented);
    QFile idx(outDir + "/index.json");
    if (idx.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        idx.write(json);
        idx.close();
    } else {
        allOk = false;
    }

    QTextStream(stdout) << "ui-stress: wrote " << cases.size()
        << " cases + index.json to " << QDir(outDir).absolutePath()
        << Qt::endl;

    return allOk ? 0 : 1;
}

} // namespace DVE

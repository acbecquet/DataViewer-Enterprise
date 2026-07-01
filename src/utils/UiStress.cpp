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

// The closed-loop no-clip check, run inside the real DataViewer.exe (where a
// fully-constructed MainWindow with live DB/python context exists -- which is
// why this guarantee is verified HERE rather than in a headless Qt Test). For
// each wrapped region, the content either fits its ScrollHost viewport or the
// host's scrollbar is active in the overflow direction. Returns false (and
// appends a human-readable reason to `failures`) if any region is clipped
// without a scrollbar. Collapsed docks / not-yet-built lazy panels are skipped.
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
        QWidget* content = w.regionWidget(key);
        if (!host || !content || !host->isVisible())
            continue;   // collapsed dock / lazy panel: nothing on screen to clip
        const QSize vp   = host->viewport()->size();
        const QSize need = content->sizeHint().expandedTo(content->minimumSizeHint());
        const bool hOk = (need.width()  <= vp.width())  || host->scrollbarActive(Qt::Horizontal);
        const bool vOk = (need.height() <= vp.height()) || host->scrollbarActive(Qt::Vertical);
        if (!hOk) {
            ok = false;
            failures << QStringLiteral("%1 H need=%2 vp=%3")
                            .arg(key).arg(need.width()).arg(vp.width());
        }
        if (!vOk) {
            ok = false;
            failures << QStringLiteral("%1 V need=%2 vp=%3")
                            .arg(key).arg(need.height()).arg(vp.height());
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

    MainWindow window;
    window.show();
    settle(150);   // a grab needs the widget realized

    QJsonArray cases;
    bool allOk = true;

    for (const qreal scale : scales) {
        // Set the scaled app font AFTER apply() (apply() already ran once, before
        // the loop) so the scaled point size actually takes effect, then re-polish
        // every widget against it. This is the in-process analogue of OS text-
        // scaling and is the whole point of the scale axis.
        QFont f = baseFont;
        f.setPointSizeF(basePt * scale);
        QApplication::setFont(f);
        for (QWidget* top : QApplication::topLevelWidgets())
            QApplication::style()->polish(top);   // refresh derived metrics

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

    // Restore the un-scaled base font (apply() already ran before the loop;
    // re-running it here just re-asserts fontDefault(), which equals baseFont).
    QApplication::setFont(baseFont);

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

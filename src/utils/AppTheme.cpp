#include "AppTheme.h"
#include <QApplication>
#include <QFont>
#include <QIcon>
#include <QHash>
#include <QFile>
#include <QFileInfo>
#include <QDir>
#include <QStandardPaths>
#include <QDebug>
#include <cmath>

namespace {

// Format a QColor as a CSS hex string for embedding in the QSS template.
QString hex(const QColor& c) {
    return QString("#%1%2%3")
        .arg(c.red(),   2, 16, QChar('0'))
        .arg(c.green(), 2, 16, QChar('0'))
        .arg(c.blue(),  2, 16, QChar('0'))
        .toUpper();
}

// Resolve the icon directory at runtime. The CLAUDE.md note about Qt 6.10 rcc
// emitting binary TSD output means we cannot use the Qt resource system; icons
// live on disk next to the executable. This mirrors MainWindow::resourcePath().
QString iconDir() {
    static QString cached;
    static bool    resolved = false;
    if (resolved) return cached;
    resolved = true;

    const QString exeDir = QFileInfo(QCoreApplication::applicationFilePath()).absolutePath();

    // 1. Production install: <app>/resources/icons
    QString candidate = exeDir + "/resources/icons";
    if (QDir(candidate).exists()) { cached = candidate; return cached; }

    // 2. Dev build running from build/release sub-dir: <build>/../resources/icons
    candidate = exeDir + "/../resources/icons";
    if (QDir(candidate).exists()) { cached = QDir(candidate).absolutePath(); return cached; }

    // 3. CWD-relative for test runs
    candidate = "resources/icons";
    if (QDir(candidate).exists()) { cached = QDir(candidate).absolutePath(); return cached; }

    qWarning() << "AppTheme: icon directory not found; tried"
               << exeDir + "/resources/icons"
               << exeDir + "/../resources/icons"
               << "resources/icons";
    return cached; // empty
}

} // anon

QFont AppTheme::fontDefault() { return QFont("Segoe UI", 9); }
QFont AppTheme::fontSmall()   { return QFont("Segoe UI", 8); }

QFont AppTheme::fontLabel() {
    QFont f("Segoe UI", 10);
    return f;
}
QFont AppTheme::fontSection() {
    QFont f("Segoe UI", 11);
    f.setWeight(QFont::DemiBold);
    return f;
}
QFont AppTheme::fontPageTitle() {
    QFont f("Segoe UI", 12);
    f.setWeight(QFont::DemiBold);
    return f;
}
QFont AppTheme::fontTitle() { return fontSection(); }  // legacy alias
QFont AppTheme::fontMono()  { return QFont("Consolas", 9); }

QIcon AppTheme::icon(const QString& name) {
    static QHash<QString, QIcon> cache;
    auto it = cache.find(name);
    if (it != cache.end()) return it.value();

    const QString path = iconDir() + "/" + name + ".svg";
    if (!QFile::exists(path)) {
        qWarning() << "AppTheme::icon — missing:" << path;
        cache.insert(name, QIcon());
        return QIcon();
    }
    QIcon ico(path);
    cache.insert(name, ico);
    return ico;
}

QString AppTheme::stylesheet() {
    // Build a token map for the QSS template
    const QString TXT_PRI   = hex(textPrimary());
    const QString TXT_SEC   = hex(textSecondary());
    const QString TXT_MUTED = hex(textMuted());
    const QString SURF_APP  = hex(surfaceApp());
    const QString SURF_PNL  = hex(surfacePanel());
    const QString SURF_RBN  = hex(surfaceRibbon());
    const QString BORD_SUB  = hex(borderSubtle());
    const QString BORD_DEF  = hex(borderDefault());
    const QString BORD_STR  = hex(borderStrong());
    const QString ACCENT    = hex(accent());
    const QString ACC_HOV   = hex(accentHover());
    const QString ACC_PRS   = hex(accentPress());
    const QString ACC_SUB   = hex(accentSubtle());
    const QString HOVER_BG  = hex(hoverBg());
    const QString SELECT_BG = hex(selectBg());
    const QString TABLE_ALT = hex(tableAlt());
    const QString TBL_HDR   = hex(tableHeader());
    const QString TBL_HOV   = hex(tableHeaderHover());
    const QString SUCCESS   = hex(success());
    const QString WARNING   = hex(warning());
    const QString ERROR_C   = hex(error());

    // Suppress unused-variable warnings for tokens reserved for future QSS blocks
    Q_UNUSED(TXT_MUTED)
    Q_UNUSED(HOVER_BG)
    Q_UNUSED(SUCCESS)
    Q_UNUSED(WARNING)

    QString s;
    s += QString(R"(
/* ═══════════════════════════════════════════════════════════════════════
   DataViewer Enterprise — token-based theme (v2.0.9)
   Generated from AppTheme color/spacing/radius accessors.
   ═══════════════════════════════════════════════════════════════════════ */

QWidget {
    background-color: %1;
    color: %2;
    font-family: "Segoe UI";
    font-size: 9pt;
}
QMainWindow { background-color: %1; }
QMainWindow::separator { background-color: %3; width: 4px; height: 4px; }
)").arg(SURF_APP, TXT_PRI, BORD_DEF);

    s += QString(R"(
/* ── Push Button ── */
QPushButton {
    background-color: #E0E0E0;
    color: %1;
    border: 1px solid %2;
    border-radius: 4px;
    padding: 5px 14px;
    font-size: 9pt;
    min-height: 24px;
}
QPushButton:hover  { background-color: %3; border-color: %4; color: %5; }
QPushButton:pressed{ background-color: #99CAFF; border-color: %6; color: %5; }
QPushButton:disabled { background-color: #EBEBEB; color: #AAAAAA; border-color: #D0D0D0; }
QPushButton:focus  { outline: none; border: 1px solid %4; }

QPushButton[primary="true"], QPushButton#btnPrimary, QPushButton.primary {
    background-color: %4; color: #FFFFFF; border: 1px solid %6; font-weight: 600;
}
QPushButton[primary="true"]:hover { background-color: %7; border-color: %4; }
QPushButton[primary="true"]:pressed { background-color: %6; border-color: #003388; }

QPushButton[destructive="true"] {
    background-color: #FFFFFF; color: %8; border: 1px solid %8;
}
QPushButton[destructive="true"]:hover {
    background-color: #FEE2E2; color: #991B1B; border-color: #991B1B;
}
)").arg(TXT_PRI, BORD_DEF, ACC_SUB, ACCENT, "#003388", ACC_PRS, ACC_HOV, ERROR_C);

    s += QString(R"(
/* ── Tool Button (ribbon, etc.) ── */
QToolButton {
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 3px;
    color: %1;
}
QToolButton:hover  { background-color: %2; border-color: #99CAFF; }
QToolButton:pressed{ background-color: #99CAFF; border-color: %3; }
QToolButton:checked{ background-color: %2; border-color: %3; }
QToolButton:disabled { color: #AAAAAA; }
QToolButton::menu-indicator { image: none; width: 0px; }
)").arg(TXT_PRI, ACC_SUB, ACCENT);

    s += QString(R"(
/* ── Inputs (LineEdit / ComboBox / Spin) ── */
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
    background-color: %1;
    color: %2;
    border: 1px solid %3;
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
    selection-background-color: %4;
    selection-color: #003388;
}
QLineEdit:hover, QComboBox:hover, QSpinBox:hover, QDoubleSpinBox:hover { border-color: %5; }
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus { border-color: %6; outline: none; }
QLineEdit:disabled, QComboBox:disabled { background-color: %7; color: #AAAAAA; }
QLineEdit:read-only { background-color: #F5F5F5; color: %8; }
QComboBox::drop-down {
    subcontrol-origin: padding; subcontrol-position: right center;
    width: 20px; border-left: 1px solid %3;
    border-top-right-radius: 4px; border-bottom-right-radius: 4px;
    background-color: %9;
}
QComboBox::down-arrow {
    width: 0; height: 0; border-left: 4px solid transparent;
    border-right: 4px solid transparent; border-top: 5px solid %8; margin: 0 4px;
}
QComboBox QAbstractItemView {
    background-color: %1; border: 1px solid %3;
    selection-background-color: %4; selection-color: #003388; outline: none;
}
)").arg(SURF_PNL, TXT_PRI, BORD_DEF, SELECT_BG, BORD_STR, ACCENT, SURF_APP, TXT_SEC, SURF_RBN);

    s += QString(R"(
/* ── Text editors ── */
QTextEdit, QPlainTextEdit {
    background-color: %1; color: %2; border: 1px solid %3; border-radius: 4px;
    padding: 4px; selection-background-color: %4; selection-color: #003388;
}
QTextEdit:focus, QPlainTextEdit:focus { border-color: %5; outline: none; }

/* ── Tables ── */
QTableWidget, QTableView {
    background-color: %1; alternate-background-color: %6;
    gridline-color: %7; border: 1px solid %3;
    selection-background-color: %4; selection-color: #003388; font-size: 9pt;
}
QTableWidget::item, QTableView::item { padding: 4px 6px; border: none; }
QTableWidget::item:selected, QTableView::item:selected { background-color: %4; color: #003388; }
QTableWidget::item:hover, QTableView::item:hover { background-color: %8; }

QHeaderView { background-color: %9; border: none; }
QHeaderView::section {
    background-color: %9; color: #FFFFFF; font-weight: 600; font-size: 9pt;
    padding: 6px 8px; border: none;
    border-right: 1px solid %10; border-bottom: 2px solid #1F2A37;
}
QHeaderView::section:hover { background-color: %10; }
QHeaderView::section:vertical {
    background-color: %11; color: %2; border-right: 1px solid %3;
    border-bottom: 1px solid %12; padding: 3px 6px; font-weight: normal;
}
QTableCornerButton::section { background-color: %9; border: none; }
)").arg(SURF_PNL, TXT_PRI, BORD_DEF, SELECT_BG, ACCENT, TABLE_ALT, BORD_SUB, ACC_SUB, TBL_HDR, TBL_HOV, SURF_RBN, BORD_SUB);

    s += QString(R"(
/* ── Tab Widget ── */
QTabWidget { background-color: %1; border: none; }
QTabWidget::pane {
    background-color: %2; border: 1px solid %3; border-top: none;
    border-bottom-left-radius: 4px; border-bottom-right-radius: 4px;
}
QTabBar { background-color: transparent; border-bottom: 1px solid %3; }
QTabBar::tab {
    background-color: %4; color: %5; border: 1px solid %3; border-bottom: none;
    border-top-left-radius: 4px; border-top-right-radius: 4px;
    padding: 6px 18px; margin-right: 2px; min-width: 80px; font-size: 9pt;
}
QTabBar::tab:selected {
    background-color: %2; color: %6; border-color: %3;
    border-bottom: 2px solid %6; font-weight: 600;
}
QTabBar::tab:hover:!selected { background-color: %1; color: %7; }
QTabBar::tab:disabled { color: #AAAAAA; }
)").arg(SURF_APP, SURF_PNL, BORD_DEF, "#D4D8DC", TXT_SEC, ACCENT, TXT_PRI);

    s += QString(R"(
/* ── Scroll bars ── */
QScrollBar:vertical   { background-color: %1; width: 10px; border: none; margin: 0; }
QScrollBar::handle:vertical { background-color: #C0C4C8; min-height: 20px; border-radius: 5px; margin: 2px; }
QScrollBar::handle:vertical:hover  { background-color: #9CA3AF; }
QScrollBar::handle:vertical:pressed{ background-color: %2; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; border: none; background: none; }

QScrollBar:horizontal { background-color: %1; height: 10px; border: none; margin: 0; }
QScrollBar::handle:horizontal { background-color: #C0C4C8; min-width: 20px; border-radius: 5px; margin: 2px; }
QScrollBar::handle:horizontal:hover  { background-color: #9CA3AF; }
QScrollBar::handle:horizontal:pressed{ background-color: %2; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; border: none; background: none; }
QScrollBar::add-page, QScrollBar::sub-page { background: none; }

/* ── Splitter ── */
QSplitter::handle { background-color: %3; }
QSplitter::handle:horizontal { width: 4px; }
QSplitter::handle:vertical { height: 4px; }
QSplitter::handle:hover  { background-color: %2; }
)").arg(SURF_APP, ACCENT, BORD_SUB);

    s += QString(R"(
/* ── GroupBox ── */
QGroupBox {
    background-color: %1; border: 1px solid %2; border-radius: 6px;
    margin-top: 12px; padding: 8px 6px 6px 6px; font-size: 9pt; font-weight: 600;
}
QGroupBox::title {
    subcontrol-origin: margin; subcontrol-position: top left;
    padding: 0 6px; color: %3; background-color: %1; left: 8px;
}

/* ── Checkbox / Radio ── */
QCheckBox { spacing: 6px; color: %4; }
QCheckBox::indicator {
    width: 14px; height: 14px; border: 1px solid %2;
    border-radius: 3px; background-color: %1;
}
QCheckBox::indicator:hover   { border-color: %5; background-color: %6; }
QCheckBox::indicator:checked { background-color: %5; border-color: #003388; }
QCheckBox::indicator:disabled{ background-color: %7; border-color: #D0D0D0; }

QRadioButton { spacing: 6px; color: %4; }
QRadioButton::indicator {
    width: 14px; height: 14px; border: 1px solid %2;
    border-radius: 7px; background-color: %1;
}
QRadioButton::indicator:hover   { border-color: %5; background-color: %6; }
QRadioButton::indicator:checked { background-color: %5; border-color: #003388; }

/* ── Tooltip ── */
QToolTip {
    background-color: #FFFFCC; color: %4; border: 1px solid %2;
    padding: 4px 6px; border-radius: 3px; font-size: 8pt;
}

/* ── Tree / List ── */
QTreeWidget, QListWidget {
    background-color: %1; alternate-background-color: #FAFBFD;
    border: 1px solid %2; selection-background-color: %8; selection-color: #003388;
}
QTreeWidget::item, QListWidget::item { padding: 3px 4px; }
QTreeWidget::item:hover, QListWidget::item:hover     { background-color: %6; }
QTreeWidget::item:selected, QListWidget::item:selected { background-color: %8; color: #003388; }

/* ── Labels ── */
QLabel { background-color: transparent; color: %4; }

/* ── Status bar — REDESIGNED in Task 6 (this is a fallback) ── */
QStatusBar { background-color: %9; color: %4; border-top: 1px solid %2; font-size: 9pt; padding: 0 4px; }
QStatusBar::item { border: none; }
QStatusBar QLabel { background-color: transparent; color: %4; padding: 0 4px; }

/* ── DialogButtonBox ── */
QDialogButtonBox QPushButton { min-width: 80px; }
)").arg(SURF_PNL, BORD_DEF, TBL_HDR, TXT_PRI, ACCENT, ACC_SUB, SURF_APP, SELECT_BG, SURF_RBN);

    return s;
}

void AppTheme::apply() {
    QApplication* app = qobject_cast<QApplication*>(QApplication::instance());
    if (!app) return;
    app->setFont(fontDefault());
    app->setStyleSheet(stylesheet());
}

QColor AppTheme::seriesColor(int idx)
{
    if (idx < 0) idx = 0;

    // Curated 20-color qualitative palette — NO yellow / yellow-adjacent hues
    // (they wash out on a projector), ordered most-distinct-first. Shared with
    // the sensory radar chart, which already vetted these for projector use.
    static const QColor kCurated[] = {
        QColor( 31, 119, 180), QColor(214,  39,  40), QColor( 44, 160,  44),
        QColor(148, 103, 189), QColor(255, 127,  14), QColor( 23, 190, 207),
        QColor(227, 119, 194), QColor(140,  86,  75), QColor(127, 127, 127),
        QColor(  0,   0, 139), QColor(139,   0,   0), QColor(  0, 100,   0),
        QColor( 75,   0, 130), QColor(255,   0, 255), QColor(  0, 191, 255),
        QColor(178,  34,  34), QColor( 70, 130, 180), QColor(255,  20, 147),
        QColor( 47,  79,  79), QColor(106,  90, 205),
    };
    static const int kN = int(sizeof(kCurated) / sizeof(kCurated[0]));

    if (idx < kN) return kCurated[idx];

    // Beyond the curated set: golden-angle hue rotation evenly spreads the rest
    // around the wheel. Skip the yellow band (45–70) and pin saturation/value
    // to a vivid, projector-safe band; nudge them each lap so colors stay
    // distinct even past a full turn.
    const int    k    = idx - kN;
    const double aDeg = std::fmod(210.0 + double(k + 1) * 137.508, 360.0);
    int hue = int(aDeg);
    if (hue >= 45 && hue <= 70) hue = (hue + 30) % 360;
    const int lap = k / 8;
    const int val = 196 - (lap % 3) * 28;    // 196 / 168 / 140
    const int sat = 178 + (lap % 2) * 40;    // 178 / 218
    return QColor::fromHsv(hue, qBound(0, sat, 255), qBound(0, val, 255));
}

QVector<QColor> AppTheme::seriesColors(int n)
{
    QVector<QColor> out;
    if (n <= 0) return out;
    out.reserve(n);
    for (int i = 0; i < n; ++i) out.append(seriesColor(i));
    return out;
}

QColor AppTheme::shade(const QColor& base, int idx, int count)
{
    if (count <= 1) return base;
    if (idx < 0) idx = 0;
    if (idx > count - 1) idx = count - 1;

    int h, s, v, a;
    base.getHsv(&h, &s, &v, &a);
    if (a < 0) a = 255;

    // Spread brightness across a projector-safe band (≈0.55→0.85 of full) so
    // even the lightest shade stays readable against white — never toward 1.0.
    const double t    = double(idx) / double(count - 1);   // 0 → 1
    const int    newV = int(std::lround(140.0 + t * (217.0 - 140.0)));

    if (s < 20)  // achromatic base (e.g. the gray family): stay neutral
        return QColor::fromHsv(0, 0, qBound(0, newV, 255), a);

    const int newS = qMax(s, 130);  // floor saturation so light shades aren't pastel
    return QColor::fromHsv(h, qBound(0, newS, 255), qBound(0, newV, 255), a);
}

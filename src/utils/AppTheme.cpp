#include "AppTheme.h"
#include <QApplication>
#include <QFont>

// ── Fonts ─────────────────────────────────────────────────────────────────────

QFont AppTheme::fontDefault() {
    return QFont("Segoe UI", 9);
}

QFont AppTheme::fontSmall() {
    return QFont("Segoe UI", 8);
}

QFont AppTheme::fontTitle() {
    QFont f("Segoe UI", 11);
    f.setWeight(QFont::DemiBold);
    return f;
}

QFont AppTheme::fontMono() {
    return QFont("Consolas", 9);
}

// ── Stylesheet ────────────────────────────────────────────────────────────────

QString AppTheme::stylesheet() {
    return R"(
/* ═══════════════════════════════════════════════════════════════════════
   DataViewer Enterprise – Professional Light Gray Theme
   Inspired by SolidWorks / Windows engineering applications
   ═══════════════════════════════════════════════════════════════════════ */

/* ── Global ─────────────────────────────────────────────────────────── */
QWidget {
    background-color: #F0F0F0;
    color: #1A1A1A;
    font-family: "Segoe UI";
    font-size: 9pt;
}

QMainWindow {
    background-color: #F0F0F0;
}

QMainWindow::separator {
    background-color: #BCBCBC;
    width: 4px;
    height: 4px;
}

/* ── Menu Bar ────────────────────────────────────────────────────────── */
QMenuBar {
    background-color: #E8E8E8;
    color: #1A1A1A;
    border-bottom: 1px solid #BCBCBC;
    padding: 2px 4px;
    spacing: 0px;
}

QMenuBar::item {
    background-color: transparent;
    padding: 4px 10px;
    border-radius: 2px;
}

QMenuBar::item:selected,
QMenuBar::item:hover {
    background-color: #CCE4FF;
    color: #003388;
}

QMenuBar::item:pressed {
    background-color: #0066CC;
    color: #FFFFFF;
}

/* ── Menu (Drop-down) ────────────────────────────────────────────────── */
QMenu {
    background-color: #FFFFFF;
    border: 1px solid #BCBCBC;
    padding: 3px 0px;
}

QMenu::item {
    padding: 5px 28px 5px 28px;
    background-color: transparent;
    color: #1A1A1A;
}

QMenu::item:selected {
    background-color: #CCE4FF;
    color: #003388;
}

QMenu::item:disabled {
    color: #AAAAAA;
}

QMenu::separator {
    height: 1px;
    background-color: #DCDCDC;
    margin: 3px 8px;
}

QMenu::icon {
    padding-left: 8px;
}

/* ── Push Button ─────────────────────────────────────────────────────── */
QPushButton {
    background-color: #E0E0E0;
    color: #1A1A1A;
    border: 1px solid #BCBCBC;
    border-radius: 3px;
    padding: 5px 14px;
    font-size: 9pt;
    min-height: 24px;
}

QPushButton:hover {
    background-color: #CCE4FF;
    border-color: #0066CC;
    color: #003388;
}

QPushButton:pressed {
    background-color: #99CAFF;
    border-color: #0044AA;
    color: #003388;
}

QPushButton:disabled {
    background-color: #EBEBEB;
    color: #AAAAAA;
    border-color: #D0D0D0;
}

QPushButton:focus {
    outline: none;
    border: 1px solid #0066CC;
}

/* Primary button (set property "role" = "primary" or use class) */
QPushButton[primary="true"],
QPushButton#btnPrimary,
QPushButton.primary {
    background-color: #0066CC;
    color: #FFFFFF;
    border: 1px solid #0055AA;
    font-weight: 600;
}

QPushButton[primary="true"]:hover,
QPushButton#btnPrimary:hover,
QPushButton.primary:hover {
    background-color: #0088FF;
    border-color: #0066CC;
}

QPushButton[primary="true"]:pressed,
QPushButton#btnPrimary:pressed,
QPushButton.primary:pressed {
    background-color: #0044AA;
    border-color: #003388;
}

/* ── Tool Button ─────────────────────────────────────────────────────── */
QToolButton {
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 3px;
    padding: 3px;
    color: #1A1A1A;
}

QToolButton:hover {
    background-color: #CCE4FF;
    border-color: #99CAFF;
}

QToolButton:pressed {
    background-color: #99CAFF;
    border-color: #0066CC;
}

QToolButton:checked {
    background-color: #CCE4FF;
    border-color: #0066CC;
}

QToolButton:disabled {
    color: #AAAAAA;
}

QToolButton::menu-indicator {
    image: none;
    width: 0px;
}

/* ── ComboBox ────────────────────────────────────────────────────────── */
QComboBox {
    background-color: #FFFFFF;
    color: #1A1A1A;
    border: 1px solid #BCBCBC;
    border-radius: 3px;
    padding: 3px 8px;
    min-height: 22px;
    selection-background-color: #CCE4FF;
}

QComboBox:hover {
    border-color: #0066CC;
}

QComboBox:focus {
    border-color: #0066CC;
    outline: none;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: right center;
    width: 20px;
    border-left: 1px solid #BCBCBC;
    border-top-right-radius: 3px;
    border-bottom-right-radius: 3px;
    background-color: #E8E8E8;
}

QComboBox::down-arrow {
    width: 0px;
    height: 0px;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid #555555;
    margin: 0px 4px;
}

QComboBox::down-arrow:hover {
    border-top-color: #0066CC;
}

QComboBox QAbstractItemView {
    background-color: #FFFFFF;
    border: 1px solid #BCBCBC;
    selection-background-color: #CCE4FF;
    selection-color: #003388;
    outline: none;
}

/* ── Line Edit ───────────────────────────────────────────────────────── */
QLineEdit {
    background-color: #FFFFFF;
    color: #1A1A1A;
    border: 1px solid #BCBCBC;
    border-radius: 3px;
    padding: 3px 6px;
    min-height: 22px;
    selection-background-color: #CCE4FF;
    selection-color: #003388;
}

QLineEdit:hover {
    border-color: #999999;
}

QLineEdit:focus {
    border-color: #0066CC;
    outline: none;
}

QLineEdit:disabled {
    background-color: #F0F0F0;
    color: #AAAAAA;
}

QLineEdit:read-only {
    background-color: #F5F5F5;
    color: #555555;
}

/* ── Text Edit / Plain Text Edit ─────────────────────────────────────── */
QTextEdit, QPlainTextEdit {
    background-color: #FFFFFF;
    color: #1A1A1A;
    border: 1px solid #BCBCBC;
    border-radius: 3px;
    padding: 4px;
    selection-background-color: #CCE4FF;
    selection-color: #003388;
}

QTextEdit:focus, QPlainTextEdit:focus {
    border-color: #0066CC;
    outline: none;
}

/* ── Table Widget ────────────────────────────────────────────────────── */
QTableWidget, QTableView {
    background-color: #FFFFFF;
    alternate-background-color: #E8F4FD;
    gridline-color: #DCDCDC;
    border: 1px solid #BCBCBC;
    selection-background-color: #CCE4FF;
    selection-color: #003388;
    font-size: 9pt;
}

QTableWidget::item, QTableView::item {
    padding: 3px 6px;
    border: none;
}

QTableWidget::item:selected, QTableView::item:selected {
    background-color: #CCE4FF;
    color: #003388;
}

QTableWidget::item:hover, QTableView::item:hover {
    background-color: #EAF4FF;
}

/* Table header */
QHeaderView {
    background-color: #1F4E79;
    border: none;
}

QHeaderView::section {
    background-color: #1F4E79;
    color: #FFFFFF;
    font-weight: 600;
    font-size: 9pt;
    padding: 5px 8px;
    border: none;
    border-right: 1px solid #2D6499;
    border-bottom: 2px solid #163A5A;
}

QHeaderView::section:hover {
    background-color: #2D6499;
}

QHeaderView::section:pressed {
    background-color: #163A5A;
}

QHeaderView::section:checked {
    background-color: #2D6499;
}

QHeaderView::section:vertical {
    background-color: #E8E8E8;
    color: #1A1A1A;
    border-right: 1px solid #BCBCBC;
    border-bottom: 1px solid #DCDCDC;
    padding: 3px 6px;
    font-weight: normal;
}

/* Table corner button */
QTableCornerButton::section {
    background-color: #1F4E79;
    border: none;
    border-right: 1px solid #2D6499;
    border-bottom: 2px solid #163A5A;
}

/* ── Tab Widget / Tab Bar ─────────────────────────────────────────────── */
QTabWidget {
    background-color: #F0F0F0;
    border: none;
}

QTabWidget::pane {
    background-color: #FFFFFF;
    border: 1px solid #BCBCBC;
    border-top: none;
    border-bottom-left-radius: 3px;
    border-bottom-right-radius: 3px;
}

QTabWidget::tab-bar {
    alignment: left;
}

QTabBar {
    background-color: transparent;
    border-bottom: 1px solid #BCBCBC;
}

QTabBar::tab {
    background-color: #D4D4D4;
    color: #555555;
    border: 1px solid #BCBCBC;
    border-bottom: none;
    border-top-left-radius: 3px;
    border-top-right-radius: 3px;
    padding: 6px 16px;
    margin-right: 2px;
    min-width: 80px;
    font-size: 9pt;
}

QTabBar::tab:selected {
    background-color: #FFFFFF;
    color: #0066CC;
    border-color: #BCBCBC;
    border-bottom: 2px solid #0066CC;
    font-weight: 600;
}

QTabBar::tab:hover:!selected {
    background-color: #E8E8E8;
    color: #1A1A1A;
}

QTabBar::tab:disabled {
    color: #AAAAAA;
}

/* ── Dock Widget ─────────────────────────────────────────────────────── */
QDockWidget {
    titlebar-close-icon: url(none);
    titlebar-normal-icon: url(none);
    font-size: 9pt;
    color: #1A1A1A;
}

QDockWidget::title {
    background-color: #E0E0E0;
    color: #1A1A1A;
    font-weight: 600;
    padding: 5px 8px;
    border-bottom: 1px solid #BCBCBC;
    text-align: left;
}

QDockWidget::close-button,
QDockWidget::float-button {
    background-color: transparent;
    border: none;
    padding: 2px;
    border-radius: 2px;
}

QDockWidget::close-button:hover,
QDockWidget::float-button:hover {
    background-color: #CCE4FF;
}

QDockWidget::close-button:pressed,
QDockWidget::float-button:pressed {
    background-color: #99CAFF;
}

/* ── Status Bar ──────────────────────────────────────────────────────── */
QStatusBar {
    background-color: #1F4E79;
    color: #FFFFFF;
    border-top: 1px solid #163A5A;
    font-size: 8pt;
    padding: 0px 4px;
}

QStatusBar::item {
    border: none;
}

QStatusBar QLabel {
    background-color: transparent;
    color: #FFFFFF;
    padding: 0px 4px;
}

/* ── Scroll Bar ──────────────────────────────────────────────────────── */
QScrollBar:vertical {
    background-color: #F0F0F0;
    width: 10px;
    border: none;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background-color: #C0C0C0;
    min-height: 20px;
    border-radius: 5px;
    margin: 2px;
}

QScrollBar::handle:vertical:hover {
    background-color: #A0A0A0;
}

QScrollBar::handle:vertical:pressed {
    background-color: #0066CC;
}

QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {
    height: 0px;
    border: none;
    background: none;
}

QScrollBar:horizontal {
    background-color: #F0F0F0;
    height: 10px;
    border: none;
    margin: 0px;
}

QScrollBar::handle:horizontal {
    background-color: #C0C0C0;
    min-width: 20px;
    border-radius: 5px;
    margin: 2px;
}

QScrollBar::handle:horizontal:hover {
    background-color: #A0A0A0;
}

QScrollBar::handle:horizontal:pressed {
    background-color: #0066CC;
}

QScrollBar::add-line:horizontal,
QScrollBar::sub-line:horizontal {
    width: 0px;
    border: none;
    background: none;
}

QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
    background: none;
}

/* ── Splitter ────────────────────────────────────────────────────────── */
QSplitter::handle {
    background-color: #DCDCDC;
}

QSplitter::handle:horizontal {
    width: 4px;
}

QSplitter::handle:vertical {
    height: 4px;
}

QSplitter::handle:hover {
    background-color: #0066CC;
}

QSplitter::handle:pressed {
    background-color: #0044AA;
}

/* ── Progress Bar ────────────────────────────────────────────────────── */
QProgressBar {
    background-color: #E0E0E0;
    border: 1px solid #BCBCBC;
    border-radius: 3px;
    text-align: center;
    color: #1A1A1A;
    font-size: 8pt;
    min-height: 16px;
}

QProgressBar::chunk {
    background-color: #0066CC;
    border-radius: 2px;
}

/* ── Group Box ───────────────────────────────────────────────────────── */
QGroupBox {
    background-color: #FFFFFF;
    border: 1px solid #BCBCBC;
    border-radius: 4px;
    margin-top: 12px;
    padding: 8px 6px 6px 6px;
    font-size: 9pt;
    font-weight: 600;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0px 6px;
    color: #1F4E79;
    background-color: #FFFFFF;
    left: 8px;
}

/* ── Spin Box ────────────────────────────────────────────────────────── */
QSpinBox, QDoubleSpinBox {
    background-color: #FFFFFF;
    color: #1A1A1A;
    border: 1px solid #BCBCBC;
    border-radius: 3px;
    padding: 3px 6px;
    min-height: 22px;
}

QSpinBox:focus, QDoubleSpinBox:focus {
    border-color: #0066CC;
}

QSpinBox::up-button, QDoubleSpinBox::up-button {
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 16px;
    border-left: 1px solid #BCBCBC;
    border-bottom: 1px solid #BCBCBC;
    background-color: #E8E8E8;
    border-top-right-radius: 3px;
}

QSpinBox::down-button, QDoubleSpinBox::down-button {
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    width: 16px;
    border-left: 1px solid #BCBCBC;
    background-color: #E8E8E8;
    border-bottom-right-radius: 3px;
}

QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover,
QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover {
    background-color: #CCE4FF;
}

/* ── Check Box ───────────────────────────────────────────────────────── */
QCheckBox {
    spacing: 6px;
    color: #1A1A1A;
}

QCheckBox::indicator {
    width: 14px;
    height: 14px;
    border: 1px solid #BCBCBC;
    border-radius: 2px;
    background-color: #FFFFFF;
}

QCheckBox::indicator:hover {
    border-color: #0066CC;
    background-color: #EAF4FF;
}

QCheckBox::indicator:checked {
    background-color: #0066CC;
    border-color: #0044AA;
}

QCheckBox::indicator:checked:hover {
    background-color: #0088FF;
}

QCheckBox::indicator:disabled {
    background-color: #F0F0F0;
    border-color: #D0D0D0;
}

/* ── Radio Button ────────────────────────────────────────────────────── */
QRadioButton {
    spacing: 6px;
    color: #1A1A1A;
}

QRadioButton::indicator {
    width: 14px;
    height: 14px;
    border: 1px solid #BCBCBC;
    border-radius: 7px;
    background-color: #FFFFFF;
}

QRadioButton::indicator:hover {
    border-color: #0066CC;
    background-color: #EAF4FF;
}

QRadioButton::indicator:checked {
    background-color: #0066CC;
    border-color: #0044AA;
}

/* ── Slider ──────────────────────────────────────────────────────────── */
QSlider::groove:horizontal {
    height: 4px;
    background-color: #D0D0D0;
    border-radius: 2px;
    margin: 0px 0px;
}

QSlider::handle:horizontal {
    background-color: #0066CC;
    border: 1px solid #0044AA;
    width: 14px;
    height: 14px;
    margin: -5px 0px;
    border-radius: 7px;
}

QSlider::handle:horizontal:hover {
    background-color: #0088FF;
}

QSlider::sub-page:horizontal {
    background-color: #0066CC;
    border-radius: 2px;
}

/* ── Tool Tips ───────────────────────────────────────────────────────── */
QToolTip {
    background-color: #FFFFCC;
    color: #1A1A1A;
    border: 1px solid #BCBCBC;
    padding: 4px 6px;
    border-radius: 2px;
    font-size: 8pt;
}

/* ── Frame ───────────────────────────────────────────────────────────── */
QFrame[frameShape="4"],
QFrame[frameShape="5"] {
    color: #BCBCBC;
}

/* ── Label ───────────────────────────────────────────────────────────── */
QLabel {
    background-color: transparent;
    color: #1A1A1A;
}

/* ── Scroll Area ─────────────────────────────────────────────────────── */
QScrollArea {
    background-color: #FFFFFF;
    border: 1px solid #BCBCBC;
}

QScrollArea > QWidget > QWidget {
    background-color: #FFFFFF;
}

/* ── Tree Widget / List Widget ───────────────────────────────────────── */
QTreeWidget, QListWidget {
    background-color: #FFFFFF;
    alternate-background-color: #F5F5F5;
    border: 1px solid #BCBCBC;
    selection-background-color: #CCE4FF;
    selection-color: #003388;
}

QTreeWidget::item, QListWidget::item {
    padding: 2px 4px;
}

QTreeWidget::item:hover, QListWidget::item:hover {
    background-color: #EAF4FF;
}

QTreeWidget::item:selected, QListWidget::item:selected {
    background-color: #CCE4FF;
    color: #003388;
}

QTreeWidget::branch:has-children:!has-siblings:closed,
QTreeWidget::branch:closed:has-children:has-siblings {
    border-image: none;
    image: none;
}

/* ── Dialog Buttons ──────────────────────────────────────────────────── */
QDialogButtonBox QPushButton {
    min-width: 80px;
}

/* ── Message Box ─────────────────────────────────────────────────────── */
QMessageBox {
    background-color: #F0F0F0;
}

QMessageBox QLabel {
    color: #1A1A1A;
    min-width: 300px;
}
)";
}

// ── Apply ─────────────────────────────────────────────────────────────────────

void AppTheme::apply() {
    QApplication* app = qobject_cast<QApplication*>(QApplication::instance());
    if (!app)
        return;

    // Set application-wide font
    app->setFont(fontDefault());

    // Apply the stylesheet
    app->setStyleSheet(stylesheet());
}

#pragma once
#include <QString>
#include <QColor>
#include <QFont>

class AppTheme {
public:
    // Apply the theme to the QApplication instance
    static void apply();

    // ── Color palette ────────────────────────────────────────────────────────
    static QColor bgMain()      { return QColor(0xF0, 0xF0, 0xF0); }  // light gray
    static QColor bgPanel()     { return QColor(0xFF, 0xFF, 0xFF); }  // white
    static QColor bgRibbon()    { return QColor(0xE8, 0xE8, 0xE8); }  // ribbon bg
    static QColor bgTabActive() { return QColor(0xFF, 0xFF, 0xFF); }  // active tab
    static QColor bgTabInact()  { return QColor(0xD4, 0xD4, 0xD4); }  // inactive tab
    static QColor accent()      { return QColor(0x00, 0x66, 0xCC); }  // blue
    static QColor accentHover() { return QColor(0x00, 0x88, 0xFF); }
    static QColor accentPress() { return QColor(0x00, 0x44, 0xAA); }
    static QColor textPrimary() { return QColor(0x1A, 0x1A, 0x1A); }
    static QColor textSec()     { return QColor(0x55, 0x55, 0x55); }
    static QColor border()      { return QColor(0xBC, 0xBC, 0xBC); }
    static QColor borderDark()  { return QColor(0x99, 0x99, 0x99); }
    static QColor tableHeader() { return QColor(0x1F, 0x4E, 0x79); }  // dark blue
    static QColor tableAlt()    { return QColor(0xE8, 0xF4, 0xFD); }  // light blue
    static QColor hoverBg()     { return QColor(0xEA, 0xF4, 0xFF); }  // hover highlight
    static QColor selectBg()    { return QColor(0xCC, 0xE4, 0xFF); }  // selection highlight
    static QColor success()     { return QColor(0x28, 0xA7, 0x45); }  // green
    static QColor warning()     { return QColor(0xED, 0x8B, 0x00); }  // orange

    // ── Fonts ────────────────────────────────────────────────────────────────
    static QFont fontDefault();   // Segoe UI 9pt
    static QFont fontSmall();     // Segoe UI 8pt
    static QFont fontTitle();     // Segoe UI Semibold 11pt
    static QFont fontMono();      // Consolas 9pt

    // ── Full QSS stylesheet ──────────────────────────────────────────────────
    static QString stylesheet();
};

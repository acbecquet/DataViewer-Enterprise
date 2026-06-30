#pragma once
#include <QString>
#include <QColor>
#include <QFont>
#include <QIcon>
#include <QVector>

class AppTheme {
public:
    // Apply the theme to the QApplication instance
    static void apply();

    // ── Surfaces ─────────────────────────────────────────────────────────────
    static QColor surfaceApp()      { return QColor(0xF5, 0xF6, 0xF8); }
    static QColor surfacePanel()    { return QColor(0xFF, 0xFF, 0xFF); }
    static QColor surfaceRibbon()   { return QColor(0xEC, 0xEE, 0xF1); }
    static QColor surfaceStatusBar(){ return QColor(0xEC, 0xEE, 0xF1); }
    static QColor surfaceTabActive(){ return QColor(0xFF, 0xFF, 0xFF); }
    static QColor surfaceTabInact() { return QColor(0xD4, 0xD8, 0xDC); }

    // ── Borders ──────────────────────────────────────────────────────────────
    static QColor borderSubtle()    { return QColor(0xE4, 0xE6, 0xEA); }
    static QColor borderDefault()   { return QColor(0xCF, 0xD3, 0xD8); }
    static QColor borderStrong()    { return QColor(0xA0, 0xA6, 0xAE); }

    // ── Accent (kept identical to old #0066CC) ───────────────────────────────
    static QColor accent()          { return QColor(0x00, 0x66, 0xCC); }
    static QColor accentHover()     { return QColor(0x00, 0x88, 0xFF); }
    static QColor accentPress()     { return QColor(0x00, 0x44, 0xAA); }
    static QColor accentSubtle()    { return QColor(0xE8, 0xF2, 0xFC); }

    // ── Text ─────────────────────────────────────────────────────────────────
    static QColor textPrimary()     { return QColor(0x1A, 0x1D, 0x21); }
    static QColor textSecondary()   { return QColor(0x5C, 0x63, 0x6A); }
    static QColor textMuted()       { return QColor(0x8A, 0x90, 0x99); }

    // ── Table header — flatter slate instead of navy ─────────────────────────
    static QColor tableHeader()     { return QColor(0x2C, 0x3E, 0x50); }
    static QColor tableHeaderHover(){ return QColor(0x3F, 0x55, 0x6B); }

    // ── Selection / hover ────────────────────────────────────────────────────
    static QColor hoverBg()         { return QColor(0xE8, 0xF2, 0xFC); }
    static QColor selectBg()        { return QColor(0xCC, 0xE4, 0xFF); }
    static QColor tableAlt()        { return QColor(0xE8, 0xF2, 0xFC); }

    // ── Semantic ─────────────────────────────────────────────────────────────
    static QColor success()         { return QColor(0x16, 0xA3, 0x4A); }
    static QColor warning()         { return QColor(0xD9, 0x77, 0x06); }
    static QColor error()           { return QColor(0xDC, 0x26, 0x26); }

    // ── Plot series palette — single source of truth for ALL chart colors ───
    // Never reuse a color within a plot. Curated 20-color, no-yellow,
    // projector-safe palette (shared with the sensory radar); golden-angle
    // generation beyond 20, unique through ~100 entries — far more than any
    // realistic plot needs.
    static QColor          seriesColor(int idx);   // the idx-th distinct color
    static QVector<QColor> seriesColors(int n);    // first n distinct colors
    // A distinct shade of a base hue for grouped charts (per-file families).
    // Only value/saturation vary (hue preserved), so shades of two different
    // base hues can never collide; the value band stays projector-safe.
    static QColor          shade(const QColor& base, int idx, int count);

    // ── BACKWARD COMPAT — legacy accessors that resolve to new tokens ───────
    // Keep these so existing call sites don't have to be updated all at once.
    static QColor bgMain()      { return surfaceApp(); }
    static QColor bgPanel()     { return surfacePanel(); }
    static QColor bgRibbon()    { return surfaceRibbon(); }
    static QColor bgTabActive() { return surfaceTabActive(); }
    static QColor bgTabInact()  { return surfaceTabInact(); }
    static QColor textSec()     { return textSecondary(); }
    static QColor border()      { return borderDefault(); }
    static QColor borderDark()  { return borderStrong(); }

    // ── Spacing scale (px) ───────────────────────────────────────────────────
    static int spaceXs()  { return 4; }
    static int spaceSm()  { return 8; }
    static int spaceMd()  { return 12; }
    static int spaceLg()  { return 16; }
    static int spaceXl()  { return 24; }
    static int space2Xl() { return 32; }

    // ── Radius scale (px) ────────────────────────────────────────────────────
    static int radiusControl() { return 4; }  // buttons, inputs
    static int radiusPanel()   { return 6; }  // cards, groupboxes
    static int radiusDialog()  { return 8; }  // dialogs

    // ── Fonts ────────────────────────────────────────────────────────────────
    static QFont fontDefault();   // Segoe UI 9pt
    static QFont fontSmall();     // Segoe UI 8pt
    static QFont fontLabel();     // Segoe UI 10pt
    static QFont fontSection();   // Segoe UI 11pt semibold
    static QFont fontPageTitle(); // Segoe UI 12pt semibold
    static QFont fontTitle();     // Segoe UI Semibold 11pt (LEGACY → fontSection)
    static QFont fontMono();      // Consolas 9pt

    // -- Font-metric sizing helpers (v2.7.0 responsive) -----------------------
    // Pure functions of the active font -- no caching, no state -- so they re-
    // evaluate correctly after a DPI/scale change (Qt re-polishes the app font).
    // Use these instead of hard-coded pixel heights so text-bearing controls
    // grow with the font under OS text-scaling rather than clipping.
    static int lineUnit(const QFont& f = fontDefault());                 // one text line's height, px
    static int controlHeight(const QFont& f = fontDefault(), int vPad = 6); // single-line control height
    static int em(qreal n, const QFont& f = fontDefault());              // n average-char-widths, px

    // ── Icon helper (Phase 1 Task 3 implements this) ─────────────────────────
    // Returns a QIcon for the given Lucide icon name (without .svg extension).
    // Loads from <resourcePath>/icons/<name>.svg. Caches the result.
    // Returns a null QIcon and logs a warning if the file is missing.
    static QIcon icon(const QString& name);

    // ── Full QSS stylesheet (now built from tokens) ──────────────────────────
    static QString stylesheet();
};

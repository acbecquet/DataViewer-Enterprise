#pragma once
#include <QWidget>
#include <QTabWidget>
#include <QToolButton>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QLabel>
#include <QFrame>
#include <QFont>
#include <QString>
#include <QList>
#include <functional>

// ─── RibbonGroup ──────────────────────────────────────────────────────────────
// A named group of buttons within a ribbon tab (e.g. "File", "Reports", "View").
// Layout (inside a fixed-height container):
//   [ large-button | large-button | ... ]  <- m_largeButtonLayout (QHBoxLayout)
//   [ small-button ]                        <- m_smallButtonLayout (QVBoxLayout)
//
// The group title label (m_titleLabel) still exists for compatibility but is no
// longer shown or laid out in either mode -- the owner found the bottom titles
// redundant, so the title row + its separator were removed to shorten the ribbon.
// The group is separated from the next group by a thin vertical line (QFrame).

class RibbonGroup : public QWidget {
    Q_OBJECT
public:
    explicit RibbonGroup(const QString& title, QWidget* parent = nullptr);

    // Add a large tool button: icon on top (32x32), text below (8pt, word-wrap).
    // The button is 80px wide × 76px tall.
    QToolButton* addLargeButton(const QString& text,
                                const QIcon&   icon,
                                const QString& tooltip = "");

    // -- Large-button sizing/wrap helpers (font-derived; spec v2.7.0 §4) ------
    // Standard large-button width. Height/wrap derive from the active font so
    // the button grows under OS text-scaling instead of clipping to 3 lines.
    static constexpr int kLargeButtonWidth = 80;

    // Owner spec (2026-07-01): exactly 5px of band padding above and below the
    // button content, in BOTH full and compact modes.
    static constexpr int kBandPadding = 5;

    // Icons-only button side and the tight group band it produces
    // (kBandPadding + kCompactButtonSide + kBandPadding).
    static constexpr int kCompactButtonSide = 36;
    static int compactGroupHeight();

    // Width this group needs to render every VISIBLE button with its label
    // (full mode), computed from text metrics so it is independent of the
    // mode currently applied. Hidden (mode-swapped) buttons contribute 0.
    int fullModeNeededWidth() const;

    // The font the large-button LABEL is actually rendered with (8pt Segoe UI,
    // matching the QSS), used for all wrap measurement.
    static QFont largeButtonFont();

    // Real available text width inside the button: width - frame - padding.
    static int largeButtonTextWidth(const QFont& f = largeButtonFont());

    // Minimum button height that holds the 32px icon band + up to two lines of
    // label at font `f` (replaces the hard-coded 76px).
    static int largeButtonHeight(const QFont& f = largeButtonFont());

    // Split `text` into at most two lines that each fit largeButtonTextWidth(f).
    // Picks the split point closest to a balanced halfway split among the splits
    // that fit; if no two-line split fits (very long single word at large
    // scale), returns the text unchanged (the button grows in width instead).
    static QString wrapLabelText(const QString& text, const QFont& f = largeButtonFont());

    // Font-derived minimum group height: top+bottom margin + large-button band.
    // The title row (separator + title line) is no longer included, so the group
    // is shorter than before. Grows under text-scaling via largeButtonHeight().
    static int groupMinimumHeight(const QFont& f = largeButtonFont());

    // Add a small tool button: icon on left (16x16), text on right (9pt).
    // The button spans full group width and is 24px tall.
    QToolButton* addSmallButton(const QString& text,
                                const QIcon&   icon,
                                const QString& tooltip = "");

    // Insert a thin vertical separator line between large buttons.
    void addSeparator();

    // Add an arbitrary widget into the large-button row.
    void addWidget(QWidget* w);

    // Toggle icons-only compact mode (hides labels and group title).
    void setCompactMode(bool compact);

private:
    // Outer vertical layout of the whole group, and the horizontal layout of its
    // content row. Held so setCompactMode can tighten their margins/spacing in
    // icons-only mode (packs groups closer) without affecting full mode.
    QVBoxLayout* m_outerVBox;
    QHBoxLayout* m_contentHBox;

    // Top area: large buttons sit side-by-side horizontally
    QHBoxLayout* m_largeButtonLayout;

    // Top-left area (inside the large row): stack of small buttons
    QVBoxLayout* m_smallButtonLayout;
    QWidget*     m_smallButtonContainer;

    // Whether any small buttons have been added yet
    bool         m_hasSmallButtons;

    QLabel*      m_titleLabel;

    // Compact-mode state + tracked large buttons
    bool                 m_compactMode  = false;
    QList<QToolButton*>  m_largeButtons;
};


// ─── RibbonTab ────────────────────────────────────────────────────────────────
// One page of the ribbon.  Contains an arbitrary number of RibbonGroups
// laid out horizontally.

class RibbonTab : public QWidget {
    Q_OBJECT
public:
    explicit RibbonTab(QWidget* parent = nullptr);

    // Create and return a new group with the given title.
    RibbonGroup* addGroup(const QString& title);

    QList<RibbonGroup*> groups() const { return m_groups; }

    // Sum of the visible groups' full-mode width needs plus tab margins.
    int fullModeNeededWidth() const;

    // Toggle icons-only compact mode on all groups in this tab.
    void setCompactMode(bool compact);

private:
    QHBoxLayout*        m_layout;
    QList<RibbonGroup*> m_groups;
    bool                m_compactMode = false;
};


// ─── RibbonWidget ─────────────────────────────────────────────────────────────
// The full ribbon bar (fixed height ≈ 100 px).  Wraps a QTabWidget whose pages
// are RibbonTab instances.  The tab bar uses flat tabs with a blue underline
// for the active tab.

class RibbonWidget : public QWidget {
    Q_OBJECT
public:
    explicit RibbonWidget(QWidget* parent = nullptr);

    // Create a new tab and return it so the caller can populate it.
    RibbonTab* addTab(const QString& label);

    void setCurrentTab(int index);
    int  currentTab() const;

    QTabWidget* tabWidget() { return m_tabs; }

    // Toggle icons-only compact mode (cascades to all tabs/groups).
    void setCompactMode(bool compact);
    bool isCompactMode() const { return m_compactMode; }

    // Widest full-labeled width across ALL tabs (0 when no tabs yet). Compact
    // mode is a single global flag, so the decision is driven by the widest
    // tab -- otherwise the icon style would flip when the user switches tabs.
    int fullModeNeededWidth() const;

    // Recompute compact vs full from the current width: labels persist until
    // they genuinely no longer fit, and only then does icons-only engage
    // (owner spec 2026-07-01). Runs on resize and tab switches; MainWindow
    // also calls it after updateRibbonForMode() changes button visibility.
    void updateResponsiveMode();

    // The ribbon's height is computed EXACTLY (live tab-bar hint + the QSS
    // pane borders + current page hint + bottom rule) rather than trusting
    // QTabWidget's sizeHint, whose style-frame allowance exceeds our 1px QSS
    // pane borders and used to leak ~2px of dead band around the buttons.
    QSize sizeHint() const override;
    QSize minimumSizeHint() const override;

protected:
    void resizeEvent(QResizeEvent* e) override;

private:
    int exactHeight() const;

    // Resolve the RibbonTab on tab page `index`, unwrapping the ScrollHost the
    // page is stored in. Returns nullptr for an out-of-range/foreign page.
    RibbonTab* tabAt(int index) const;

    QTabWidget*  m_tabs;
    QVBoxLayout* m_layout;
    bool         m_compactMode = false;
};

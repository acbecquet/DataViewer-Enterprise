#pragma once
#include <QWidget>
#include <QTabWidget>
#include <QToolButton>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QLabel>
#include <QFrame>
#include <QString>
#include <QList>
#include <functional>

// ─── RibbonGroup ──────────────────────────────────────────────────────────────
// A named group of buttons within a ribbon tab (e.g. "File", "Reports", "View").
// Layout (inside a fixed-height ~90px container):
//   [ large-button | large-button | ... ]  <- m_largeButtonLayout (QHBoxLayout)
//   [ small-button ]                        <- m_smallButtonLayout (QVBoxLayout)
//   ─────────────────────────────────────  <- thin separator line
//   [         Group Title                ] <- m_titleLabel, small centered text
//
// The group is separated from the next group by a thin vertical line (QFrame).

class RibbonGroup : public QWidget {
    Q_OBJECT
public:
    explicit RibbonGroup(const QString& title, QWidget* parent = nullptr);

    // Add a large tool button: icon on top (32x32), text below (8pt, word-wrap).
    // The button is 56px wide × 70px tall.
    QToolButton* addLargeButton(const QString& text,
                                const QIcon&   icon,
                                const QString& tooltip = "");

    // Add a small tool button: icon on left (16x16), text on right (9pt).
    // The button spans full group width and is 24px tall.
    QToolButton* addSmallButton(const QString& text,
                                const QIcon&   icon,
                                const QString& tooltip = "");

    // Insert a thin vertical separator line between large buttons.
    void addSeparator();

    // Add an arbitrary widget into the large-button row.
    void addWidget(QWidget* w);

private:
    // Top area: large buttons sit side-by-side horizontally
    QHBoxLayout* m_largeButtonLayout;

    // Top-left area (inside the large row): stack of small buttons
    QVBoxLayout* m_smallButtonLayout;
    QWidget*     m_smallButtonContainer;

    // Whether any small buttons have been added yet
    bool         m_hasSmallButtons;

    QLabel*      m_titleLabel;
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

private:
    QHBoxLayout*        m_layout;
    QList<RibbonGroup*> m_groups;
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

private:
    QTabWidget*  m_tabs;
    QVBoxLayout* m_layout;
};

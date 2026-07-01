#include "RibbonWidget.h"

#include <QToolButton>
#include <QLabel>
#include <QFrame>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QTabWidget>
#include <QTabBar>
#include <QIcon>
#include <QSizePolicy>
#include <QFont>
#include <QPainter>
#include <QStyleOption>
#include <QFontMetrics>
#include "ScrollHost.h"
#include "utils/AppTheme.h"
#include <QStyle>
#include <QStyleOptionToolButton>

// ═══════════════════════════════════════════════════════════════════════════════
// RibbonGroup
// ═══════════════════════════════════════════════════════════════════════════════

int RibbonGroup::groupMinimumHeight(const QFont& f)
{
    // 4px top + 4px bottom margin + button band + 2px slack. The group-title
    // row (1px separator + title line) is no longer shown in either mode, so
    // its height is not reserved -- the ribbon is shorter by exactly that band.
    return 4 + largeButtonHeight(f) + 4 + 2;
}

RibbonGroup::RibbonGroup(const QString& title, QWidget* parent)
    : QWidget(parent)
    , m_hasSmallButtons(false)
{
    setObjectName("RibbonGroup");

    // Font-derived minimum so the group-title row never clips under scaling;
    // equals ~98 at standard scale. Vertical policy stays Fixed.
    setMinimumHeight(groupMinimumHeight());
    setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Fixed);

    // ── Outer vertical layout ─────────────────────────────────────────────────
    //  Row 0 (stretch): content area  – large buttons  + optional small buttons
    //
    // The group-title row (a thin separator + a centered title label) was
    // removed: the owner found the titles redundant (the buttons are already
    // labelled) and they wasted vertical space. m_titleLabel still exists (it's
    // referenced by setCompactMode) but is never added to a layout and stays
    // hidden, so it reserves NO height in either mode -- the ribbon is shorter.

    QVBoxLayout* outerVBox = new QVBoxLayout(this);
    outerVBox->setContentsMargins(4, 4, 4, 4);
    outerVBox->setSpacing(0);
    m_outerVBox = outerVBox;

    // ── Content area (large buttons + small buttons side-by-side) ─────────────
    QWidget* contentArea = new QWidget(this);
    contentArea->setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Expanding);

    QHBoxLayout* contentHBox = new QHBoxLayout(contentArea);
    contentHBox->setContentsMargins(0, 0, 0, 0);
    contentHBox->setSpacing(2);
    m_contentHBox = contentHBox;

    // Small-button container (left side, stacked vertically)
    m_smallButtonContainer = new QWidget(contentArea);
    m_smallButtonContainer->setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Expanding);
    m_smallButtonContainer->setVisible(false);  // hidden until first small button

    m_smallButtonLayout = new QVBoxLayout(m_smallButtonContainer);
    m_smallButtonLayout->setContentsMargins(0, 0, 0, 0);
    m_smallButtonLayout->setSpacing(1);
    m_smallButtonLayout->addStretch(1);

    // Large-button row (right/main side)
    QWidget* largeBtnContainer = new QWidget(contentArea);
    largeBtnContainer->setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Expanding);

    m_largeButtonLayout = new QHBoxLayout(largeBtnContainer);
    m_largeButtonLayout->setContentsMargins(0, 0, 0, 0);
    m_largeButtonLayout->setSpacing(2);
    m_largeButtonLayout->setAlignment(Qt::AlignLeft | Qt::AlignVCenter);

    contentHBox->addWidget(m_smallButtonContainer);
    contentHBox->addWidget(largeBtnContainer);

    outerVBox->addWidget(contentArea, 1);

    // ── Group title label (created but intentionally NOT laid out) ─────────────
    // Kept as a hidden member so setCompactMode's title toggle stays valid, but
    // it is neither added to outerVBox nor given a separator row, so it reserves
    // no vertical space. The title is shown in NEITHER full nor compact mode.
    m_titleLabel = new QLabel(title, this);
    m_titleLabel->setVisible(false);

    // ── Right-side vertical border (thin separator between groups) ─────────────
    // Drawn via QSS on the widget itself.
    setStyleSheet(
        "RibbonGroup {"
        "  background-color: transparent;"
        "  border-right: 1px solid #C8C8C8;"
        "}"
    );
}

QFont RibbonGroup::largeButtonFont()
{
    // Matches the QSS `font-size: 8pt; font-family: 'Segoe UI'` below.
    QFont f("Segoe UI", 8);
    return f;
}

int RibbonGroup::largeButtonTextWidth(const QFont& f)
{
    // Button width minus: 1px QSS border each side, 2px QSS padding each side,
    // plus a 2px QToolButton internal text inset each side. Floor at a few px
    // so a tiny font never yields a non-positive width.
    Q_UNUSED(f);
    const int frameAndPad = 2 * (1 /*border*/ + 2 /*padding*/ + 2 /*tool inset*/);
    return qMax(8, kLargeButtonWidth - frameAndPad);   // 80 - 10 = 70px standard
}

int RibbonGroup::largeButtonHeight(const QFont& f)
{
    // Two label lines + the 32px icon band + the QSS vertical padding/border.
    // AppTheme::lineUnit(f) is QFontMetrics(f).height(); two of them is the
    // 2-line label block. 32 = icon, +10 = 2px padding + 1px border (x2) + 4
    // spacing. At 8pt Segoe UI this evaluates to ~76, preserving today's look;
    // it grows with the font under scaling.
    const int iconBand = 32;
    const int chrome = 10;
    return iconBand + 2 * AppTheme::lineUnit(f) + chrome;
}

QString RibbonGroup::wrapLabelText(const QString& text, const QFont& f)
{
    const QFontMetrics fm(f);
    const int maxW = largeButtonTextWidth(f);

    if (fm.horizontalAdvance(text) <= maxW)
        return text;   // already fits on one line

    // Candidate split points are the space positions. Choose the split where
    // BOTH halves fit AND the split is closest to the visual midpoint, so the
    // two lines are balanced (not ragged splits that leave line 2 wide enough
    // to re-wrap).
    const int mid = text.length() / 2;
    int bestSplit = -1;
    int bestDist  = text.length() + 1;
    for (int i = 1; i < text.length(); ++i) {
        if (text[i] != ' ')
            continue;
        const QString l1 = text.left(i);
        const QString l2 = text.mid(i + 1);
        if (fm.horizontalAdvance(l1) <= maxW && fm.horizontalAdvance(l2) <= maxW) {
            const int dist = qAbs(i - mid);
            if (dist < bestDist) {
                bestDist  = dist;
                bestSplit = i;
            }
        }
    }

    if (bestSplit < 0)
        return text;   // no fitting 2-line split: let the button grow in width

    return text.left(bestSplit) + "\n" + text.mid(bestSplit + 1);
}

QToolButton* RibbonGroup::addLargeButton(const QString& text,
                                         const QIcon&   icon,
                                         const QString& tooltip)
{
    QToolButton* btn = new QToolButton(this);
    btn->setText(text);
    btn->setIcon(icon);
    btn->setIconSize(QSize(32, 32));
    btn->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
    btn->setToolTip(tooltip.isEmpty() ? text : tooltip);
    btn->setAutoRaise(true);
    btn->setFocusPolicy(Qt::NoFocus);

    // Balanced <=2-line wrap measured against the real available text width and
    // the actual label font (8pt), so QToolButton never spawns a clipped 3rd
    // row (the "View Raw Data" overflow bug).
    const QFont labelFont = largeButtonFont();
    btn->setText(wrapLabelText(text, labelFont));

    // Min-size, not fixed: standard scale renders 80x76 (no visual regression);
    // under text-scaling the button grows instead of clipping. Vertical policy
    // Fixed keeps the ribbon row tidy; horizontal Minimum lets a too-long
    // single word widen the button rather than re-wrap to a 3rd line.
    btn->setMinimumSize(kLargeButtonWidth, largeButtonHeight(labelFont));
    btn->setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Fixed);

    btn->setStyleSheet(R"(
        QToolButton {
            background-color: transparent;
            border: 1px solid transparent;
            border-radius: 3px;
            padding: 2px;
            font-family: 'Segoe UI';
            font-size: 8pt;
            color: #1A1A1A;
            text-align: center;
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
    )");

    m_largeButtonLayout->addWidget(btn);
    m_largeButtons.append(btn);
    return btn;
}

void RibbonGroup::setCompactMode(bool compact)
{
    if (m_compactMode == compact) return;
    m_compactMode = compact;
    for (QToolButton* b : m_largeButtons) {
        b->setToolButtonStyle(compact ? Qt::ToolButtonIconOnly
                                      : Qt::ToolButtonTextUnderIcon);
        if (compact) {
            b->setMinimumSize(36, 36);
            b->setMaximumSize(36, 36);
            b->setIconSize(QSize(20, 20));
        } else {
            b->setMaximumSize(QWIDGETSIZE_MAX, QWIDGETSIZE_MAX);
            b->setMinimumSize(kLargeButtonWidth, largeButtonHeight(largeButtonFont()));
            b->setIconSize(QSize(32, 32));
        }
    }

    // Pack the icon-only ribbon tighter: in compact mode the per-group side
    // margins and the intra-group button spacing dominate the inter-group gap
    // (there is no group title to justify the width), so shrink them; restore
    // the roomier full-mode values otherwise. These are the ONLY horizontal
    // margins/spacing between one group's icons and the next.
    if (compact) {
        m_outerVBox->setContentsMargins(1, 4, 1, 4);
        m_contentHBox->setSpacing(0);
        m_largeButtonLayout->setSpacing(1);
    } else {
        m_outerVBox->setContentsMargins(4, 4, 4, 4);
        m_contentHBox->setSpacing(2);
        m_largeButtonLayout->setSpacing(2);
    }

    // Group titles are shown in NEITHER mode now (owner: redundant, wasteful);
    // keep the label hidden regardless of compact state.
    if (m_titleLabel) m_titleLabel->setVisible(false);
    updateGeometry();
}

QToolButton* RibbonGroup::addSmallButton(const QString& text,
                                         const QIcon&   icon,
                                         const QString& tooltip)
{
    QToolButton* btn = new QToolButton(this);
    btn->setText(text);
    btn->setIcon(icon);
    btn->setIconSize(QSize(16, 16));
    btn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    btn->setToolTip(tooltip.isEmpty() ? text : tooltip);
    btn->setAutoRaise(true);
    btn->setFocusPolicy(Qt::NoFocus);

    // Height fixed at 24px; stretch full width
    btn->setFixedHeight(24);
    btn->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);

    btn->setStyleSheet(R"(
        QToolButton {
            background-color: transparent;
            border: 1px solid transparent;
            border-radius: 2px;
            padding: 1px 4px;
            font-family: 'Segoe UI';
            font-size: 9pt;
            color: #1A1A1A;
            text-align: left;
        }
        QToolButton:hover {
            background-color: #CCE4FF;
            border-color: #99CAFF;
        }
        QToolButton:pressed {
            background-color: #99CAFF;
            border-color: #0066CC;
        }
        QToolButton:disabled {
            color: #AAAAAA;
        }
    )");

    // Insert before the trailing stretch
    if (!m_hasSmallButtons) {
        m_smallButtonContainer->setVisible(true);
        m_hasSmallButtons = true;
        // Remove the trailing stretch so we can add the button before it,
        // then re-add the stretch.
        QLayoutItem* stretch = m_smallButtonLayout->takeAt(
            m_smallButtonLayout->count() - 1);
        m_smallButtonLayout->addWidget(btn);
        m_smallButtonLayout->addItem(stretch);
    } else {
        // Find the stretch item (always last) and insert before it.
        int stretchIdx = m_smallButtonLayout->count() - 1;
        m_smallButtonLayout->insertWidget(stretchIdx, btn);
    }

    return btn;
}

void RibbonGroup::addSeparator()
{
    QFrame* sep = new QFrame(this);
    sep->setFrameShape(QFrame::VLine);
    sep->setFrameShadow(QFrame::Plain);
    sep->setFixedWidth(1);
    sep->setSizePolicy(QSizePolicy::Fixed, QSizePolicy::Expanding);
    sep->setStyleSheet("color: #C8C8C8; background-color: #C8C8C8; border: none;");
    m_largeButtonLayout->addWidget(sep);
}

void RibbonGroup::addWidget(QWidget* w)
{
    m_largeButtonLayout->addWidget(w);
}


// ═══════════════════════════════════════════════════════════════════════════════
// RibbonTab
// ═══════════════════════════════════════════════════════════════════════════════

RibbonTab::RibbonTab(QWidget* parent)
    : QWidget(parent)
{
    setObjectName("RibbonTab");
    setStyleSheet("RibbonTab { background-color: #E8E8E8; }");

    m_layout = new QHBoxLayout(this);
    m_layout->setContentsMargins(4, 2, 4, 2);
    m_layout->setSpacing(0);
    m_layout->setAlignment(Qt::AlignLeft | Qt::AlignVCenter);
    m_layout->addStretch(1);
}

RibbonGroup* RibbonTab::addGroup(const QString& title)
{
    RibbonGroup* grp = new RibbonGroup(title, this);

    // Insert before the trailing stretch
    int stretchIdx = m_layout->count() - 1;
    m_layout->insertWidget(stretchIdx, grp);

    m_groups.append(grp);
    return grp;
}

void RibbonTab::setCompactMode(bool compact)
{
    if (m_compactMode == compact) return;
    m_compactMode = compact;
    for (RibbonGroup* g : m_groups) g->setCompactMode(compact);
}


// ═══════════════════════════════════════════════════════════════════════════════
// RibbonWidget
// ═══════════════════════════════════════════════════════════════════════════════

int RibbonWidget::ribbonMinimumHeight(const QFont& tabFont, const QFont& btnFont)
{
    const int tabBarH = AppTheme::lineUnit(tabFont) + 14;
    return tabBarH + RibbonGroup::groupMinimumHeight(btnFont) + 1;
}

RibbonWidget::RibbonWidget(QWidget* parent)
    : QWidget(parent)
{
    setObjectName("RibbonWidget");
    setMinimumHeight(ribbonMinimumHeight());
    setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);

    m_layout = new QVBoxLayout(this);
    m_layout->setContentsMargins(0, 0, 0, 0);
    m_layout->setSpacing(0);

    m_tabs = new QTabWidget(this);
    m_tabs->setObjectName("RibbonTabWidget");
    m_tabs->setDocumentMode(false);
    m_tabs->setTabPosition(QTabWidget::North);
    m_tabs->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);

    // ── Ribbon-specific QSS ───────────────────────────────────────────────────
    // Override global theme for the ribbon tab widget to get the flat
    // engineering-app look with the blue underline on the active tab.
    m_tabs->setStyleSheet(R"(
        QTabWidget#RibbonTabWidget {
            background-color: #E8E8E8;
            border: none;
        }
        QTabWidget#RibbonTabWidget::pane {
            background-color: #E8E8E8;
            border: none;
            border-top: 1px solid #BCBCBC;
            border-bottom: 1px solid #BCBCBC;
        }
        QTabWidget#RibbonTabWidget > QTabBar::tab {
            background-color: transparent;
            color: #333333;
            border: none;
            border-bottom: 2px solid transparent;
            padding: 5px 14px 4px 14px;
            margin-right: 1px;
            font-family: 'Segoe UI';
            font-size: 9pt;
            min-width: 60px;
        }
        QTabWidget#RibbonTabWidget > QTabBar::tab:selected {
            background-color: #E8E8E8;
            color: #0066CC;
            border-bottom: 2px solid #0066CC;
            font-weight: 600;
        }
        QTabWidget#RibbonTabWidget > QTabBar::tab:hover:!selected {
            background-color: #D4EAFF;
            color: #003388;
            border-bottom: 2px solid #99CAFF;
        }
        QTabWidget#RibbonTabWidget > QTabBar {
            background-color: #E8E8E8;
            border-bottom: none;
        }
        QTabBar::scroller {
            width: 18px;
        }
        QTabBar QToolButton {
            background-color: #E8E8E8;
            border: none;
        }
    )");

    m_layout->addWidget(m_tabs);

    // Bottom border line
    QFrame* bottomLine = new QFrame(this);
    bottomLine->setFrameShape(QFrame::HLine);
    bottomLine->setFrameShadow(QFrame::Plain);
    bottomLine->setFixedHeight(1);
    bottomLine->setStyleSheet("background-color: #BCBCBC; border: none;");
    m_layout->addWidget(bottomLine);
}

RibbonTab* RibbonWidget::addTab(const QString& label)
{
    RibbonTab* tab = new RibbonTab(m_tabs);
    // Horizontal scroll fallback: when groups overflow even in compact mode,
    // the row scrolls sideways rather than clipping off-screen. Vertical scroll
    // is disabled (the ribbon is a single fixed-height row). ScrollHost::wrap
    // re-parents `tab` into the returned scroll area.
    DVE::ScrollHost* host = DVE::ScrollHost::wrap(tab, Qt::Horizontal);
    m_tabs->addTab(host, label);
    return tab;
}

void RibbonWidget::setCurrentTab(int index)
{
    m_tabs->setCurrentIndex(index);
}

int RibbonWidget::currentTab() const
{
    return m_tabs->currentIndex();
}

void RibbonWidget::setCompactMode(bool compact)
{
    if (m_compactMode == compact) return;
    m_compactMode = compact;
    for (int i = 0; i < m_tabs->count(); ++i) {
        QWidget* page = m_tabs->widget(i);
        RibbonTab* tab = qobject_cast<RibbonTab*>(page);
        if (!tab) {
            if (auto* host = qobject_cast<DVE::ScrollHost*>(page))
                tab = qobject_cast<RibbonTab*>(host->widget());
        }
        if (tab)
            tab->setCompactMode(compact);
    }
    if (compact) {
        // Reuse AppTheme's 9pt tab font (single source of truth) rather than a
        // fresh QFont("Segoe UI", 9) literal -- matches ribbonMinimumHeight().
        // Compact group band = 36px icon button + 8px group vertical margin.
        const int tabBarH = AppTheme::lineUnit(AppTheme::fontDefault()) + 14;
        setMinimumHeight(tabBarH + 36 + 8 + 1);
    } else {
        setMinimumHeight(ribbonMinimumHeight());
    }
}

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

// ═══════════════════════════════════════════════════════════════════════════════
// RibbonGroup
// ═══════════════════════════════════════════════════════════════════════════════

RibbonGroup::RibbonGroup(const QString& title, QWidget* parent)
    : QWidget(parent)
    , m_hasSmallButtons(false)
{
    setObjectName("RibbonGroup");

    // The group has a fixed height so the ribbon stays at ~90px content height.
    setFixedHeight(90);
    setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Fixed);

    // ── Outer vertical layout ─────────────────────────────────────────────────
    //  Row 0 (stretch): content area  – large buttons  + optional small buttons
    //  Row 1 (fixed 1px): separator line
    //  Row 2 (fixed 16px): group title

    QVBoxLayout* outerVBox = new QVBoxLayout(this);
    outerVBox->setContentsMargins(4, 4, 4, 0);
    outerVBox->setSpacing(0);

    // ── Content area (large buttons + small buttons side-by-side) ─────────────
    QWidget* contentArea = new QWidget(this);
    contentArea->setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Expanding);

    QHBoxLayout* contentHBox = new QHBoxLayout(contentArea);
    contentHBox->setContentsMargins(0, 0, 0, 0);
    contentHBox->setSpacing(2);

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

    // ── Thin separator line above title ───────────────────────────────────────
    QFrame* sep = new QFrame(this);
    sep->setFrameShape(QFrame::HLine);
    sep->setFrameShadow(QFrame::Plain);
    sep->setFixedHeight(1);
    sep->setStyleSheet("background-color: #C8C8C8; border: none;");

    outerVBox->addWidget(sep);

    // ── Group title label ──────────────────────────────────────────────────────
    m_titleLabel = new QLabel(title, this);
    m_titleLabel->setAlignment(Qt::AlignHCenter | Qt::AlignVCenter);
    m_titleLabel->setFixedHeight(16);
    m_titleLabel->setStyleSheet(
        "font-family: 'Segoe UI'; font-size: 8pt; color: #555555;"
        "background-color: transparent; border: none;"
    );

    outerVBox->addWidget(m_titleLabel);

    // ── Right-side vertical border (thin separator between groups) ─────────────
    // Drawn via QSS on the widget itself.
    setStyleSheet(
        "RibbonGroup {"
        "  background-color: transparent;"
        "  border-right: 1px solid #C8C8C8;"
        "}"
    );
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

    // Fixed size: 72 wide × 70 tall
    btn->setFixedSize(72, 70);

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
    return btn;
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


// ═══════════════════════════════════════════════════════════════════════════════
// RibbonWidget
// ═══════════════════════════════════════════════════════════════════════════════

RibbonWidget::RibbonWidget(QWidget* parent)
    : QWidget(parent)
{
    setObjectName("RibbonWidget");
    setFixedHeight(100);
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
    m_tabs->addTab(tab, label);
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

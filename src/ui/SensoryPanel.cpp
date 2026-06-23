#include "SensoryPanel.h"

#include "database/LiveSync.h"
#include "utils/AppTheme.h"
#include "utils/OutputPaths.h"
#include "utils/ResponsiveLayout.h"

#include <QHBoxLayout>
#include <QFormLayout>
#include <QPushButton>
#include <QFileDialog>
#include <QMessageBox>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QDateTime>
#include <QFile>
#include <QTextStream>
#include <QStyle>
#include <QLabel>
#include <QScrollBar>
#include <QWheelEvent>
#include <QApplication>
#include <QScreen>
#include <QDir>
#include <QPainter>
#include <QPixmap>
#include <QImage>
#include <QBuffer>
#include <QFrame>
#include <QFileInfo>
#include <QHeaderView>
#include <QTreeWidget>
#include <QListWidget>
#include <QAction>
#include <QToolButton>
#include <QWidget>
#include <QVBoxLayout>
#include <QAbstractItemView>
#include <QFontMetrics>
#include <QDebug>
#include <functional>
#include <cmath>
#include <climits>
#include <utility>

namespace DVE {
namespace {

// v2.0.4 helper: attach a trailing ▼ action to a QLineEdit that opens
// a popup menu of preset values. The provider callback is invoked on
// every click so the menu always reflects the latest DB state — no
// caching, no NOTIFY plumbing needed. Empty / single-entry results
// show a disabled "(no saved values)" or single item; selecting any
// item populates the line edit's text.
// v2.4.9: the saved-values picker is a popup anchored BELOW the field with a
// scrollable list. Using Qt::Popup (not QMenu) means it never auto-flips upward
// or clips off-screen, and QListWidget gives a real scrollbar + mouse-wheel for
// a long backlog. optional onDelete adds a per-row red ✕ that removes the value
// from the pool (deleteSensoryHeaderPreset) and hides the row in place, so the
// user can prune several without reopening. Empty onDelete = pick-only rows.
void attachPresetDropdown(QLineEdit* edit,
                          std::function<QStringList()> provider,
                          std::function<void(const QString&)> onDelete)
{
    if (!edit) return;
    QAction* act = edit->addAction(
        edit->style()->standardIcon(QStyle::SP_ArrowDown),
        QLineEdit::TrailingPosition);
    act->setToolTip(QObject::tr("Pick from saved values"));
    QObject::connect(act, &QAction::triggered, edit,
        [edit, provider = std::move(provider), onDelete = std::move(onDelete)]() {
            const QStringList values = provider();

            auto* popup = new QFrame(edit, Qt::Popup);
            popup->setAttribute(Qt::WA_DeleteOnClose);
            popup->setFrameShape(QFrame::StyledPanel);
            auto* vlay = new QVBoxLayout(popup);
            vlay->setContentsMargins(0, 0, 0, 0);
            vlay->setSpacing(0);
            auto* list = new QListWidget(popup);
            list->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
            list->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
            list->setSelectionMode(QAbstractItemView::NoSelection);
            list->setFocusPolicy(Qt::NoFocus);
            vlay->addWidget(list);

            if (values.isEmpty()) {
                auto* it = new QListWidgetItem(
                    QObject::tr("(no saved values — type and Save Test Headers)"), list);
                it->setFlags(Qt::NoItemFlags);   // greyed, non-interactive
            } else {
                for (const QString& v : values) {
                    auto* item = new QListWidgetItem(list);
                    auto* row  = new QWidget(list);
                    auto* h    = new QHBoxLayout(row);
                    h->setContentsMargins(6, 1, 4, 1);
                    h->setSpacing(8);
                    auto* pick = new QToolButton(row);
                    pick->setText(v);
                    pick->setAutoRaise(true);
                    pick->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);
                    pick->setStyleSheet(QStringLiteral(
                        "QToolButton{border:none;text-align:left;padding:2px 4px;}"
                        "QToolButton:hover{background:palette(highlight);"
                        "color:palette(highlighted-text);}"));
                    h->addWidget(pick, 1);
                    QObject::connect(pick, &QToolButton::clicked, popup,
                        [edit, v, popup]() { edit->setText(v); popup->close(); });
                    if (onDelete) {
                        auto* del = new QToolButton(row);
                        del->setText(QStringLiteral("✕"));
                        del->setAutoRaise(true);
                        del->setToolTip(
                            QObject::tr("Remove “%1” from saved values").arg(v));
                        del->setStyleSheet(QStringLiteral(
                            "QToolButton{border:none;color:#b00020;font-weight:bold;"
                            "padding:2px 6px;}QToolButton:hover{color:#ff0000;}"));
                        h->addWidget(del, 0);
                        QObject::connect(del, &QToolButton::clicked, popup,
                            [onDelete, v, item]() {
                                onDelete(v);
                                item->setHidden(true);   // drop row in place; popup stays open
                            });
                    }
                    item->setSizeHint(row->sizeHint());
                    list->setItemWidget(item, row);
                }
            }

            // Anchor below the field; bound height to the room on-screen below
            // (cap ~16 rows) so it ALWAYS drops down and scrolls rather than
            // covering the screen or flipping upward. Width fits the longest value.
            const int rowH = qMax(24, edit->sizeHint().height());
            const QPoint below = edit->mapToGlobal(QPoint(0, edit->height()));
            const QRect avail  = edit->screen() ? edit->screen()->availableGeometry()
                                                : QRect(0, 0, 1920, 1080);
            const int roomBelow = qMax(rowH + 8, avail.bottom() - below.y() - 6);
            const int wanted    = qMin(qMax(1, list->count()), 16) * rowH + 6;
            const int popupH    = qMin(wanted, roomBelow);
            QFontMetrics fm(edit->font());
            int textW = 0;
            for (const QString& v : values) textW = qMax(textW, fm.horizontalAdvance(v));
            const int popupW = qBound(220, qMax(textW + 80, edit->width()), 540);
            popup->setFixedSize(popupW, popupH);
            popup->move(below);
            popup->show();
        });
}

} // anonymous
} // namespace DVE

#include "xlsxdocument.h"
#include "reporting/PptxWriter.h"
#include "reporting/SensoryReportSource.h"
#include "ReportPreviewDialog.h"
#include "ui/TesterRound.h"
#include "utils/ImageUtils.h"

namespace DVE {

namespace {
// QTableWidget's default sort uses QTableWidgetItem::operator< which compares
// the DisplayRole as a string — so "10" sorts before "9". This subclass
// stores the raw double in UserRole and compares against UserRole instead,
// giving numeric sort order while preserving the formatted display string.
class NumericTableItem : public QTableWidgetItem {
public:
    explicit NumericTableItem(double value)
        : QTableWidgetItem(QString::number(value, 'f', 1)) {
        setData(Qt::UserRole, value);
        setTextAlignment(Qt::AlignCenter);
    }
    bool operator<(const QTableWidgetItem& other) const override {
        return data(Qt::UserRole).toDouble()
             < other.data(Qt::UserRole).toDouble();
    }
};
} // anonymous namespace

// Ignores wheel events unless the spinbox has keyboard focus.
// Prevents accidental value changes when scrolling past sample cards.
class NoWheelDoubleSpinBox : public QDoubleSpinBox
{
public:
    explicit NoWheelDoubleSpinBox(QWidget* parent = nullptr) : QDoubleSpinBox(parent) {
        setFocusPolicy(Qt::StrongFocus);
    }
protected:
    void wheelEvent(QWheelEvent* e) override {
        e->ignore();  // never consume scroll — always let parent scroll area handle it
    }
};

// Ignores wheel events unless focused — prevents accidental Round changes
// when scrolling past the header row.
class NoWheelComboBox : public QComboBox
{
public:
    explicit NoWheelComboBox(QWidget* parent = nullptr) : QComboBox(parent) {
        setFocusPolicy(Qt::StrongFocus);
    }
protected:
    void wheelEvent(QWheelEvent* e) override {
        e->ignore();
    }
};

// ─────────────────────────────────────────────────────────────────────────────
// FlowLayout — cards wrap left-to-right, then down
// (Adapted from Qt's Flow Layout example.)
// ─────────────────────────────────────────────────────────────────────────────

FlowLayout::FlowLayout(QWidget* parent, int margin, int hSpacing, int vSpacing)
    : QLayout(parent), m_hSpace(hSpacing), m_vSpace(vSpacing)
{
    setContentsMargins(margin, margin, margin, margin);
}

FlowLayout::~FlowLayout()
{
    QLayoutItem* item;
    while ((item = takeAt(0)))
        delete item;
}

void FlowLayout::addItem(QLayoutItem* item) { m_items.append(item); }

int FlowLayout::horizontalSpacing() const
{
    return (m_hSpace >= 0) ? m_hSpace : 6;
}

int FlowLayout::verticalSpacing() const
{
    return (m_vSpace >= 0) ? m_vSpace : 6;
}

int FlowLayout::count() const { return m_items.size(); }

QLayoutItem* FlowLayout::itemAt(int index) const
{
    return (index >= 0 && index < m_items.size()) ? m_items.at(index) : nullptr;
}

QLayoutItem* FlowLayout::takeAt(int index)
{
    if (index >= 0 && index < m_items.size())
        return m_items.takeAt(index);
    return nullptr;
}

Qt::Orientations FlowLayout::expandingDirections() const { return {}; }

bool FlowLayout::hasHeightForWidth() const { return true; }

int FlowLayout::heightForWidth(int width) const
{
    return doLayout(QRect(0, 0, width, 0), true);
}

void FlowLayout::setGeometry(const QRect& rect)
{
    QLayout::setGeometry(rect);
    doLayout(rect, false);
}

QSize FlowLayout::sizeHint() const
{
    return minimumSize();
}

QSize FlowLayout::minimumSize() const
{
    QSize size;
    for (const QLayoutItem* item : m_items)
        size = size.expandedTo(item->minimumSize());
    int m = contentsMargins().left() + contentsMargins().right();
    return QSize(size.width() + m, size.height() + m);
}

int FlowLayout::doLayout(const QRect& rect, bool testOnly) const
{
    int left, top, right, bottom;
    getContentsMargins(&left, &top, &right, &bottom);
    QRect effectiveRect = rect.adjusted(left, top, -right, -bottom);
    int availW = effectiveRect.width();
    int spaceX = horizontalSpacing();
    int spaceY = verticalSpacing();

    QVector<QSize> sizeHints;
    sizeHints.reserve(m_items.size());
    for (int i = 0; i < m_items.size(); ++i)
        sizeHints.append(m_items[i]->sizeHint());

    struct Row { int firstIdx; int count; int totalWidth; int maxHeight; };
    QVector<Row> rows;
    {
        int rowWidth = 0;
        int rowStart = 0;
        int rowHeight = 0;
        for (int i = 0; i < m_items.size(); ++i) {
            const QSize& sz = sizeHints[i];
            int needed = (i == rowStart) ? sz.width() : rowWidth + spaceX + sz.width();
            if (needed > availW && i != rowStart) {
                rows.append({rowStart, i - rowStart, rowWidth, rowHeight});
                rowStart = i;
                rowWidth = sz.width();
                rowHeight = sz.height();
            } else {
                rowWidth = needed;
                rowHeight = qMax(rowHeight, sz.height());
            }
        }
        if (rowStart < m_items.size())
            rows.append({rowStart, static_cast<int>(m_items.size() - rowStart), rowWidth, rowHeight});
    }

    int y = effectiveRect.y();
    for (const Row& row : rows) {
        int xOffset = (availW - row.totalWidth) / 2;
        int x = effectiveRect.x() + xOffset;
        for (int i = row.firstIdx; i < row.firstIdx + row.count; ++i) {
            const QSize& sz = sizeHints[i];
            if (!testOnly)
                m_items[i]->setGeometry(QRect(QPoint(x, y), sz));
            x += sz.width() + spaceX;
        }
        y += row.maxHeight + spaceY;
    }

    return y - spaceY - rect.y() + bottom;
}

// ─────────────────────────────────────────────────────────────────────────────
// SampleCard
// ─────────────────────────────────────────────────────────────────────────────

SampleCard::SampleCard(int index, QWidget* parent)
    : QGroupBox(QString("Sample %1").arg(index + 1), parent)
{
    // v2.0.10: bumped from 220 → 245 so the V/R/HT row fits without the
    // heating-tech combo getting clipped by the card edge.
    setFixedWidth(245); // base width; updated by SensoryPanel widthChanged

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(3);
    mainLayout->setContentsMargins(6, 16, 6, 4);

    auto* nameRow = new QHBoxLayout;
    nameRow->setSpacing(3);
    nameRow->addWidget(new QLabel("Name:"));
    m_nameEdit = new QLineEdit;
    m_nameEdit->setPlaceholderText(QString("Sample %1").arg(index + 1));
    nameRow->addWidget(m_nameEdit);
    mainLayout->addLayout(nameRow);

    // ── Per-sample device properties ──
    auto* devGrid = new QGridLayout;
    devGrid->setSpacing(2);
    devGrid->setContentsMargins(0, 0, 0, 0);

    auto makeSmallLabel = [](const QString& text) {
        auto* lbl = new QLabel(text);
        lbl->setStyleSheet("font-size: 7pt; color: #555;");
        return lbl;
    };

    devGrid->addWidget(makeSmallLabel("V:"), 0, 0);
    m_voltageEdit = new QLineEdit;
    m_voltageEdit->setFixedWidth(52);
    m_voltageEdit->setFixedHeight(20);
    m_voltageEdit->setPlaceholderText("0.00");
    m_voltageEdit->setStyleSheet("font-size: 7pt;");
    devGrid->addWidget(m_voltageEdit, 0, 1);

    devGrid->addWidget(makeSmallLabel("R:"), 0, 2);
    m_resistanceEdit = new QLineEdit;
    m_resistanceEdit->setFixedWidth(52);
    m_resistanceEdit->setFixedHeight(20);
    m_resistanceEdit->setPlaceholderText("0.000");
    m_resistanceEdit->setStyleSheet("font-size: 7pt;");
    devGrid->addWidget(m_resistanceEdit, 0, 3);

    devGrid->addWidget(makeSmallLabel("HT:"), 0, 4);
    m_heatingTechCombo = new QComboBox;
    m_heatingTechCombo->setFixedWidth(72);
    m_heatingTechCombo->setFixedHeight(20);
    m_heatingTechCombo->setStyleSheet("font-size: 7pt;");
    m_heatingTechCombo->addItems({"", "EVO", "EVOMAX", "SE", "CCELL3.0", "T58G", "T51", "Competitor"});
    m_heatingTechCombo->setEditable(true);
    devGrid->addWidget(m_heatingTechCombo, 0, 5);

    m_powerLabel = new QLabel;
    m_powerLabel->setStyleSheet("font-size: 7pt; color: #333;");
    devGrid->addWidget(makeSmallLabel("P:"), 1, 0);
    devGrid->addWidget(m_powerLabel, 1, 1, 1, 5);

    // #7: power-type combo lives in the device grid below V/R/P.
    devGrid->addWidget(makeSmallLabel("PT:"), 2, 0);
    m_powerTypeCombo = new QComboBox;
    m_powerTypeCombo->setFixedHeight(20);
    m_powerTypeCombo->setStyleSheet("font-size: 7pt;");
    m_powerTypeCombo->addItems({
        tr("Constant Voltage"), tr("Constant Power"),
        tr("Variable Voltage"), tr("Variable Power")
    });
    devGrid->addWidget(m_powerTypeCombo, 2, 1, 1, 5);

    mainLayout->addLayout(devGrid);

    // Wire up power recalculation
    connect(m_voltageEdit, &QLineEdit::textChanged, this, [this]() { recalcPower(); emit changed(); });
    connect(m_resistanceEdit, &QLineEdit::textChanged, this, [this]() { recalcPower(); emit changed(); });
    connect(m_heatingTechCombo, &QComboBox::currentTextChanged, this, [this]() { recalcPower(); emit changed(); });
    // Voltage is only an input under Constant Voltage; the other three power
    // types either derive V or vary it across the run, so the field stops
    // making sense and we disable it. R and HT remain editable in all modes.
    connect(m_powerTypeCombo, &QComboBox::currentTextChanged,
            this, [this](const QString& pt) {
        m_voltageEdit->setEnabled(pt == tr("Constant Voltage"));
        emit changed();
    });

    // v2.0.1: per-widget commit events for LiveSync. Field paths must match
    // the canonical JSON serializer (SensoryData.cpp): metrics are flat keys
    // at the sample level, scalars use snake_case where the serializer does.
    connect(m_voltageEdit, &QLineEdit::editingFinished, this, [this]() {
        emit cellCommitted(QStringLiteral("voltage"),
                           m_voltageEdit->text().toDouble());
    });
    connect(m_resistanceEdit, &QLineEdit::editingFinished, this, [this]() {
        emit cellCommitted(QStringLiteral("resistance"),
                           m_resistanceEdit->text().toDouble());
    });
    connect(m_heatingTechCombo, &QComboBox::currentTextChanged, this,
            [this](const QString& s) {
        emit cellCommitted(QStringLiteral("heating_technology"), s);
    });
    connect(m_powerTypeCombo, &QComboBox::currentTextChanged, this,
            [this](const QString& s) {
        emit cellCommitted(QStringLiteral("power_type"), s);
    });

    // ── Sensory score spinboxes (decimal, 0.1 step) ──
    auto* formLayout = new QFormLayout;
    formLayout->setSpacing(4);
    formLayout->setContentsMargins(0, 2, 0, 2);
    formLayout->setLabelAlignment(Qt::AlignRight | Qt::AlignVCenter);
    formLayout->setFormAlignment(Qt::AlignHCenter | Qt::AlignTop);
    formLayout->setRowWrapPolicy(QFormLayout::DontWrapRows);
    formLayout->setFieldGrowthPolicy(QFormLayout::ExpandingFieldsGrow);
    formLayout->setHorizontalSpacing(8);
    formLayout->setVerticalSpacing(6);
    // Tooltips describe what each end of the 1-9 scale means. Hovering either
    // the label or the spinbox surfaces the same guidance.
    static const QMap<QString, QString> kMetricTooltips = {
        { "Burnt Taste",    QObject::tr("Rank 1-9. 9 means no burn, 1 means bad burn.") },
        { "Vapor Volume",   QObject::tr("Rank 1-9. 9 means big cloud, 1 means no cloud.") },
        { "Overall Flavor", QObject::tr("Rank 1-9. 9 means best flavor, 1 means worst flavor.") },
        { "Smoothness",     QObject::tr("Rank 1-9. 9 means extremely smooth, 1 means extremely harsh.") },
        { "Overall Liking", QObject::tr("Rank 1-9. 9 means it's the best, 1 means it's the worst.") },
    };
    for (const QString& metric : kSensoryMetrics) {
        auto* spin = new NoWheelDoubleSpinBox;
        spin->setRange(1.0, 9.0);
        spin->setSingleStep(0.1);
        spin->setDecimals(1);
        spin->setValue(5.0);
        spin->setFixedWidth(65);
        spin->setFixedHeight(22);
        spin->setButtonSymbols(QAbstractSpinBox::NoButtons);
        const QString tip = kMetricTooltips.value(metric);
        if (!tip.isEmpty()) spin->setToolTip(tip);
        m_spinBoxes[metric] = spin;
        formLayout->addRow(metric + ":", spin);
        if (!tip.isEmpty()) {
            if (auto* lbl = formLayout->labelForField(spin)) lbl->setToolTip(tip);
        }
        // Fix 72 px label column so metric names align across cards.
        if (auto* lbl = formLayout->labelForField(spin)) {
            lbl->setMinimumWidth(72);
            lbl->setMaximumWidth(72);
        }
        connect(spin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, &SampleCard::changed);
        // v2.0.1: per-metric LiveSync emission. The serializer stores metric
        // scores as flat keys on the sample object (sObj[metric] = value),
        // so the JSON path under the sample is just the metric name — no
        // "scores." prefix.
        connect(spin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, [this, metric](double v) {
                    emit cellCommitted(metric, v);
                });
    }
    mainLayout->addLayout(formLayout);

    // #7: puff length sits between scoring and comments per spec.
    auto* puffRow = new QHBoxLayout;
    puffRow->setSpacing(4);
    auto* puffLabel = new QLabel(tr("Puff length:"));
    puffLabel->setStyleSheet("font-size: 7pt; color: #555;");
    puffRow->addWidget(puffLabel);
    m_puffLengthSpin = new NoWheelDoubleSpinBox;
    m_puffLengthSpin->setRange(0.1, 60.0);
    m_puffLengthSpin->setSingleStep(0.5);
    m_puffLengthSpin->setDecimals(1);
    m_puffLengthSpin->setSuffix(QStringLiteral(" s"));
    m_puffLengthSpin->setValue(3.0);
    m_puffLengthSpin->setFixedWidth(72);
    m_puffLengthSpin->setFixedHeight(20);
    m_puffLengthSpin->setStyleSheet("font-size: 7pt;");
    m_puffLengthSpin->setButtonSymbols(QAbstractSpinBox::NoButtons);
    puffRow->addWidget(m_puffLengthSpin);
    puffRow->addStretch();
    mainLayout->addLayout(puffRow);
    connect(m_puffLengthSpin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
            this, [this](double) { emit changed(); });
    connect(m_puffLengthSpin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
            this, [this](double v) {
                emit cellCommitted(QStringLiteral("puff_length_sec"), v);
            });

    mainLayout->addWidget(new QLabel("Comments:"));
    // v2.1.0+: 4 px gap so the box doesn't visually touch the label.
    mainLayout->addSpacing(4);
    m_commentsEdit = new QTextEdit;
    m_commentsEdit->setMinimumHeight(36);
    m_commentsEdit->setMaximumHeight(50);
    // v2.1.0+: every variant of `border:` on the QTextEdit itself (type
    // selector, ID selector, no-selector inline, per-edge, QAbstractScrollArea
    // selector) was suppressed under our AppTheme stylesheet — only the
    // bottom + right edges painted, the rest disappeared. See
    // tests/outline_harness/. Drawing the border on a wrapping QFrame and
    // stripping the QTextEdit's own frame is the only approach that
    // reliably renders all four sides.
    m_commentsEdit->setFrameShape(QFrame::NoFrame);
    auto* commentsFrame = new QFrame;
    commentsFrame->setObjectName(QStringLiteral("sensoryCommentsFrame"));
    commentsFrame->setStyleSheet(
        "QFrame#sensoryCommentsFrame { border: 3px solid #A0A6AE; "
        "border-radius: 4px; background: white; }"
        "QFrame#sensoryCommentsFrame:focus { border-color: #0066CC; }");
    auto* commentsFrameLayout = new QVBoxLayout(commentsFrame);
    commentsFrameLayout->setContentsMargins(4, 4, 4, 4);
    commentsFrameLayout->addWidget(m_commentsEdit);
    mainLayout->addWidget(commentsFrame, 1);
    connect(m_commentsEdit, &QTextEdit::textChanged, this, &SampleCard::changed);
    connect(m_nameEdit, &QLineEdit::textChanged, this, &SampleCard::changed);

    // v2.0.1: editing-finished commits for the textual fields. Name commits
    // synchronously on editingFinished; comments need debouncing because
    // QTextEdit::textChanged fires on every keystroke and LiveSync does NOT
    // coalesce server-side — a 50-char comment would produce 50 separate
    // BEGIN/UPDATE/COMMIT transactions. A 500 ms single-shot timer collapses
    // a typing burst into a single commit.
    connect(m_nameEdit, &QLineEdit::editingFinished, this, [this]() {
        emit cellCommitted(QStringLiteral("name"), m_nameEdit->text());
    });
    m_commentsCommitTimer = new QTimer(this);
    m_commentsCommitTimer->setSingleShot(true);
    m_commentsCommitTimer->setInterval(500);
    connect(m_commentsCommitTimer, &QTimer::timeout, this, [this]() {
        emit cellCommitted(QStringLiteral("comments"),
                           m_commentsEdit->toPlainText());
    });
    connect(m_commentsEdit, &QTextEdit::textChanged, this, [this]() {
        m_commentsCommitTimer->start();
    });

    auto* removeBtn = new QPushButton("Remove");
    removeBtn->setFixedWidth(60);
    removeBtn->setFixedHeight(20);
    removeBtn->setStyleSheet("font-size: 7pt;");
    auto* btnRow = new QHBoxLayout;
    btnRow->addStretch();
    btnRow->addWidget(removeBtn);
    mainLayout->addLayout(btnRow);

    connect(removeBtn, &QPushButton::clicked, this, [this]() {
        emit removeRequested(this);
    });
}

void SampleCard::recalcPower()
{
    double v = m_voltageEdit->text().toDouble();
    double r = m_resistanceEdit->text().toDouble();
    QString tech = m_heatingTechCombo->currentText().trimmed().toUpper();

    double rOffset = 0.0;
    if (tech == "CCELL3.0" || tech == "CCELL 3.0" || tech == "T58G")
        rOffset = 0.78;
    else if (tech == "T51")
        rOffset = 0.25;

    double denom = r + rOffset;
    double power = (v > 0 && denom > 0) ? (v * v) / denom : 0.0;

    if (power > 0)
        m_powerLabel->setText(QString("%1W").arg(power, 0, 'f', 2));
    else
        m_powerLabel->setText("");
}

void SampleCard::attachNamePresetDropdown(std::function<QStringList()> provider,
                                          std::function<void(const QString&)> onDelete)
{
    attachPresetDropdown(m_nameEdit, std::move(provider), std::move(onDelete));
}

SensorySample SampleCard::toSample() const
{
    SensorySample s;
    s.name     = m_nameEdit->text().trimmed();
    s.comments = m_commentsEdit->toPlainText().trimmed();
    for (auto it = m_spinBoxes.constBegin(); it != m_spinBoxes.constEnd(); ++it) {
        s.scores[it.key()] = it.value()->value();
    }
    s.voltage           = m_voltageEdit->text().toDouble();
    s.resistance        = m_resistanceEdit->text().toDouble();
    s.heatingTechnology = m_heatingTechCombo->currentText().trimmed();
    s.powerType         = m_powerTypeCombo->currentText();
    s.puffLengthSec     = m_puffLengthSpin->value();

    // Compute power
    double rOffset = 0.0;
    QString tech = s.heatingTechnology.trimmed().toUpper();
    if (tech == "CCELL3.0" || tech == "CCELL 3.0" || tech == "T58G")
        rOffset = 0.78;
    else if (tech == "T51")
        rOffset = 0.25;
    double denom = s.resistance + rOffset;
    s.power = (s.voltage > 0 && denom > 0) ? (s.voltage * s.voltage) / denom : 0.0;

    return s;
}

void SampleCard::fromSample(const SensorySample& s)
{
    m_nameEdit->setText(s.name);
    m_commentsEdit->setPlainText(s.comments);
    for (auto it = m_spinBoxes.constBegin(); it != m_spinBoxes.constEnd(); ++it) {
        if (s.scores.contains(it.key())) {
            it.value()->blockSignals(true);
            it.value()->setValue(s.scores.value(it.key(), 5.0));
            it.value()->blockSignals(false);
        }
    }

    // Per-sample device properties
    m_voltageEdit->blockSignals(true);
    m_resistanceEdit->blockSignals(true);
    m_heatingTechCombo->blockSignals(true);

    m_voltageEdit->setText(s.voltage > 0 ? QString::number(s.voltage, 'f', 2) : QString());
    m_resistanceEdit->setText(s.resistance > 0 ? QString::number(s.resistance, 'f', 3) : QString());
    int htIdx = m_heatingTechCombo->findText(s.heatingTechnology, Qt::MatchFixedString);
    if (htIdx >= 0)
        m_heatingTechCombo->setCurrentIndex(htIdx);
    else
        m_heatingTechCombo->setCurrentText(s.heatingTechnology);

    m_voltageEdit->blockSignals(false);
    m_resistanceEdit->blockSignals(false);
    m_heatingTechCombo->blockSignals(false);

    // #7: per-sample test conditions. Default-aware: tolerate older JSON that
    // didn't carry these fields (deserializer returns "" / 0.0 in that case).
    m_powerTypeCombo->blockSignals(true);
    m_puffLengthSpin->blockSignals(true);
    int ptIdx = m_powerTypeCombo->findText(s.powerType);
    m_powerTypeCombo->setCurrentIndex(ptIdx >= 0 ? ptIdx : 0);
    m_puffLengthSpin->setValue(s.puffLengthSec > 0 ? s.puffLengthSec : 3.0);
    m_powerTypeCombo->blockSignals(false);
    m_puffLengthSpin->blockSignals(false);

    recalcPower();
}

// ─────────────────────────────────────────────────────────────────────────────
// SensoryPanel
// ─────────────────────────────────────────────────────────────────────────────

SensoryPanel::SensoryPanel(DatabaseManager* db, QWidget* parent)
    : QWidget(parent)
    , m_db(db)
{
    m_refreshTimer = new QTimer(this);
    m_refreshTimer->setSingleShot(true);
    m_refreshTimer->setInterval(150);
    connect(m_refreshTimer, &QTimer::timeout, this, &SensoryPanel::onRefreshChart);

    auto* outerLayout = new QVBoxLayout(this);
    outerLayout->setSpacing(4);
    outerLayout->setContentsMargins(4, 4, 4, 4);

    // ── Header row (test title, assessor, tester, media, date) ──
    auto* headerWidget = new QWidget;
    buildHeaderRow(headerWidget);
    outerLayout->addWidget(headerWidget);

    // ── Horizontal splitter: cards (left) + radar chart (right) ──
    m_splitter = new QSplitter(Qt::Horizontal);

    // Left: scroll area with flow layout for sample cards + Add button
    auto* cardsPanel = new QWidget;
    auto* cardsLayout = new QVBoxLayout(cardsPanel);
    cardsLayout->setContentsMargins(0, 0, 0, 0);
    cardsLayout->setSpacing(4);

    m_scrollArea = new QScrollArea;
    m_scrollArea->setWidgetResizable(true);
    m_scrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);

    m_flowContainer = new QWidget;
    m_flowLayout    = new FlowLayout(m_flowContainer, 4, 6, 6);
    m_scrollArea->setWidget(m_flowContainer);
    cardsLayout->addWidget(m_scrollArea, 1);

    m_addSampleBtn = new QPushButton("+ Add Sample");
    auto* addBtn = m_addSampleBtn;
    addBtn->setProperty("primary", true);
    addBtn->setIcon(AppTheme::icon("file-plus"));
    addBtn->setMinimumHeight(32);
    addBtn->style()->unpolish(addBtn);
    addBtn->style()->polish(addBtn);
    cardsLayout->addWidget(addBtn);
    connect(addBtn, &QPushButton::clicked, this, &SensoryPanel::onAddSample);

    // Averaged table overlay (shown when test avg selected, replaces cards).
    // Clicking a column header sorts by that column: alphabetical for the
    // Device column (plain QTableWidgetItem default), numeric for every
    // metric column (NumericTableItem overrides operator< — see below).
    m_avgOverlayTable = new QTableWidget;
    m_avgOverlayTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    m_avgOverlayTable->setAlternatingRowColors(true);
    m_avgOverlayTable->setSelectionMode(QAbstractItemView::NoSelection);
    m_avgOverlayTable->setSortingEnabled(true);
    m_avgOverlayTable->horizontalHeader()->setSectionsClickable(true);
    m_avgOverlayTable->verticalHeader()->setVisible(false);
    m_avgOverlayTable->verticalHeader()->setDefaultSectionSize(22);
    m_avgOverlayTable->setStyleSheet(
        "QTableWidget { font-size: 9pt; }"
        "QTableWidget::item { padding: 2px 4px; }"
        "QHeaderView::section { background: #1F4E79; color: white; font-weight: 600;"
        "  font-size: 8pt; padding: 2px 4px; border: none; }");

    // Stack: index 0 = cards panel, index 1 = averaged table
    m_leftStack = new QStackedWidget;
    m_leftStack->addWidget(cardsPanel);
    m_leftStack->addWidget(m_avgOverlayTable);
    m_leftStack->setCurrentIndex(0);

    m_splitter->addWidget(m_leftStack);

    // Right: radar chart fills remaining space
    m_chart = new RadarChartWidget;
    m_splitter->addWidget(m_chart);

    m_splitter->setStretchFactor(0, 50);   // cards: 50%
    m_splitter->setStretchFactor(1, 50);  // chart: 50%
    m_splitter->setSizes({10000, 10000}); // force equal initial split
    m_splitter->setChildrenCollapsible(false);

    outerLayout->addWidget(m_splitter, 1);

    // Start with one empty session
    SensorySession empty;
    empty.sessionName = QStringLiteral("New Session");
    empty.date = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    empty.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    m_sessions.append(empty);
    m_currentTesterIdx = 0;
    addSampleCard();

    // Responsive card width: 3-up (>=1100) -> 2-up (700-1099) -> 1-up (<700).
    connect(&DVE::ResponsiveLayout::instance(),
            &DVE::ResponsiveLayout::widthChanged,
            this, [this](int w) {
        int targetCardWidth = 245;  // 3-up at >=1100
        if (w < DVE::ResponsiveLayout::kSensoryNarrowThreshold)
            targetCardWidth = qMax(245, w - 60);  // 1-up
        else if (w < DVE::ResponsiveLayout::kCompactThreshold)
            targetCardWidth = 265;                // 2-up
        for (auto* card : m_cards) card->setFixedWidth(targetCardWidth);
        m_flowLayout->invalidate();
        m_flowLayout->activate();
    });
}

void SensoryPanel::buildHeaderRow(QWidget* container)
{
    auto* layout = new QHBoxLayout(container);
    layout->setContentsMargins(0, 0, 0, 0);
    layout->setSpacing(10);

    auto addField = [&](const QString& label, QLineEdit*& edit) {
        layout->addWidget(new QLabel(label));
        edit = new QLineEdit;
        edit->setFixedWidth(150);
        layout->addWidget(edit);
    };

    addField("Test Title:", m_testTitleEdit);
    addField("Assessor:",  m_assessorEdit);
    addField("Tester:",    m_testerEdit);

    // Bug 2: Round selector for double-blind R1/R2. Folded into testerName via
    // combineTesterRound() at buildSession() time; split back out in
    // applySession(). Default "1" (most tests are double-blind round 1 first).
    layout->addWidget(new QLabel("Round:"));
    m_roundCombo = new NoWheelComboBox;
    m_roundCombo->addItems({QStringLiteral("1"), QStringLiteral("2"), QStringLiteral("N/A")});
    m_roundCombo->setCurrentIndex(0);
    m_roundCombo->setFixedWidth(60);
    layout->addWidget(m_roundCombo);

    addField("Media:",     m_mediaEdit);

    // Plan C (C6 fix): the header fields previously had NO change wiring at
    // all — typing a Test Title / Assessor / Tester / Media or picking a Round
    // mutated the in-memory session (via buildSession() at snapshot time) but
    // never told MainWindow, so the crash snapshot missed header edits. Emit
    // dataEdited() on every header change so RecoveryManager::noteDirty() fires.
    // These do not touch the chart, so no scheduleChartRefresh here.
    for (QLineEdit* edit : {m_testTitleEdit, m_assessorEdit,
                            m_testerEdit, m_mediaEdit}) {
        connect(edit, &QLineEdit::textChanged, this, &SensoryPanel::dataEdited);
    }
    connect(m_roundCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &SensoryPanel::dataEdited);

    // v2.0.4: trailing ▼ dropdowns on the two fields that are most
    // commonly re-entered verbatim across sessions. The provider
    // lambdas fetch fresh from the DB on every click so values saved
    // on another workstation appear without a restart.
    attachPresetDropdown(m_testTitleEdit, [this]() -> QStringList {
        return m_db ? m_db->loadSensoryHeaderPresets("test_name") : QStringList{};
    }, [this](const QString& v) {
        if (m_db) m_db->deleteSensoryHeaderPreset("test_name", v);
    });
    attachPresetDropdown(m_mediaEdit, [this]() -> QStringList {
        return m_db ? m_db->loadSensoryHeaderPresets("media") : QStringList{};
    }, [this](const QString& v) {
        if (m_db) m_db->deleteSensoryHeaderPreset("media", v);
    });

    // v2.0.4: "Save Test Headers" button next to Media. Records the
    // current Test Title, Media, and per-sample names into the shared
    // preset pool so coworkers see them in their dropdowns next time.
    m_saveHeadersBtn = new QPushButton(tr("Save Test Headers"));
    auto* saveHeadersBtn = m_saveHeadersBtn;
    saveHeadersBtn->setProperty("primary", true);
    saveHeadersBtn->setMinimumHeight(28);
    saveHeadersBtn->style()->unpolish(saveHeadersBtn);
    saveHeadersBtn->style()->polish(saveHeadersBtn);
    saveHeadersBtn->setToolTip(
        tr("Add the current Test Title, Media, and sample names to the\n"
           "shared dropdown pool so coworkers can pick them instead of\n"
           "retyping. Safe to click repeatedly — duplicates are ignored."));
    connect(saveHeadersBtn, &QPushButton::clicked, this, [this]() {
        if (!m_db) return;
        QStringList sampleNames;
        sampleNames.reserve(m_cards.size());
        for (SampleCard* card : m_cards) {
            if (!card) continue;
            const QString n = card->toSample().name.trimmed();
            if (!n.isEmpty()) sampleNames << n;
        }
        const QString testTitle = m_testTitleEdit->text().trimmed();
        const QString media     = m_mediaEdit->text().trimmed();
        if (m_db->saveSensoryHeaderPresets(testTitle, media, sampleNames)) {
            // v2.0.5: surface what was saved so the user can confirm
            // sample names actually went into the pool. Duplicates are
            // silently skipped by the DB; this lists everything that
            // was attempted, not what's new.
            QStringList parts;
            if (!testTitle.isEmpty())
                parts << tr("Test Title: %1").arg(testTitle);
            if (!media.isEmpty())
                parts << tr("Media: %1").arg(media);
            if (!sampleNames.isEmpty())
                parts << tr("Sample names (%1): %2")
                             .arg(sampleNames.size())
                             .arg(sampleNames.join(QStringLiteral(", ")));
            const QString body = parts.isEmpty()
                ? tr("Nothing to save — all fields are empty.")
                : (tr("Saved to the shared dropdown pool:\n\n") +
                   parts.join(QStringLiteral("\n")) +
                   tr("\n\nDuplicates were skipped automatically. "
                      "Coworkers will see the new values in their "
                      "dropdowns next time they open a session."));
            QMessageBox::information(this, tr("Save Test Headers"), body);
        } else {
            QMessageBox::warning(this, tr("Save Test Headers"),
                tr("Could not save test headers to the database.\n%1")
                    .arg(m_db->lastError()));
        }
    });
    layout->addWidget(saveHeadersBtn);

    layout->addWidget(new QLabel("Date:"));
    m_dateLabel = new QLabel(QDateTime::currentDateTime().toString("yyyy-MM-dd"));
    layout->addWidget(m_dateLabel);

    auto* saveChartBtn = new QPushButton;
    saveChartBtn->setIcon(style()->standardIcon(QStyle::SP_DialogSaveButton));
    saveChartBtn->setToolTip("Save chart as image");
    saveChartBtn->setFixedSize(24, 24);
    saveChartBtn->setFlat(true);
    connect(saveChartBtn, &QPushButton::clicked, this, &SensoryPanel::onSaveChart);
    layout->addWidget(saveChartBtn);

    layout->addStretch();
}

// ─────────────────────────────────────────────────────────────────────────────
// Card management
// ─────────────────────────────────────────────────────────────────────────────

int SensoryPanel::activeSessionId() const
{
    if (m_currentTesterIdx < 0 || m_currentTesterIdx >= m_sessions.size())
        return -1;
    const int id = m_sessions[m_currentTesterIdx].id;
    return id > 0 ? id : -1;
}

void SensoryPanel::addSampleCard(const SensorySample& sample)
{
    int idx = m_cards.size();
    auto* card = new SampleCard(idx, m_flowContainer);
    if (!sample.name.isEmpty() || !sample.scores.isEmpty()) {
        card->fromSample(sample);
    }
    connect(card, &SampleCard::changed,          this, &SensoryPanel::scheduleChartRefresh);
    // Plan C (C6 fix): SampleCard::changed fires on every per-field value edit
    // (score, comment, name, V/R/HT/PT/puff). The card's value is folded into
    // m_sessions lazily by buildSession()/allSessions(), so the recovery
    // snapshot captures it — but only if MainWindow is told the snapshot is
    // dirty. scheduleChartRefresh just repaints the chart; dataEdited() is what
    // reaches RecoveryManager::noteDirty().
    connect(card, &SampleCard::changed,          this, &SensoryPanel::dataEdited);
    connect(card, &SampleCard::removeRequested,  this, &SensoryPanel::onRemoveCard);

    // v2.0.1: route per-cell commits through LiveSync. activeSessionId()
    // gates both the in-range check and the s.id > 0 placeholder check.
    //
    // v2.5.0 Task 3 (RC2): BEFORE (and regardless of) the LiveSync gate, record
    // the touched SCORE cell into the current session's dirtyCells so the
    // whole-session save's dirty-aware merge keeps this local edit. The LiveSync
    // stream is gated on activeSessionId()>0 and on a working sync connection;
    // when either is absent the merge would otherwise revert the edit to the DB
    // value (the RC2 data loss). The merge only arbitrates score keys, so only
    // kSensoryMetrics fields are marked. Programmatic loads (fromSample) set the
    // score widgets under blockSignals, so cellCommitted never fires for them.
    connect(card, &SampleCard::cellCommitted, this,
            [this, card](const QString& fieldPath, const QVariant& value) {
                const int idx = m_cards.indexOf(card);
                if (idx < 0) return;
                if (kSensoryMetrics.contains(fieldPath)
                    && m_currentTesterIdx >= 0
                    && m_currentTesterIdx < m_sessions.size()) {
                    m_sessions[m_currentTesterIdx].dirtyCells.insert(
                        QStringLiteral("samples[%1].%2").arg(idx).arg(fieldPath));
                }
                if (!m_liveSync) return;
                const int sessionId = activeSessionId();
                if (sessionId < 0) return;
                const QString jsonPath =
                    QStringLiteral("json_path:samples[%1].%2").arg(idx).arg(fieldPath);
                m_liveSync->commitCell(QStringLiteral("sensory_sessions"),
                                       sessionId, jsonPath, value);
            });

    m_flowLayout->addWidget(card);
    m_cards.append(card);

    // DATAVIEWER-2: scope the sample-name dropdown to the current Test
    // Title instead of the global pool (which flooded the screen). Uses
    // the SAME trimmed Test Title that save/backfill key the presets on,
    // so the scoped list lines up exactly; blank title -> empty.
    card->attachNamePresetDropdown([this]() -> QStringList {
        if (!m_db) return {};
        const QString test = m_testTitleEdit->text().trimmed();
        if (test.isEmpty()) return {};
        return m_db->loadSampleNamesForTest(test);
    }, [this](const QString& v) {
        // sample_name presets are scoped to the current Test Title (matches save).
        if (m_db) m_db->deleteSensoryHeaderPreset(
            "sample_name", v, m_testTitleEdit->text().trimmed());
    });

    scheduleChartRefresh();
}

void SensoryPanel::clearAllCards()
{
    for (SampleCard* card : m_cards) {
        m_flowLayout->removeWidget(card);
        card->deleteLater();
    }
    m_cards.clear();
}

// ─────────────────────────────────────────────────────────────────────────────
// Session serialisation
// ─────────────────────────────────────────────────────────────────────────────

SensorySession SensoryPanel::buildSession() const
{
    SensorySession sess;
    const QString testTitle = m_testTitleEdit->text().trimmed();
    QString existing;
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size())
        existing = m_sessions[m_currentTesterIdx].sessionName;

    // Sync sessionName to testTitle whenever the user has provided one.
    // Renaming a saved session changes its natural-key column; if the new
    // name collides with another existing row, the save loop in
    // MainWindow::onUpdateDatabase will surface a UniqueViolation and
    // prompt the user to override or cancel.
    //
    // Fallbacks (in order): keep an existing non-placeholder sessionName
    // if testTitle is blank, otherwise generate a timestamped unique name
    // so a "New Session" with no Test Title still gets a writable key.
    if (!testTitle.isEmpty()) {
        sess.sessionName = testTitle;
    } else if (!existing.isEmpty() && existing != QLatin1String("New Session")) {
        sess.sessionName = existing;
    } else {
        sess.sessionName = QString("Session_%1").arg(
            QDateTime::currentDateTime().toString("yyyyMMdd_HHmmss"));
    }
    sess.testTitle    = testTitle;
    sess.assessorName = m_assessorEdit->text().trimmed();
    sess.testerName   = combineTesterRound(m_testerEdit->text(), m_roundCombo->currentText());
    sess.media        = m_mediaEdit->text().trimmed();
    sess.date         = m_dateLabel->text();
    sess.timestamp    = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    sess.samples.reserve(m_cards.size());
    for (const SampleCard* card : m_cards) {
        sess.samples.append(card->toSample());
    }

    // Carry over fields not represented by SensoryPanel UI widgets
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()) {
        const SensorySession& stored = m_sessions[m_currentTesterIdx];

        // Preserve DB primary key + version. Without this, navigating between
        // sessions clobbers id back to -1, the next save attempts INSERT, and
        // the (session_name, tester_name, date) UNIQUE constraint trips
        // UniqueViolationDialog on every selection change. DetailedSensoryPanel
        // already does this implicitly via `sess = m_sessions[idx]`.
        sess.id      = stored.id;
        sess.version = stored.version;
        // v2.1.0+: carry over the loaded sessionName snapshot. Used by
        // MainWindow::onUpdateDatabase to detect Test Title renames and
        // route them to INSERT (new row) rather than UPDATE-in-place,
        // preserving the old row.
        sess.originalSessionName = stored.originalSessionName;

        // v2.5.0 Task 3 (RC2): carry the per-run dirty-cell set so the
        // whole-session save keeps locally-edited scores authoritative. The
        // per-cell handler writes into m_sessions[current].dirtyCells; without
        // this carry buildSession would emit a struct with an empty set and the
        // merge would revert the edits.
        sess.dirtyCells = stored.dirtyCells;

        sess.imagePaths         = stored.imagePaths;
        sess.imageLayouts       = stored.imageLayouts;
        sess.imageCrops         = stored.imageCrops;
        sess.sourceImagePath    = stored.sourceImagePath;

        // Session-level test properties (edited via MainWindow props panel)
        sess.control            = stored.control;
        sess.isBlind            = stored.isBlind;
        sess.primaryDifferences = stored.primaryDifferences;

        // Legacy fields (preserved for backward compat)
        sess.puffLength         = stored.puffLength;
        sess.burnStatus         = stored.burnStatus;
        sess.clogStatus         = stored.clogStatus;
        sess.leakStatus         = stored.leakStatus;
        sess.resistance         = stored.resistance;
        sess.voltage            = stored.voltage;
        sess.power              = stored.power;
        sess.heatingTechnology  = stored.heatingTechnology;
    }

    return sess;
}

void SensoryPanel::applySession(const SensorySession& session)
{
    clearAllCards();
    m_testTitleEdit->setText(session.testTitle);
    m_assessorEdit->setText(session.assessorName);
    const TesterRound tr = splitTesterRound(session.testerName);
    m_testerEdit->setText(tr.tester);
    if (m_roundCombo) m_roundCombo->setCurrentText(tr.round);
    m_mediaEdit->setText(session.media);
    if (!session.date.isEmpty()) m_dateLabel->setText(session.date);

    for (const SensorySample& s : session.samples) {
        addSampleCard(s);
    }

    onRefreshChart();
}

void SensoryPanel::saveCurrentTester()
{
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()) {
        m_sessions[m_currentTesterIdx] = buildSession();
    }
}

void SensoryPanel::inheritExistingIdsAndVersions()
{
    // One bulk SELECT instead of one per session — v2.0.5's per-key form
    // issued N round-trips on the UI thread and was invoked from the
    // 5-second auto-save tick, producing "Not Responding" freezes on
    // slower LANs. v2.0.6 moved this to load time (see loadSessions) and
    // replaced the per-key form with the bulk findSensorySessionsByKeys.
    //
    // Server-side lookup keeps the May-14 heap-corruption shape out of
    // the picture: we never construct a QVector<…> of QString-bearing
    // structs over the entire sensory_sessions table — only over the
    // small subset that actually matches the imported keys.
    if (!m_db) return;
    QVector<DatabaseManager::NaturalKey> keys;
    keys.reserve(m_sessions.size());
    for (const SensorySession& s : m_sessions) {
        if (s.id > 0) continue;
        DatabaseManager::NaturalKey k;
        k.sessionName = s.sessionName.trimmed();
        k.testerName  = s.testerName.trimmed();
        k.date        = s.date.trimmed();
        if (k.sessionName.isEmpty() && k.testerName.isEmpty() && k.date.isEmpty())
            continue;
        keys.append(k);
    }
    if (keys.isEmpty()) return;

    const auto matches = m_db->findSensorySessionsByKeys(keys);
    if (matches.isEmpty()) return;

    QHash<QString, DatabaseManager::SessionKey> byKey;
    byKey.reserve(matches.size());
    for (const auto& m : matches) {
        const QString k = m.sessionName + QChar('\x1f')
                        + m.testerName + QChar('\x1f') + m.date;
        byKey.insert(k, {m.id, m.version});
    }
    for (SensorySession& s : m_sessions) {
        if (s.id > 0) continue;
        const QString k = s.sessionName.trimmed() + QChar('\x1f')
                        + s.testerName.trimmed() + QChar('\x1f')
                        + s.date.trimmed();
        const auto it = byKey.constFind(k);
        if (it != byKey.constEnd()) {
            s.id      = static_cast<int>(it->id);
            s.version = it->version;
        }
    }
    // v2.1.0+: stamp originalSessionName on every persisted session so
    // subsequent Test Title edits can be detected as renames in the save
    // flow. We do this for ALL id>0 sessions (not just newly-resolved
    // ones) — fresh loads from DB go through loadSessions(), but we want
    // the same invariant after every reconciliation pass.
    for (SensorySession& s : m_sessions) {
        if (s.id > 0 && s.originalSessionName.isEmpty())
            s.originalSessionName = s.sessionName;
    }
}

void SensoryPanel::syncSavedSessionState(const QVector<SensorySession>& saved)
{
    const int n = qMin(saved.size(), m_sessions.size());
    for (int i = 0; i < n; ++i) {
        const SensorySession& src = saved[i];
        SensorySession&       dst = m_sessions[i];
        // Only adopt server-assigned identity when the save actually
        // landed (id > 0 after byRef back-fill).
        if (src.id > 0) {
            dst.id      = src.id;
            dst.version = src.version;
            // v2.5.0 RC4: the auto-suffix wrapper may have changed sessionName
            // AND testTitle (e.g. "T" -> "T_1") to resolve a collision. Adopt
            // BOTH into the panel copy so a subsequent buildSession() (which
            // regenerates sessionName FROM testTitle) cannot reproduce the
            // colliding name. Without this the panel keeps the unsuffixed title
            // and the next save re-collides forever (the June-10 loop).
            dst.sessionName = src.sessionName;
            dst.testTitle   = src.testTitle;
            // The session is now committed under the current name; future
            // renames are detected against this baseline.
            dst.originalSessionName = src.sessionName;
            // If this is the session currently shown in the header, refresh the
            // visible Test Title widget too (under blockSignals so the field's
            // editingFinished handler doesn't re-fire a spurious rename). The
            // navigator list refreshes via the sessionsChanged emit below.
            // v2.5.0 Task-5 review (5a): skip the rewrite when the field has
            // focus — a background save (5 s autosave / close-flush) must never
            // yank the title out from under the user's cursor mid-type. The
            // user's in-flight text is the freshest value; the adopt will catch
            // up on the next save once they've finished editing.
            if (i == m_currentTesterIdx && m_testTitleEdit
                && !m_testTitleEdit->hasFocus()
                && m_testTitleEdit->text() != src.testTitle) {
                QSignalBlocker block(m_testTitleEdit);
                m_testTitleEdit->setText(src.testTitle);
            }
        }
        // v2.5.0 Task 3 (RC2 review, CRITICAL 2): ADOPT the caller's dirty set
        // rather than unconditionally clearing it. A previously-persisted session
        // keeps id>0 even when THIS tick's tryWrite FAILED (the wrapper back-fills
        // id/version only on Success; on failure src is untouched and still
        // carries the dirty set). The save is synchronous on the UI thread, so no
        // edits interleave: the caller clears src.dirtyCells on its LOCAL copy
        // ONLY for sessions whose WriteResult == Success. Net (per the pure rule
        // in adoptedDirtyCellsAfterSave): id<=0 -> keep dst's set; id>0+Success ->
        // src cleared -> dst cleared; id>0+failure -> src keeps the set -> dst
        // keeps its protection so the retry still treats local edits as
        // authoritative. Extracted as a pure helper so it is unit-testable
        // without the GUI dependency tree.
        dst.dirtyCells =
            DVE::adoptedDirtyCellsAfterSave(src.id, src.dirtyCells, dst.dirtyCells);
        // Per-image identity back-fill: tryWriteSensorySession also writes
        // back imageIds + imageVersions for any new image rows it inserted.
        if (!src.imageIds.isEmpty())
            dst.imageIds = src.imageIds;
        if (!src.imageVersions.isEmpty())
            dst.imageVersions = src.imageVersions;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Session management (called by MainWindow)
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::selectSession(int index)
{
    if (index < 0 || index >= m_sessions.size()) return;
    if (index == m_currentTesterIdx) return;

    saveCurrentTester();
    m_currentTesterIdx = index;
    applySession(m_sessions[index]);
}

void SensoryPanel::showAveragedChart(const QVector<int>& sessionIndices)
{
    if (sessionIndices.isEmpty()) return;

    // Build a synthetic session with per-device averaged scores
    // across all selected sessions
    saveCurrentTester();

    struct DeviceAccum { QMap<QString, double> sums; int count = 0; };
    QMap<QString, DeviceAccum> deviceMap;
    QStringList deviceOrder;

    for (int si : sessionIndices) {
        if (si < 0 || si >= m_sessions.size()) continue;
        const SensorySession& sess = m_sessions[si];
        for (const SensorySample& s : sess.samples) {
            QString key = s.name.isEmpty() ? QStringLiteral("Sample") : s.name;
            if (!deviceMap.contains(key)) deviceOrder.append(key);
            DeviceAccum& acc = deviceMap[key];
            for (const QString& m : kSensoryMetrics)
                acc.sums[m] += s.scores.value(m, 5.0);
            acc.count++;
        }
    }

    SensorySession avgSess;
    avgSess.testTitle = "Averaged";
    for (const QString& devName : deviceOrder) {
        const DeviceAccum& acc = deviceMap[devName];
        SensorySample avgSample;
        avgSample.name = devName;
        for (const QString& m : kSensoryMetrics)
            avgSample.scores[m] = acc.sums.value(m, 0) / qMax(1, acc.count);
        avgSess.samples.append(avgSample);
    }

    m_chart->setSessions({avgSess});
}

void SensoryPanel::showAveragedTable(const QStringList& deviceNames,
                                      const QVector<QMap<QString, double>>& deviceAvgs)
{
    if (!m_leftStack || !m_avgOverlayTable) return;

    int nMetrics = kSensoryMetrics.size();
    m_avgOverlayTable->setColumnCount(1 + nMetrics);
    QStringList headers;
    headers << "Device";
    for (const QString& m : kSensoryMetrics)
        headers << m;
    m_avgOverlayTable->setHorizontalHeaderLabels(headers);
    m_avgOverlayTable->horizontalHeader()->setStretchLastSection(false);
    m_avgOverlayTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);

    // Distribute column widths: device gets 1.5x share, metrics share the rest
    int totalWidth = m_avgOverlayTable->viewport()->width();
    if (totalWidth < 100) totalWidth = 600;  // fallback before first layout
    double totalShares = 1.5 + nMetrics;
    int metricW = static_cast<int>(totalWidth / totalShares);
    int deviceW = static_cast<int>(metricW * 1.5);
    m_avgOverlayTable->setColumnWidth(0, deviceW);
    for (int c = 1; c <= nMetrics; ++c)
        m_avgOverlayTable->setColumnWidth(c, metricW);

    // Disable sorting while we mutate row contents so Qt doesn't reshuffle
    // mid-fill (when the user has a sort column active and rows get inserted
    // out of order). Re-enable afterwards so header clicks still sort.
    const bool wasSorting = m_avgOverlayTable->isSortingEnabled();
    m_avgOverlayTable->setSortingEnabled(false);

    m_avgOverlayTable->setRowCount(deviceNames.size());
    for (int i = 0; i < deviceNames.size(); ++i) {
        auto* nameItem = new QTableWidgetItem(deviceNames[i]);
        nameItem->setFlags(Qt::ItemIsEnabled);
        QFont f = nameItem->font(); f.setBold(true); nameItem->setFont(f);
        m_avgOverlayTable->setItem(i, 0, nameItem);

        const QMap<QString, double>& avgs = deviceAvgs[i];
        for (int c = 0; c < nMetrics; ++c) {
            double avg = avgs.value(kSensoryMetrics[c], 0.0);
            auto* valItem = new NumericTableItem(avg);
            valItem->setFlags(Qt::ItemIsEnabled);
            m_avgOverlayTable->setItem(i, 1 + c, valItem);
        }
    }

    m_avgOverlayTable->setSortingEnabled(wasSorting);
    m_leftStack->setCurrentIndex(1);
}

void SensoryPanel::showNormalView()
{
    if (m_leftStack)
        m_leftStack->setCurrentIndex(0);
}

void SensoryPanel::newSession()
{
    saveCurrentTester();

    SensorySession empty;
    empty.sessionName = QStringLiteral("New Session");
    empty.date = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    empty.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    m_sessions.append(empty);

    m_currentTesterIdx = m_sessions.size() - 1;

    clearAllCards();
    m_testTitleEdit->clear();
    m_assessorEdit->clear();
    m_testerEdit->clear();
    if (m_roundCombo) m_roundCombo->setCurrentIndex(0);   // default round "1"
    m_mediaEdit->clear();
    m_dateLabel->setText(empty.date);
    addSampleCard();

    emit sessionsChanged();
}

void SensoryPanel::closeSessions(const QVector<int>& indices)
{
    if (indices.isEmpty()) return;

    // Sort descending so removal doesn't shift later indices
    QVector<int> sorted = indices;
    std::sort(sorted.begin(), sorted.end(), std::greater<int>());

    for (int idx : sorted) {
        if (idx < 0 || idx >= m_sessions.size()) continue;
        m_sessions.removeAt(idx);
    }

    if (m_sessions.isEmpty()) {
        // No sessions remain — show empty panels
        m_currentTesterIdx = -1;
        clearAllCards();
        m_testTitleEdit->clear();
        m_assessorEdit->clear();
        m_testerEdit->clear();
        if (m_roundCombo) m_roundCombo->setCurrentIndex(0);   // default round "1"
        m_mediaEdit->clear();
        m_dateLabel->clear();
        m_chart->setSessions({});
    } else {
        // Clamp current index and display that session
        m_currentTesterIdx = qBound(0, m_currentTesterIdx, m_sessions.size() - 1);
        applySession(m_sessions[m_currentTesterIdx]);
    }

    emit sessionsChanged();
}

// v2.5.0 RC5 (close options): drop one session from the panel without the
// "do you really want to keep it open" machinery that closeSessions assumes.
// Mirrors closeSessions' single-index removal + index fixup, but for exactly
// one session — the caller (MainWindow's close-flow Discard option) has
// already removed the DB row. Disk autosave files are deliberately left alone.
void SensoryPanel::discardSession(int index)
{
    if (index < 0 || index >= m_sessions.size()) return;
    m_sessions.removeAt(index);

    // Keep m_currentTesterIdx pointing at a valid session (or -1 when empty).
    // A removal at or before the current index shifts it down by one; clamp
    // into range so applySession below never reads past the end.
    if (m_sessions.isEmpty()) {
        m_currentTesterIdx = -1;
        clearAllCards();
        m_testTitleEdit->clear();
        m_assessorEdit->clear();
        m_testerEdit->clear();
        if (m_roundCombo) m_roundCombo->setCurrentIndex(0);
        m_mediaEdit->clear();
        m_dateLabel->clear();
        m_chart->setSessions({});
    } else {
        if (index <= m_currentTesterIdx && m_currentTesterIdx > 0)
            --m_currentTesterIdx;
        m_currentTesterIdx = qBound(0, m_currentTesterIdx, m_sessions.size() - 1);
        applySession(m_sessions[m_currentTesterIdx]);
    }

    emit sessionsChanged();
}

// v2.5.0 RC5 (close options): bring `index` to the foreground and focus the
// Test Title field so the "Name It Now" close option drops the user's cursor
// exactly where the missing name goes.
void SensoryPanel::focusTitleForSession(int index)
{
    selectSession(index);   // no-op if already current; flushes the prior one
    if (m_testTitleEdit) {
        m_testTitleEdit->setFocus(Qt::OtherFocusReason);
        m_testTitleEdit->selectAll();
    }
}

void SensoryPanel::renameSession(int index, const QString& newLabel)
{
    if (index < 0 || index >= m_sessions.size()) return;
    // sessionLabel() formats as "testTitle - testerName". We update testTitle only.
    // Split at the FIRST " - " to avoid consuming testerName.
    // Known limitation: titles containing " - " will be split incorrectly.
    QString title = newLabel;
    QString tester = m_sessions[index].testerName;
    int sep = newLabel.indexOf(" - ");
    if (sep >= 0 && !tester.isEmpty()) {
        title = newLabel.left(sep).trimmed();
    }
    m_sessions[index].testTitle = title;

    // If renaming the current session, update the header field too
    if (index == m_currentTesterIdx)
        m_testTitleEdit->setText(title);

    emit sessionsChanged();
}

// Forward decl -- defined alongside loadFile below; loadSessions needs it too.
static bool isSameSensorySession(const SensorySession& a, const SensorySession& b);

void SensoryPanel::loadSessions(const QVector<SensorySession>& sessions)
{
    saveCurrentTester();

    // Remove the initial empty session if it's still in default state
    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    // v2.4.13: dedup against already-open sessions instead of blindly appending.
    // Loading a session from the Database Browser that is ALREADY open used to fork
    // a duplicate (then a "_1" split on save). A DB-loaded session carries a real
    // row id -- the most reliable identity; fall back to the natural key for id<=0.
    int lastIdx = -1;
    for (const SensorySession& s : sessions) {
        if (s.samples.isEmpty())
            continue;

        int existing = -1;
        for (int i = 0; i < m_sessions.size(); ++i)
            if ((s.id > 0 && m_sessions[i].id == s.id)
                || isSameSensorySession(m_sessions[i], s)) {
                existing = i;
                break;
            }
        if (existing >= 0) {          // already open -> switch to it, no duplicate
            lastIdx = existing;
            continue;
        }

        SensorySession copy = s;
        // v2.1.0+: capture sessionName at load time so a later Test Title edit can
        // be detected as a rename in the save flow.
        if (copy.originalSessionName.isEmpty() && copy.id > 0)
            copy.originalSessionName = copy.sessionName;
        m_sessions.append(copy);
        lastIdx = m_sessions.size() - 1;
    }

    if (lastIdx >= 0) {
        m_currentTesterIdx = lastIdx;
        applySession(m_sessions[m_currentTesterIdx]);
    }

    // v2.0.6: reconcile fresh-from-disk sessions (id <= 0) with the DB
    // exactly once here, where new sessions enter m_sessions. The
    // previous home for this call was MainWindow::onUpdateDatabase, but
    // that path runs every Ctrl+U and every 5-second auto-save tick —
    // re-doing the lookup forever after the first save was both
    // pointless (subsequent calls were no-ops once id > 0) and the
    // source of the v2.0.5 "Not Responding" freeze.
    inheritExistingIdsAndVersions();

    emit sessionsChanged();
}

QVector<SensorySession> SensoryPanel::allSessions()
{
    saveCurrentTester();
    return m_sessions;
}

SensorySession* SensoryPanel::currentSession()
{
    saveCurrentTester();
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size())
        return &m_sessions[m_currentTesterIdx];
    return nullptr;
}

// v2.4.2 SP2-T4 (reset-to-5 keystone). Overlay-only twin of the blessed
// dbAuthoritativeSessions() overlay (see that function's comment). MainWindow
// computes `mergedSession` via mergeSensoryPreservingDbScores() and hands it
// here; we copy ONLY the kSensoryMetrics scalar scores onto the current
// in-memory session — never round-tripping through sensorySessionFromJson,
// which would drop images and reset id/version. Then we re-render the cards
// from the updated struct via applySession() so a subsequent buildSession()
// (driven by the navigator refresh) reads the merged values instead of the
// stale on-screen widgets and reverts nothing.
void SensoryPanel::applyMergedScoresToCurrentSession(const QJsonObject& mergedSession)
{
    if (m_currentTesterIdx < 0 || m_currentTesterIdx >= m_sessions.size())
        return;
    SensorySession& sess = m_sessions[m_currentTesterIdx];
    // SP2-T4 hardening: overlay the scalar scores via the shared pipeline
    // helper so this production path and the e2e regression test exercise the
    // SAME loop (kills the test/prod drift that let 9be0550 pass while broken).
    overlayMergedScores(sess, mergedSession);
    // Re-render the visible cards from the merged struct. applySession()
    // rebuilds every SampleCard (selectSession() would early-return without
    // repainting). m_currentSampleIdx / m_currentTesterIdx are unchanged.
    applySession(sess);
}

QString SensoryPanel::sessionLabel(const SensorySession& s) const
{
    QString title  = s.testTitle.isEmpty() ? QString() : s.testTitle;
    QString tester = s.testerName.isEmpty() ? s.assessorName : s.testerName;

    if (!title.isEmpty() && !tester.isEmpty())
        return title + " - " + tester;
    if (!title.isEmpty())
        return title;
    if (!tester.isEmpty())
        return tester;
    return s.sessionName.isEmpty() ? QStringLiteral("(unnamed)") : s.sessionName;
}

// ─────────────────────────────────────────────────────────────────────────────
// Chart
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::scheduleChartRefresh()
{
    m_refreshTimer->start();
}

void SensoryPanel::onRefreshChart()
{
    SensorySession sess = buildSession();
    m_chart->setSessions({sess});
}

void SensoryPanel::onAddSample()
{
    addSampleCard();
}

void SensoryPanel::onRemoveCard(SampleCard* card)
{
    // v2.5.0 Task 3 (RC2 review, CRITICAL 1): the card's position in m_cards is
    // the sample index buildSession() will emit, and dirtyCells paths embed that
    // index. Remap the CURRENT session's dirty set across this removal BEFORE the
    // card leaves m_cards, or later samples' dirty paths would point at the wrong
    // (now-shifted) rows and the merge would revert their edits.
    const int removedIdx = m_cards.indexOf(card);
    if (removedIdx >= 0
        && m_currentTesterIdx >= 0
        && m_currentTesterIdx < m_sessions.size()) {
        SensorySession& sess = m_sessions[m_currentTesterIdx];
        sess.dirtyCells =
            DVE::remapDirtyCellsAfterSampleRemoval(sess.dirtyCells, removedIdx);
    }

    m_flowLayout->removeWidget(card);
    m_cards.removeOne(card);
    card->deleteLater();
    scheduleChartRefresh();
}

void SensoryPanel::onSaveChart()
{
    if (!m_chart) return;

    QString path = QFileDialog::getSaveFileName(
        this, "Save Chart Image", OutputPaths::resolveDir(ReportMode::Sensory,m_lastBrowseDir),
        "PNG Image (*.png);;JPEG Image (*.jpg);;BMP Image (*.bmp)");
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    QPixmap pixmap(m_chart->size());
    pixmap.fill(Qt::white);
    m_chart->render(&pixmap);

    if (!pixmap.save(path)) {
        QMessageBox::warning(this, "Save Chart",
            "Failed to save chart image to:\n" + path);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Last browse directory
// ─────────────────────────────────────────────────────────────────────────────

QString SensoryPanel::lastBrowseDir() const
{
    if (!m_lastBrowseDir.isEmpty() && QDir(m_lastBrowseDir).exists())
        return m_lastBrowseDir;

    return OutputPaths::documentsDir();
}

void SensoryPanel::setLastBrowseDir(const QString& filePath)
{
    m_lastBrowseDir = QFileInfo(filePath).absolutePath();
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

bool SensoryPanel::isDefaultState() const
{
    if (m_cards.size() != 1) return false;
    if (!m_testTitleEdit->text().trimmed().isEmpty()) return false;
    if (!m_assessorEdit->text().trimmed().isEmpty()) return false;
    if (!m_testerEdit->text().trimmed().isEmpty()) return false;
    if (!m_mediaEdit->text().trimmed().isEmpty()) return false;

    SensorySample s = m_cards.first()->toSample();
    if (!s.name.isEmpty()) return false;
    if (!s.comments.isEmpty()) return false;
    for (auto it = s.scores.constBegin(); it != s.scores.constEnd(); ++it) {
        if (qAbs(it.value() - 5.0) > 0.01) return false;
    }
    return true;
}

// ─────────────────────────────────────────────────────────────────────────────
// Save
// ─────────────────────────────────────────────────────────────────────────────

// DATAVIEWER-8: the currently-displayed session is savable iff it has both a
// non-empty test name and a non-empty tester (whitespace-only counts as empty).
// Mirrors the save() hard-guard so callers can decide BEFORE invoking save()
// whether it would surface a "name required" modal.
bool SensoryPanel::currentSessionSavable() const
{
    return !m_testTitleEdit->text().trimmed().isEmpty()
        && !m_testerEdit->text().trimmed().isEmpty();
}

void SensoryPanel::save()
{
    if (m_cards.isEmpty()) {
        QMessageBox::warning(this, "No Data", "Add at least one sample before saving.");
        return;
    }

    // DATAVIEWER-8: a test name + tester are required so the session has a
    // reliable natural key and a deterministic on-disk filename. This hard
    // guard replaces the old "auto-generate a default test name" soft-prompt.
    if (m_testTitleEdit->text().trimmed().isEmpty()
        || m_testerEdit->text().trimmed().isEmpty()) {
        QMessageBox::warning(this, tr("Test Name and Tester Required"),
            tr("Enter a test name and a tester before saving - both are needed "
               "to store and find this session reliably."));
        if (m_testTitleEdit->text().trimmed().isEmpty()) m_testTitleEdit->setFocus();
        else m_testerEdit->setFocus();
        return;
    }

    // Flush current UI state into m_sessions
    saveCurrentTester();

    // DATAVIEWER-8: silent, auto-named save — no Save-As dialog. The on-disk
    // base path is derived from the session label (title - tester[ - round]),
    // sanitised for the filesystem. m_savePath is a BASE path (no extension);
    // saveToJson/saveToExcel append ".json"/".xlsx". m_lastBrowseDir is only a
    // resolver hint here and is intentionally NOT overwritten with the auto dir.
    const SensorySession cur = buildSession();
    m_savePath = OutputPaths::autoSavePath(ReportMode::Sensory, sessionLabel(cur),
                                           m_lastBrowseDir, QString());

    // Save current session's JSON/Excel files. DATAVIEWER-4: the exported
    // scores must be DB-authoritative (LiveSync may hold newer per-cell values
    // than this struct). One helper call flushes once and serves both writers.
    SensorySession sess = dbAuthoritativeSessions({cur}).value(0, cur);
    saveToJson(m_savePath + ".json", sess);
    saveToExcel(m_savePath + ".xlsx", sess);

    // Save ALL sessions to the database (not just the current one).
    // v2.0.1: LiveSync owns DB persistence for per-cell edits; this loop
    // remains as a fallback so manual Save still flushes any session that
    // wasn't yet committed (e.g., fresh imports). m_db->saveSensorySession
    // returns bool — no dialogs are surfaced for conflicts since LiveSync
    // resolves on a per-cell basis.
    int dbSaved = 0;
    if (m_db) {
        for (const SensorySession& s : m_sessions) {
            if (s.samples.isEmpty()) continue;
            // DATAVIEWER-8: never push a half-filled non-current session to the
            // DB with a junk ('','',date) natural key. The current session is
            // already gated by the hard guard at the top of save().
            if (!DVE::isSensorySessionSavable(s)) continue;
            if (m_db->saveSensorySession(s))
                ++dbSaved;
        }
    }

    emit sessionsChanged();

    {
        QMessageBox msg(QMessageBox::Information, "Saved",
            QString("Session saved to:\n%1.json\n%2.xlsx\nDatabase (%3 session%4)")
                .arg(QFileInfo(m_savePath).fileName(),
                     QFileInfo(m_savePath).fileName(),
                     QString::number(dbSaved),
                     dbSaved == 1 ? "" : "s"),
            QMessageBox::Ok, this);
        // Remove the extra spacing Qt adds between icon and text
        if (auto* grid = qobject_cast<QGridLayout*>(msg.layout())) {
            grid->setHorizontalSpacing(8);
            grid->setContentsMargins(10, 10, 10, 10);
        }
        msg.exec();
    }
}

void SensoryPanel::saveToJson(const QString& path, const SensorySession& sess)
{
    // Routes through the canonical pipeline-layer encoder so the on-disk
    // .json wire format stays byte-identical to the Postgres JSONB blob
    // and the offline-snapshot copy. Any new SensorySession field added
    // there is automatically included here.
    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Text)) return;
    f.write(QJsonDocument(sensorySessionToJson(sess)).toJson(QJsonDocument::Indented));
}

void SensoryPanel::saveToExcel(const QString& path, const SensorySession& sess)
{
    QXlsx::Document xlsx;

    // Write header row once with bold formatting (avoids read-back cycle).
    QXlsx::Format hdrFmt;
    hdrFmt.setFontBold(true);
    xlsx.write(1, 1, QStringLiteral("Sample"), hdrFmt);
    for (int i = 0; i < kSensoryMetrics.size(); ++i)
        xlsx.write(1, i + 2, kSensoryMetrics[i], hdrFmt);
    xlsx.write(1, kSensoryMetrics.size() + 2, QStringLiteral("Comments"), hdrFmt);

    int row = 2;
    for (const SensorySample& s : sess.samples) {
        xlsx.write(row, 1, s.name.isEmpty() ? QString("Sample %1").arg(row - 1) : s.name);
        for (int i = 0; i < kSensoryMetrics.size(); ++i)
            xlsx.write(row, i + 2, s.scores.value(kSensoryMetrics[i], 5.0));
        xlsx.write(row, kSensoryMetrics.size() + 2, s.comments);
        ++row;
    }

    row += 2;
    xlsx.write(row, 1, "Test Title");  xlsx.write(row, 2, sess.testTitle);   ++row;
    xlsx.write(row, 1, "Tester");      xlsx.write(row, 2, sess.testerName);  ++row;
    xlsx.write(row, 1, "Assessor");    xlsx.write(row, 2, sess.assessorName);++row;
    xlsx.write(row, 1, "Media");       xlsx.write(row, 2, sess.media);       ++row;
    xlsx.write(row, 1, "Date");        xlsx.write(row, 2, sess.date);        ++row;
    xlsx.write(row, 1, "Control");     xlsx.write(row, 2, sess.control);     ++row;
    xlsx.write(row, 1, "Blind?");      xlsx.write(row, 2, sess.isBlind ? "Y" : "N"); ++row;
    xlsx.write(row, 1, "Primary Difference(s)"); xlsx.write(row, 2, sess.primaryDifferences); ++row;

    xlsx.saveAs(path);
}

// ─────────────────────────────────────────────────────────────────────────────
// Load Excel
// ─────────────────────────────────────────────────────────────────────────────

// v2.4.13: the reliable "same file" identity is the source path (set on every
// loaded session in loadFile). Windows paths are case-insensitive and may differ
// by separator, so normalize before comparing. Empty never matches.
static bool isSameSourcePath(const QString& a, const QString& b)
{
    if (a.isEmpty() || b.isEmpty()) return false;
    return QFileInfo(a).absoluteFilePath()
               .compare(QFileInfo(b).absoluteFilePath(), Qt::CaseInsensitive) == 0;
}

// Fallback identity for sessions that carry NO source path (e.g. recovered from a
// crash snapshot): the DB's natural key. Keyed on the EFFECTIVE title (testTitle,
// falling back to sessionName -- the Excel "saved format" loader sets testTitle but
// leaves sessionName empty until the first save, which is why an earlier
// sessionName-only compare missed) + tester + date, trimmed + case-insensitive.
static bool isSameSensorySession(const SensorySession& a, const SensorySession& b)
{
    auto title = [](const SensorySession& s) {
        return (s.testTitle.isEmpty() ? s.sessionName : s.testTitle).trimmed();
    };
    return !title(a).isEmpty()
        && title(a).compare(title(b), Qt::CaseInsensitive) == 0
        && a.testerName.trimmed().compare(b.testerName.trimmed(), Qt::CaseInsensitive) == 0
        && a.date.trimmed() == b.date.trimmed();
}

void SensoryPanel::loadFile(const QString& path)
{
    saveCurrentTester();

    // v2.4.13: if a session loaded from THIS exact file is already open, just switch
    // to it -- never reload (which forks a duplicate that later trips
    // idx_sensory_sessions_key and saves as a "_1" split). sourceFilePath is set on
    // every loaded session below; it's the reliable identity for a re-open.
    for (int i = 0; i < m_sessions.size(); ++i) {
        if (isSameSourcePath(m_sessions[i].sourceFilePath, path)) {
            m_currentTesterIdx = i;
            applySession(m_sessions[i]);
            emit sessionsChanged();
            return;
        }
    }

    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    const int firstNew = m_sessions.size();   // sessions already open before this load
    int loaded = 0;
    QString ext = QFileInfo(path).suffix().toLower();

    if (ext == "json") {
        QFile f(path);
        if (!f.open(QIODevice::ReadOnly | QIODevice::Text)) return;
        QJsonDocument doc = QJsonDocument::fromJson(f.readAll());
        if (doc.isNull() || !doc.isObject()) return;

        // Route through the canonical decoder so this user-facing import
        // accepts the same wire format as the DB JSONB column and the
        // offline-snapshot copy. Backward-compat defaults for new keys
        // (power_type / puff_length_sec) are handled there.
        SensorySession sess = sensorySessionFromJson(doc.object());

        if (sess.sessionName.isEmpty())
            sess.sessionName = QFileInfo(path).baseName();

        if (!sess.samples.isEmpty()) {
            m_sessions.append(sess);
            ++loaded;
        }
    } else {
        QXlsx::Document xlsx(path);
        if (!xlsx.load()) return;

        QString testTitle = QFileInfo(path).baseName();
        QStringList excelMetrics;
        for (int col = 2; col <= 6; ++col) {
            QString hdr = xlsx.read(1, col).toString().trimmed();
            if (!hdr.isEmpty()) excelMetrics << hdr;
        }
        if (excelMetrics.size() != 5)
            excelMetrics = QStringList(kSensoryMetrics.begin(), kSensoryMetrics.end());

        QString a1 = xlsx.read(1, 1).toString().trimmed();
        bool isSavedFormat = (a1.toLower() == "sample");

        if (isSavedFormat)
            loaded += loadExcelSavedFormat(xlsx, excelMetrics, testTitle, path);
        else
            loaded += loadExcelStandardFormat(xlsx, excelMetrics, testTitle);
    }

    if (loaded == 0) return;

    // v2.4.13: tag freshly-loaded sessions with their source file so a future
    // re-open of this file is recognized by the switch at the top of loadFile.
    for (int i = firstNew; i < m_sessions.size(); ++i)
        m_sessions[i].sourceFilePath = path;

    // Fallback dedup for already-open sessions that carry NO source path (recovered
    // from a crash snapshot): drop a just-loaded session matching one of them on the
    // DB natural key and switch to it. (Re-opens of a file already open are handled
    // by the source-path switch above.)
    int switchTo = -1;
    for (int n = m_sessions.size() - 1; n >= firstNew; --n) {
        bool removed = false;
        for (int e = 0; e < firstNew; ++e)
            if (isSameSensorySession(m_sessions[e], m_sessions[n])) {
                switchTo = e;
                m_sessions.remove(n);
                removed = true;
                break;
            }
        if (!removed) {
            // genuinely new -> log keys so any residual duplicate is traceable.
            qInfo().noquote() << "[sensory-load] NEW session kept:"
                << "title=" << m_sessions[n].testTitle << "name=" << m_sessions[n].sessionName
                << "tester=" << m_sessions[n].testerName << "date=" << m_sessions[n].date
                << "src=" << m_sessions[n].sourceFilePath;
            for (int e = 0; e < firstNew; ++e)
                qInfo().noquote() << "   vs existing:"
                    << "title=" << m_sessions[e].testTitle << "name=" << m_sessions[e].sessionName
                    << "tester=" << m_sessions[e].testerName << "date=" << m_sessions[e].date
                    << "src=" << m_sessions[e].sourceFilePath;
        }
    }
    if (m_sessions.size() == firstNew) {
        if (switchTo >= 0) {
            m_currentTesterIdx = switchTo;
            applySession(m_sessions[switchTo]);
            emit sessionsChanged();
        }
        return;
    }

    // v2.0.10: re-import paths used to skip this and INSERT blindly, tripping
    // idx_sensory_sessions_key when a row with the same natural key already
    // existed. The v2.0.6 author wired inherit into loadSessions() only —
    // Excel-import is the other entry point and needs it too. Bulk SELECT,
    // idempotent if every session already has an id.
    inheritExistingIdsAndVersions();

    if (m_db) {
        for (const SensorySession& s : m_sessions) {
            if (!s.samples.isEmpty())
                m_db->saveSensorySession(s);
        }
    }

    m_currentTesterIdx = m_sessions.size() - 1;
    applySession(m_sessions[m_currentTesterIdx]);
    emit sessionsChanged();
}

void SensoryPanel::loadFiles()
{
    QStringList paths = QFileDialog::getOpenFileNames(
        this, "Load Sensory Files", OutputPaths::resolveDir(ReportMode::Sensory, lastBrowseDir()),
        "Sensory files (*.xlsx *.json);;Excel files (*.xlsx);;JSON files (*.json);;All files (*)");
    if (paths.isEmpty()) return;
    setLastBrowseDir(paths.first());

    saveCurrentTester();

    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    int loaded = 0;
    for (const QString& path : paths) {
        QString ext = QFileInfo(path).suffix().toLower();

        if (ext == "json") {
            // ── Load JSON file ──────────────────────────────────────────
            QFile f(path);
            if (!f.open(QIODevice::ReadOnly | QIODevice::Text)) {
                QMessageBox::warning(this, "Load Error",
                                     "Could not open JSON file:\n" + path);
                continue;
            }
            QJsonDocument doc = QJsonDocument::fromJson(f.readAll());
            if (doc.isNull() || !doc.isObject()) {
                QMessageBox::warning(this, "Parse Error",
                                     "Invalid JSON format:\n" + path);
                continue;
            }

            // Route through the canonical decoder so this user-facing import
            // accepts the same wire format as the DB JSONB column and the
            // offline-snapshot copy.
            SensorySession sess = sensorySessionFromJson(doc.object());

            if (sess.sessionName.isEmpty())
                sess.sessionName = QFileInfo(path).baseName();

            if (!sess.samples.isEmpty()) {
                m_sessions.append(sess);
                ++loaded;
            }
        } else {
            // ── Load Excel file ─────────────────────────────────────────
            QXlsx::Document xlsx(path);
            if (!xlsx.load()) {
                QMessageBox::warning(this, "Load Error",
                                     "Could not open Excel file:\n" + path);
                continue;
            }

            QString testTitle = QFileInfo(path).baseName();
            QString dateNow   = QDateTime::currentDateTime().toString("yyyy-MM-dd");
            QString tsNow     = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

            // Read metric headers from columns B-F (row 1)
            QStringList excelMetrics;
            for (int col = 2; col <= 6; ++col) {
                QString hdr = xlsx.read(1, col).toString().trimmed();
                if (!hdr.isEmpty()) excelMetrics << hdr;
            }
            if (excelMetrics.size() != 5)
                excelMetrics = QStringList(kSensoryMetrics.begin(), kSensoryMetrics.end());

            // ── Format detection ────────────────────────────────────────
            // Saved format: A1 == "Sample", flat table, metadata below
            // Standardized template: A1 == device name, multi-tester blocks
            QString a1 = xlsx.read(1, 1).toString().trimmed();
            bool isSavedFormat = (a1.toLower() == "sample");

            if (isSavedFormat) {
                // ── Parse saved format (single-tester, flat table) ──────
                loaded += loadExcelSavedFormat(xlsx, excelMetrics, testTitle, path);
            } else {
                // ── Parse standardized template (multi-tester blocks) ───
                loaded += loadExcelStandardFormat(xlsx, excelMetrics, testTitle);
            }
        }
    }

    if (loaded == 0) {
        QMessageBox::warning(this, "No Data",
                             "No sample data found in the selected file(s).");
        return;
    }

    inheritExistingIdsAndVersions();

    // Immediately save all loaded sessions to the database
    if (m_db) {
        for (const SensorySession& s : m_sessions) {
            if (!s.samples.isEmpty())
                m_db->saveSensorySession(s);
        }
    }

    m_currentTesterIdx = m_sessions.size() - 1;
    applySession(m_sessions[m_currentTesterIdx]);
    emit sessionsChanged();
}

// ─────────────────────────────────────────────────────────────────────────────
// Load Excel — saved format (single tester, flat "Sample | Metrics | Comments")
// ─────────────────────────────────────────────────────────────────────────────

int SensoryPanel::loadExcelSavedFormat(QXlsx::Document& xlsx,
                                        const QStringList& excelMetrics,
                                        const QString& testTitle,
                                        const QString& filePath)
{
    QString dateNow = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    QString tsNow   = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    int nMetrics    = excelMetrics.size();

    // Read sample rows (starting at row 2, until empty row)
    QVector<SensorySample> samples;
    for (int row = 2; row <= 100; ++row) {
        QString sampleName = xlsx.read(row, 1).toString().trimmed();
        if (sampleName.isEmpty()) break;

        SensorySample sample;
        sample.name = sampleName;
        bool hasData = false;
        for (int col = 2; col <= nMetrics + 1; ++col) {
            QVariant val = xlsx.read(row, col);
            if (val.isValid() && !val.toString().trimmed().isEmpty()) {
                double score = qBound(1.0, val.toDouble(), 9.0);
                if (score == 0.0) score = 5.0;
                int metricIdx = col - 2;
                QString metricName = (metricIdx < excelMetrics.size())
                    ? excelMetrics[metricIdx] : kSensoryMetrics[metricIdx];
                for (const QString& km : kSensoryMetrics) {
                    if (km.toLower() == metricName.toLower() ||
                        metricName.toLower().contains(km.toLower().left(6))) {
                        sample.scores[km] = score;
                        hasData = true;
                        break;
                    }
                }
            }
        }
        sample.comments = xlsx.read(row, nMetrics + 2).toString().trimmed();
        for (const QString& metric : kSensoryMetrics) {
            if (!sample.scores.contains(metric))
                sample.scores[metric] = 5.0;
        }
        if (hasData) samples.append(sample);
    }

    if (samples.isEmpty()) return 0;

    // Read metadata section (below data, after a gap)
    QString testerName, assessorName, media, date;
    QString control, primaryDifferences;
    bool isBlind = false;
    for (int row = samples.size() + 3; row <= samples.size() + 16; ++row) {
        QString label = xlsx.read(row, 1).toString().trimmed().toLower();
        QString value = xlsx.read(row, 2).toString().trimmed();
        if (label == "test title")              { /* testTitle already from filename */ }
        else if (label == "tester")               testerName         = value;
        else if (label == "assessor")             assessorName       = value;
        else if (label == "media")                media              = value;
        else if (label == "date")                 date               = value;
        else if (label == "control")              control            = value;
        else if (label == "blind?")               isBlind            = (value.toUpper() == "Y");
        else if (label == "primary difference(s)") primaryDifferences = value;
    }

    SensorySession sess;
    sess.sessionName       = testTitle + (testerName.isEmpty() ? "" : " - " + testerName);
    sess.testTitle         = testTitle;
    sess.testerName        = testerName;
    sess.assessorName      = assessorName;
    sess.media             = media;
    sess.date              = date.isEmpty() ? dateNow : date;
    sess.timestamp         = tsNow;
    sess.control           = control;
    sess.isBlind           = isBlind;
    sess.primaryDifferences = primaryDifferences;
    sess.samples           = samples;

    m_sessions.append(sess);
    m_savePath = filePath;
    return 1;
}

// ─────────────────────────────────────────────────────────────────────────────
// Load Excel — standardized template (multi-tester, device-block structure)
// ─────────────────────────────────────────────────────────────────────────────

int SensoryPanel::loadExcelStandardFormat(QXlsx::Document& xlsx,
                                           const QStringList& excelMetrics,
                                           const QString& testTitle)
{
    QString dateNow = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    QString tsNow   = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    struct PanelistEntry {
        QString testerName;
        QString deviceName;
        QMap<QString, double> scores;
        QString comments;
    };

    QVector<PanelistEntry> allEntries;
    QString currentDevice;
    bool inBlock = false;

    QString a1 = xlsx.read(1, 1).toString().trimmed();
    if (!a1.isEmpty() && !a1.toLower().contains("averages")) {
        currentDevice = a1;
        inBlock = true;
    }

    for (int row = 2; row <= 29; ++row) {
        QString cellA = xlsx.read(row, 1).toString().trimmed();

        bool rowEmpty = cellA.isEmpty();
        if (rowEmpty) {
            bool anyData = false;
            for (int c = 2; c <= 7; ++c) {
                if (!xlsx.read(row, c).toString().trimmed().isEmpty()) {
                    anyData = true;
                    break;
                }
            }
            if (!anyData) {
                inBlock = false;
                continue;
            }
        }

        if (cellA.toLower().contains("average"))
            break;

        if (!inBlock && !cellA.isEmpty()) {
            bool hasScores = false;
            for (int c = 2; c <= 6; ++c) {
                QVariant v = xlsx.read(row, c);
                if (v.isValid() && !v.toString().trimmed().isEmpty()) {
                    bool ok;
                    double val = v.toDouble(&ok);
                    if (ok && val >= 1.0 && val <= 9.0) { hasScores = true; break; }
                }
            }

            if (!hasScores) {
                currentDevice = cellA;
                inBlock = true;
                continue;
            } else {
                inBlock = true;
            }
        }

        if (!inBlock && cellA.isEmpty())
            continue;

        if (!cellA.isEmpty()) {
            PanelistEntry entry;
            entry.testerName = cellA;
            entry.deviceName = currentDevice;

            bool hasData = false;
            for (int col = 2; col <= 6; ++col) {
                QVariant val = xlsx.read(row, col);
                if (val.isValid() && !val.toString().trimmed().isEmpty()) {
                    double score = qBound(1.0, val.toDouble(), 9.0);
                    if (score == 0.0) score = 5.0;
                    int metricIdx = col - 2;
                    QString metricName = (metricIdx < excelMetrics.size())
                        ? excelMetrics[metricIdx] : kSensoryMetrics[metricIdx];
                    for (const QString& km : kSensoryMetrics) {
                        if (km.toLower() == metricName.toLower() ||
                            metricName.toLower().contains(km.toLower().left(6))) {
                            entry.scores[km] = score;
                            hasData = true;
                            break;
                        }
                    }
                }
            }
            entry.comments = xlsx.read(row, 7).toString().trimmed();

            if (hasData)
                allEntries.append(entry);
        }
    }

    QMap<QString, QVector<int>> testerIndices;
    QStringList testerOrder;
    for (int i = 0; i < allEntries.size(); ++i) {
        const QString& tn = allEntries[i].testerName;
        if (!testerIndices.contains(tn)) testerOrder.append(tn);
        testerIndices[tn].append(i);
    }

    int loaded = 0;
    for (const QString& testerName : testerOrder) {
        SensorySession sess;
        sess.sessionName = testTitle + " - " + testerName;
        sess.testTitle   = testTitle;
        sess.testerName  = testerName;
        sess.date        = dateNow;
        sess.timestamp   = tsNow;

        for (int idx : testerIndices[testerName]) {
            const PanelistEntry& e = allEntries[idx];
            SensorySample sample;
            sample.name     = e.deviceName;
            sample.scores   = e.scores;
            sample.comments = e.comments;
            for (const QString& metric : kSensoryMetrics) {
                if (!sample.scores.contains(metric))
                    sample.scores[metric] = 5.0;
            }
            sess.samples.append(sample);
        }

        if (!sess.samples.isEmpty()) {
            m_sessions.append(sess);
            ++loaded;
        }
    }
    return loaded;
}

// ─────────────────────────────────────────────────────────────────────────────
// Load from Database
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::loadFromDatabase()
{
    if (!m_db) {
        QMessageBox::warning(this, "Database", "No database connection available.");
        return;
    }

    QVector<SensoryRecord> records = m_db->listSensoryRecords();
    if (records.isEmpty()) {
        QMessageBox::information(this, "Database", "No sensory sessions found in the database.");
        return;
    }

    QDialog picker(this);
    picker.setWindowTitle("Load from Database");
    picker.setMinimumSize(700, 400);
    picker.resize(750, 500);

    auto* layout = new QVBoxLayout(&picker);

    auto* tree = new QTreeWidget;
    tree->setHeaderLabels({"Test Title / Tester", "Assessor", "Media", "Date", "Samples"});
    tree->setSelectionMode(QAbstractItemView::ExtendedSelection);
    tree->setRootIsDecorated(true);

    QMap<QString, QVector<const SensoryRecord*>> groups;
    QStringList order;
    for (const SensoryRecord& r : records) {
        QString title = r.testTitle.isEmpty() ? QStringLiteral("(untitled)") : r.testTitle;
        if (!groups.contains(title)) order.append(title);
        groups[title].append(&r);
    }

    for (const QString& title : order) {
        const auto& recs = groups[title];
        auto* parent = new QTreeWidgetItem(tree);
        parent->setText(0, title);
        parent->setText(4, QString("%1 tester(s)").arg(recs.size()));
        parent->setData(0, Qt::UserRole, -1);
        QFont f = parent->font(0);
        f.setBold(true);
        parent->setFont(0, f);

        for (const SensoryRecord* r : recs) {
            auto* child = new QTreeWidgetItem(parent);
            child->setText(0, r->testerName.isEmpty() ? r->sessionName : r->testerName);
            child->setText(1, r->assessorName);
            child->setText(2, r->media);
            child->setText(3, r->date);
            child->setText(4, QString::number(r->sampleCount));
            child->setData(0, Qt::UserRole, r->id);
        }

        parent->setExpanded(true);
    }

    tree->header()->setSectionResizeMode(0, QHeaderView::Stretch);
    for (int i = 1; i < 5; ++i)
        tree->header()->setSectionResizeMode(i, QHeaderView::ResizeToContents);

    layout->addWidget(tree, 1);

    auto* btnRow = new QHBoxLayout;
    btnRow->addStretch();
    auto* loadBtn = new QPushButton("Load Selected");
    auto* cancelBtn = new QPushButton("Cancel");
    btnRow->addWidget(loadBtn);
    btnRow->addWidget(cancelBtn);
    layout->addLayout(btnRow);

    connect(cancelBtn, &QPushButton::clicked, &picker, &QDialog::reject);
    connect(loadBtn, &QPushButton::clicked, &picker, &QDialog::accept);
    connect(tree, &QTreeWidget::itemDoubleClicked, &picker, &QDialog::accept);

    if (picker.exec() != QDialog::Accepted) return;

    QVector<int> selectedIds;
    const auto selected = tree->selectedItems();
    for (QTreeWidgetItem* item : selected) {
        int id = item->data(0, Qt::UserRole).toInt();
        if (id > 0) {
            selectedIds.append(id);
        } else if (id == -1) {
            for (int c = 0; c < item->childCount(); ++c) {
                int childId = item->child(c)->data(0, Qt::UserRole).toInt();
                if (childId > 0 && !selectedIds.contains(childId))
                    selectedIds.append(childId);
            }
        }
    }

    if (selectedIds.isEmpty()) return;

    QVector<SensorySession> sessions;
    for (int id : selectedIds) {
        SensorySession sess = m_db->loadSensorySession(id);
        if (!sess.samples.isEmpty())
            sessions.append(sess);
    }

    if (sessions.isEmpty()) {
        QMessageBox::warning(this, "Database", "Could not load the selected session(s).");
        return;
    }

    loadSessions(sessions);
}

// ─────────────────────────────────────────────────────────────────────────────
// DATAVIEWER-4: DB-authoritative export source
// ─────────────────────────────────────────────────────────────────────────────
//
// Every sensory export (PowerPoint report, Excel, CSV, JSON) reads the in-memory
// session model, which can hold stale per-metric scores — this client may have
// missed a LiveSync NOTIFY, or another user edited a score concurrently. The DB
// blob is the single source of truth for scores. This helper makes exports
// score-authoritative:
//   1. Flush our own pending per-cell edits to the DB once, so the DB row holds
//      this client's latest scores before we re-fetch.
//   2. For each persisted session, re-fetch the DB row and overlay ONLY the
//      kSensoryMetrics score values onto a copy of the in-memory session.
//
// Scores are taken from the DB; everything else (header metadata, sample names,
// comments, device props, AND non-serialized fields the report needs such as
// imagePaths/imageLayouts/imageCrops/id/version) stays in-memory-authoritative.
// We deliberately do NOT round-trip the struct through sensorySessionFromJson:
// that JSON contract omits the image fields and the persistence anchors, so a
// round-trip would silently drop every image from exported reports. Instead we
// reuse the pure mergeSensoryPreservingDbScores() at the JSON layer purely to
// compute the authoritative per-metric values, then copy those scalars back into
// the live struct. Unsaved sessions (id <= 0) and rows that have since
// disappeared pass through unchanged.
QVector<SensorySession> SensoryPanel::dbAuthoritativeSessions(
        const QVector<SensorySession>& inMem)
{
    if (m_liveSync && !m_liveSync->flushNowAndWait()) {
        // v2.5.0 Task 3 (RC2): flushNowAndWait() returns false on EITHER a drain
        // timeout OR a nested re-entrant flush. Either way it's safe to proceed —
        // the dirty-aware merge below keeps locally-edited scores authoritative
        // even if the worker drain didn't confirm in time.
        qWarning() << "SensoryPanel::dbAuthoritativeSessions: LiveSync flush did not "
                      "complete (timeout or nested flush); proceeding with dirty-aware "
                      "merge (pending="
                   << m_liveSync->pendingCount() << ")";
    }
    if (!m_db) return inMem;

    QVector<SensorySession> out;
    out.reserve(inMem.size());
    for (const SensorySession& s : inMem) {
        if (s.id <= 0) { out.append(s); continue; }            // never persisted
        const SensorySession dbSess = m_db->loadSensorySession(s.id);
        if (dbSess.id <= 0) { out.append(s); continue; }       // row gone / load failed

        // Compute DB-authoritative scores at the JSON layer, then overlay them
        // back onto a copy of the in-memory struct so all non-JSON fields
        // (images, anchors) survive untouched.
        // v2.5.0 Task 3 (RC2): exports honor the same dirty-cell set as the
        // whole-session save, so a report reflects scores the user edited this
        // run even when LiveSync never streamed them.
        const QJsonObject merged = mergeSensoryPreservingDbScores(
            sensorySessionToJson(s), sensorySessionToJson(dbSess), s.dirtyCells);
        const QJsonArray mergedSamples = merged.value("samples").toArray();

        SensorySession authoritative = s;  // keep every in-memory field
        for (int i = 0; i < authoritative.samples.size()
                        && i < mergedSamples.size(); ++i) {
            const QJsonObject ms = mergedSamples[i].toObject();
            for (const QString& metric : kSensoryMetrics) {
                if (ms.contains(metric))
                    authoritative.samples[i].scores[metric] = ms.value(metric).toDouble();
            }
        }
        out.append(authoritative);
    }
    return out;
}

// ─────────────────────────────────────────────────────────────────────────────
// Generate Sensory Report
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::generateFullReport()
{
    saveCurrentTester();

    if (m_sessions.isEmpty()) {
        QMessageBox::warning(this, "No Data", "No sessions loaded.");
        return;
    }

    // Show a selection dialog — user can Ctrl+Click or Shift+Click to select
    QDialog picker(this);
    picker.setWindowTitle("Select Sessions for Combined Report");
    picker.setMinimumSize(500, 400);
    picker.resize(550, 450);

    auto* layout = new QVBoxLayout(&picker);
    layout->addWidget(new QLabel("Select sessions to include (Ctrl+Click or Shift+Click):"));

    auto* list = new QListWidget;
    list->setSelectionMode(QAbstractItemView::ExtendedSelection);
    for (int i = 0; i < m_sessions.size(); ++i) {
        auto* item = new QListWidgetItem(sessionLabel(m_sessions[i]));
        item->setData(Qt::UserRole, i);
        list->addItem(item);
        item->setSelected(true);  // select all by default
    }
    layout->addWidget(list, 1);

    auto* btnRow = new QHBoxLayout;
    btnRow->addStretch();
    auto* genBtn = new QPushButton("Generate Report");
    auto* cancelBtn = new QPushButton("Cancel");
    btnRow->addWidget(genBtn);
    btnRow->addWidget(cancelBtn);
    layout->addLayout(btnRow);

    connect(cancelBtn, &QPushButton::clicked, &picker, &QDialog::reject);
    connect(genBtn, &QPushButton::clicked, &picker, &QDialog::accept);

    if (picker.exec() != QDialog::Accepted) return;

    QVector<SensorySession> selected;
    for (QListWidgetItem* item : list->selectedItems()) {
        int idx = item->data(Qt::UserRole).toInt();
        if (idx >= 0 && idx < m_sessions.size())
            selected.append(m_sessions[idx]);
    }

    if (selected.isEmpty()) {
        QMessageBox::warning(this, "No Selection", "No sessions selected.");
        return;
    }

    // DATAVIEWER-4: scores in the report must come from the DB, not stale
    // in-memory state. Everything else (metadata, images) stays in-memory.
    selected = dbAuthoritativeSessions(selected);

    auto* src = new SensoryReportSource(selected, m_db);
    ReportPreviewDialog dlg(src, this);
    const int rc = dlg.exec();
    delete src;

    // Dialog handles its own QFileDialog + writePptx + success confirmation.
    // Nothing more to do here; just return.
    Q_UNUSED(rc);
}

// ─────────────────────────────────────────────────────────────────────────────
// Generate Statistics Report (.csv)
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::generateStats()
{
    if (m_cards.isEmpty()) {
        QMessageBox::warning(this, "No Data", "Add at least one sample before generating stats.");
        return;
    }

    QString path = QFileDialog::getSaveFileName(
        this, "Save Statistics Report", OutputPaths::resolveDir(ReportMode::Sensory,m_lastBrowseDir),
        "CSV files (*.csv);;All files (*)");
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    writeStatsCsv(path);
}

void SensoryPanel::writeStatsCsv(const QString& path)
{
    // DATAVIEWER-4: statistics are computed over scores, so they must reflect
    // the DB-authoritative values rather than possibly-stale in-memory state.
    SensorySession sess = buildSession();
    sess = dbAuthoritativeSessions({sess}).value(0, sess);
    int n = sess.samples.size();

    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QMessageBox::warning(this, "Save Error",
                             "Could not save statistics report:\n" + path);
        return;
    }
    QTextStream out(&f);

    out << "Sensory Evaluation Statistics Report\n";
    if (!sess.testTitle.isEmpty())
        out << "Test Title," << sess.testTitle << "\n";
    out << "Date," << sess.date << "\n";
    out << "Assessor," << sess.assessorName << "\n";
    if (!sess.testerName.isEmpty())
        out << "Tester," << sess.testerName << "\n";
    out << "Media," << sess.media << "\n";
    out << "Sample Count," << n << "\n\n";

    out << "Metric,Sample Size,Mean,Std Dev,Std Error,Min,Max,Range,Outliers,Chi-Square\n";

    for (const QString& metric : kSensoryMetrics) {
        QVector<double> vals;
        double sum = 0;
        for (const SensorySample& s : sess.samples) {
            double v = s.scores.value(metric, 5.0);
            vals.append(v);
            sum += v;
        }

        double mean = (n > 0) ? sum / n : 0;
        double sumSq = 0;
        for (double v : vals) sumSq += (v - mean) * (v - mean);
        double variance = (n > 1) ? sumSq / (n - 1) : 0;
        double stddev = std::sqrt(qMax(0.0, variance));
        double stderr_ = (n > 0) ? stddev / std::sqrt(static_cast<double>(n)) : 0;

        double minVal = 9.0, maxVal = 1.0;
        for (double v : vals) { minVal = qMin(minVal, v); maxVal = qMax(maxVal, v); }

        QStringList outliers;
        for (int i = 0; i < vals.size(); ++i) {
            if (stddev > 0 && std::abs(vals[i] - mean) > 2.0 * stddev) {
                QString sName = (i < sess.samples.size() && !sess.samples[i].name.isEmpty())
                    ? sess.samples[i].name : QString("Sample %1").arg(i + 1);
                outliers << QString("%1(%2)").arg(sName).arg(vals[i], 0, 'f', 1);
            }
        }

        double chiSq = 0;
        double expected = static_cast<double>(n) / 9.0;
        if (expected > 0) {
            // Use rounded values for chi-square (integer bins)
            QMap<int, int> freq;
            for (double v : vals) freq[qRound(v)]++;
            for (int score = 1; score <= 9; ++score) {
                double obs = freq.value(score, 0);
                chiSq += (obs - expected) * (obs - expected) / expected;
            }
        }

        out << "\"" << metric << "\""
            << "," << n
            << "," << QString::number(mean, 'f', 2)
            << "," << QString::number(stddev, 'f', 2)
            << "," << QString::number(stderr_, 'f', 3)
            << "," << QString::number(minVal, 'f', 1)
            << "," << QString::number(maxVal, 'f', 1)
            << "," << QString::number(maxVal - minVal, 'f', 1)
            << "," << (outliers.isEmpty() ? "None" : outliers.join("; "))
            << "," << QString::number(chiSq, 'f', 2)
            << "\n";
    }

    out << "\nRaw Data\n";
    out << "Sample";
    for (const QString& m : kSensoryMetrics) out << "," << "\"" << m << "\"";
    out << ",Comments\n";

    for (const SensorySample& s : sess.samples) {
        out << "\"" << (s.name.isEmpty() ? QString("Sample") : s.name) << "\"";
        for (const QString& m : kSensoryMetrics)
            out << "," << QString::number(s.scores.value(m, 5.0), 'f', 1);
        out << ",\"" << QString(s.comments).replace("\"", "\"\"") << "\"\n";
    }

    f.close();

    QMessageBox::information(this, "Statistics Report Saved",
        "Statistics report saved:\n" + QFileInfo(path).fileName());
}

// ─────────────────────────────────────────────────────────────────────────────
// Combined report: one slide per session (static)
// ─────────────────────────────────────────────────────────────────────────────

bool SensoryPanel::generateCombinedPptx(const QVector<SensorySession>& sessions,
                                         const QString& filePath,
                                         QString& errorOut)
{
    // DATAVIEWER-4 note: this static entry point does NOT route through
    // dbAuthoritativeSessions() — and must not. It has no instance access to
    // m_db / m_liveSync, and its sole caller (DatabaseBrowserDialog) builds its
    // `sessions` list by loading rows straight from the DB, so the scores are
    // already authoritative. The in-memory-model exports (generateFullReport,
    // save, writeStatsCsv) are the ones that need the reconciliation.
    //
    // Thin caller for the legacy fast path. The whole body now lives in
    // SensoryReportSource::writeSensoryPptx so the new IReportSource entry
    // point and this legacy entry point share one renderer.
    //
    // Pass an empty ReportLayout so the 5-arg addContentSlide falls through to
    // the legacy in-line positions (wrap-aware table height, aspect-matched
    // radar). computeDefaultLayout is intentionally NOT used here — it's a
    // simpler shape designed for the preview dialog (square radar; flat
    // 0.50 + N/3 table height; cumulative sized from the FIRST session's
    // sample count) and would override the legacy in-line wrap-aware math via
    // the 5-arg PptxWriter::addContentSlide overload, breaking visual
    // equivalence with the pre-Task-8 output.
    const ReportLayout emptyLayout;
    QString err;
    const bool ok = SensoryReportSource::writeSensoryPptx(
        sessions, emptyLayout, /*excludedSamples=*/{}, filePath, &err);
    if (!ok) errorOut = err;
    return ok;
}

// ─────────────────────────────────────────────────────────────────────────────
// v2.0.1 LiveSync wiring
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::setLiveSync(LiveSync* sync)
{
    if (m_liveSync == sync) return;
    if (m_liveSync) disconnect(m_liveSync, nullptr, this, nullptr);
    m_liveSync = sync;
    if (m_liveSync) {
        connect(m_liveSync, &LiveSync::cellChanged, this,
                &SensoryPanel::onRemoteCellChanged);
    }
}

void SensoryPanel::onRemoteCellChanged(const QString& table, qint64 rowId,
                                       const QString& column,
                                       const QVariant& newValue)
{
    if (table != QLatin1String("sensory_sessions")) return;
    if (!column.startsWith(QLatin1String("json_path:samples["))) return;

    // Parse "json_path:samples[N].<field>" into (sampleIdx, fieldPath).
    const QString rest = column.mid(QStringLiteral("json_path:samples[").size());
    const int rbr = rest.indexOf(QLatin1Char(']'));
    if (rbr < 0) return;
    bool okIdx = false;
    const int idx = rest.left(rbr).toInt(&okIdx);
    if (!okIdx) return;
    if (rest.size() <= rbr + 2) return;
    const QString fieldPath = rest.mid(rbr + 2);  // skip "]."

    int sessIdx = -1;
    for (int i = 0; i < m_sessions.size(); ++i) {
        if (static_cast<qint64>(m_sessions[i].id) == rowId) { sessIdx = i; break; }
    }
    if (sessIdx < 0) return;
    if (idx < 0 || idx >= m_sessions[sessIdx].samples.size()) return;
    applyRemoteFieldToSample(m_sessions[sessIdx].samples[idx], fieldPath, newValue);

    if (sessIdx != m_currentTesterIdx) return;  // not visible
    if (idx >= m_cards.size()) return;
    SampleCard* card = m_cards[idx];
    QSignalBlocker b(card);
    card->fromSample(m_sessions[sessIdx].samples[idx]);
}

void SensoryPanel::applyRemoteFieldToSample(SensorySample& s,
                                            const QString& fieldPath,
                                            const QVariant& value)
{
    if (fieldPath == QLatin1String("name"))               s.name = value.toString();
    else if (fieldPath == QLatin1String("comments"))      s.comments = value.toString();
    else if (fieldPath == QLatin1String("voltage"))       s.voltage = value.toDouble();
    else if (fieldPath == QLatin1String("resistance"))    s.resistance = value.toDouble();
    else if (fieldPath == QLatin1String("power"))         s.power = value.toDouble();
    else if (fieldPath == QLatin1String("heating_technology"))
                                                          s.heatingTechnology = value.toString();
    else if (fieldPath == QLatin1String("power_type"))    s.powerType = value.toString();
    else if (fieldPath == QLatin1String("puff_length_sec")) s.puffLengthSec = value.toDouble();
    else if (kSensoryMetrics.contains(fieldPath))         s.scores[fieldPath] = value.toDouble();
}

} // namespace DVE

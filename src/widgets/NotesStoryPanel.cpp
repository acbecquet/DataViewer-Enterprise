#include "NotesStoryPanel.h"
#include "FlowLayout.h"
#include "../utils/AppTheme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QScrollArea>
#include <QScrollBar>
#include <QFrame>
#include <QLabel>
#include <QPlainTextEdit>
#include <QLineEdit>
#include <QValidator>
#include <QRegularExpression>
#include <QTimer>
#include <QToolButton>
#include <QTableWidget>
#include <QHeaderView>

namespace DVE {

// Canonical per-row puffing-regime format. The standard is "<vol>mL/<puff>s/
// <interval>s" (e.g. 60mL/3s/30s, 100mL/2.5s/15s), but real templates also
// store the abbreviated "<puff>s/<interval>s" (e.g. 3s/30s) — so the mL prefix
// is OPTIONAL here, and only genuinely malformed text (typos that would
// fragment regime grouping in processing) flags red. Empty = "(unspecified)",
// also treated as valid (no red).
static const QRegularExpression kRegimeFormat(
    QStringLiteral("^(\\d+(?:\\.\\d+)?mL/)?\\d+(?:\\.\\d+)?s/\\d+(?:\\.\\d+)?s$"),
    QRegularExpression::CaseInsensitiveOption);

static bool regimeLooksValid(const QString& text) {
    const QString t = text.trimmed();
    return t.isEmpty() || kRegimeFormat.match(t).hasMatch();
}

NotesStoryPanel::NotesStoryPanel(QWidget* parent) : QWidget(parent) {
    auto* outer = new QVBoxLayout(this);
    outer->setContentsMargins(0, 0, 0, 0);
    m_scroll = new QScrollArea(this);
    m_scroll->setWidgetResizable(true);
    m_scroll->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_scroll->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    m_scroll->setFrameShape(QFrame::NoFrame);
    m_content = new QWidget(m_scroll);
    m_vbox = new QVBoxLayout(m_content);
    m_vbox->setContentsMargins(8, 8, 8, 8);
    m_vbox->setSpacing(8);
    m_vbox->addStretch(1);
    m_scroll->setWidget(m_content);
    outer->addWidget(m_scroll);
    // Mirror the Navigator dock's min width (see MainWindow setupCentralWidget)
    // so the panel can shrink with it; content uses FlowLayout/stretch tables
    // so nothing clips down to this width.
    setMinimumWidth(220);
}

void NotesStoryPanel::clear() { m_sample = SampleResult{}; m_excluded.clear(); rebuild(); }

void NotesStoryPanel::setSample(const SampleResult& sample, const QSet<int>& excluded, bool perRowRegime) {
    m_sample = sample; m_excluded = excluded; m_perRowRegime = perRowRegime; rebuild();
}

void NotesStoryPanel::rebuild() {
    m_cardByRow.clear();
    QLayoutItem* item;
    while ((item = m_vbox->takeAt(0)) != nullptr) {
        if (QWidget* w = item->widget()) w->deleteLater();
        delete item;
    }
    const QVector<StorySegment> segs = NotesStory::build(m_sample, m_excluded);
    for (int i = 0; i < segs.size(); ++i) {
        const StorySegment& seg = segs[i];
        QWidget* w = (seg.kind == StorySegment::Note)
            ? buildNoteCard(m_sample, seg.rowIndex, seg.excluded, m_perRowRegime)
            : buildSummaryBar(m_sample, seg.summary, i, m_perRowRegime);
        m_vbox->addWidget(w);
        if (seg.kind == StorySegment::Note) m_cardByRow.insert(seg.rowIndex, w);
    }
    m_vbox->addStretch(1);
}

// ── Task 3: editable note card ────────────────────────────────────────────────
QWidget* NotesStoryPanel::buildNoteCard(const SampleResult& s, int rowIndex, bool excluded, bool perRowRegime) {
    const DataRow& dr = s.rows[rowIndex];
    auto* card = new QFrame;
    card->setObjectName("storyCard");
    card->setStyleSheet(QString("#storyCard{border:1px solid %1;border-radius:6px;background:%2;}")
        .arg(AppTheme::border().name(), excluded ? QStringLiteral("#F8F0F0") : QStringLiteral("#FFFFFF")));
    auto* v = new QVBoxLayout(card);
    v->setContentsMargins(8, 8, 8, 8); v->setSpacing(6);

    auto* hdr = new QLabel(QString("Puff %1%2").arg(int(dr.puffs))
        .arg(excluded ? QStringLiteral("   • excluded") : QString()));
    hdr->setStyleSheet(QString("font-weight:600;color:%1;").arg(AppTheme::accent().name()));
    v->addWidget(hdr);

    auto* ctx = new QLabel(QString("TPM %1 · avg %2 · var %3")
        .arg(dr.tpm,0,'f',2).arg(s.averageTPM,0,'f',2).arg(dr.variationTPM,0,'f',2));
    ctx->setStyleSheet("color:#666;font-size:8pt;");
    v->addWidget(ctx);

    auto* note = new QPlainTextEdit(dr.notes);     // initial text set in ctor, before connect
    note->setPlaceholderText(QStringLiteral("Add a note…"));
    note->setFixedHeight(48);
    note->setLineWrapMode(QPlainTextEdit::WidgetWidth);
    auto* noteTimer = new QTimer(note);
    noteTimer->setSingleShot(true); noteTimer->setInterval(700);
    connect(note, &QPlainTextEdit::textChanged, noteTimer, qOverload<>(&QTimer::start));
    connect(noteTimer, &QTimer::timeout, this, [this, rowIndex, note]() {
        emit cellEdited(rowIndex, Cols::NOTES, note->toPlainText());
    });
    v->addWidget(note);

    auto neighbourTpm = [&s, rowIndex](int dir)->QString {
        for (int i = rowIndex + dir; i >= 0 && i < s.rows.size(); i += dir)
            if (s.rows[i].beforeWeight != 0.0 && s.rows[i].afterWeight != 0.0)
                return QString::number(s.rows[i].tpm, 'f', 2);
        return QStringLiteral("—");
    };
    // Read-only context chips — in a FlowLayout so they wrap to the next line
    // instead of clipping as the panel narrows (no horizontal scrollbar).
    auto roChip = [](const QString& t){ auto* l=new QLabel(t); l->setStyleSheet(
        QStringLiteral("background:#F0F0F0;color:#555;border-radius:6px;padding:2px 6px;font-size:8pt;"));
        l->setSizePolicy(QSizePolicy::Fixed, QSizePolicy::Fixed); return l; };
    auto* chips = new FlowLayout(nullptr, 0, 6, 4);
    chips->setRowAlignment(Qt::AlignLeft);
    chips->addWidget(roChip(QString("TPM before %1").arg(neighbourTpm(-1))));
    chips->addWidget(roChip(QString("TPM after %1").arg(neighbourTpm(+1))));
    chips->addWidget(roChip(QString("Draw %1").arg(dr.drawPressure,0,'f',1)));
    v->addLayout(chips);

    // Editor controls — each label+control bundled in a small container so the
    // pair wraps together (never a label on one row, its control on the next).
    // The FlowLayout wraps them to additional rows as the panel narrows, which
    // is what splits Clog/Smell/Regime onto two rows at tight widths.
    auto bundle = [](const QString& labelText, QWidget* ctrl)->QWidget* {
        auto* host = new QWidget;
        auto* hl = new QHBoxLayout(host);
        hl->setContentsMargins(0,0,0,0); hl->setSpacing(3);
        auto* lbl = new QLabel(labelText);
        lbl->setStyleSheet(QStringLiteral("font-size:8pt;color:#555;"));
        hl->addWidget(lbl);
        hl->addWidget(ctrl);
        host->setSizePolicy(QSizePolicy::Fixed, QSizePolicy::Fixed);
        return host;
    };
    auto* edits = new FlowLayout(nullptr, 0, 6, 4);
    edits->setRowAlignment(Qt::AlignLeft);

    // Clog is a plain Y/N text field (no dropdown arrow) — type the value.
    auto* clog = new QLineEdit;
    clog->setMaxLength(1); clog->setFixedWidth(34); clog->setAlignment(Qt::AlignCenter);
    clog->setValidator(new QRegularExpressionValidator(QRegularExpression(QStringLiteral("[YyNn]?")), clog));
    clog->setText(dr.clog.trimmed().toUpper() == QLatin1String("Y") ? QStringLiteral("Y") : QStringLiteral("N")); // before connect
    connect(clog, &QLineEdit::editingFinished, this, [this, rowIndex, clog]() {
        QString t = clog->text().trimmed().toUpper();
        if (t != QLatin1String("Y") && t != QLatin1String("N")) t = QStringLiteral("N");
        if (clog->text() != t) clog->setText(t);
        emit cellEdited(rowIndex, Cols::CLOG, t);
    });
    edits->addWidget(bundle(QStringLiteral("Clog"), clog));

    // Smell is a 0-4 text field (no spin arrows) — the validator limits input.
    auto* smell = new QLineEdit;
    smell->setMaxLength(1); smell->setFixedWidth(34); smell->setAlignment(Qt::AlignCenter);
    smell->setValidator(new QIntValidator(0, 4, smell));
    smell->setText(QString::number(dr.smell.toInt()));   // before connect
    connect(smell, &QLineEdit::editingFinished, this, [this, rowIndex, smell]() {
        QString t = smell->text().trimmed();
        if (t.isEmpty()) t = QStringLiteral("0");
        emit cellEdited(rowIndex, Cols::SMELL, t);
    });
    edits->addWidget(bundle(QStringLiteral("Smell"), smell));

    if (perRowRegime) {
        // Free-text regime — flag red when it doesn't match the standard format
        // so a typo can't silently fragment regime grouping in processing.
        auto* regime = new QLineEdit;
        regime->setMinimumWidth(110);
        regime->setPlaceholderText(QStringLiteral("e.g. 60mL/3s/30s"));
        regime->setText(dr.puffingRegime);               // before connect
        auto styleRegime = [regime]() {
            regime->setStyleSheet(regimeLooksValid(regime->text())
                ? QString()
                : QStringLiteral("border:1px solid #CC0000; background:#FFF0F0;"));
        };
        styleRegime();
        connect(regime, &QLineEdit::textChanged, regime, [styleRegime](const QString&){ styleRegime(); });
        connect(regime, &QLineEdit::editingFinished, this, [this, rowIndex, regime]() {
            const QString t = regime->text().trimmed();
            if (regimeLooksValid(t))                      // only persist a recognisable regime
                emit cellEdited(rowIndex, Cols::RESISTANCE, t);
        });
        edits->addWidget(bundle(QStringLiteral("Regime"), regime));
    }
    v->addLayout(edits);
    return card;
}
QWidget* NotesStoryPanel::buildSummaryBar(const SampleResult& s, const StorySummary& sum, int /*segIndex*/, bool /*perRowRegime*/) {
    auto* box = new QWidget;
    auto* bv = new QVBoxLayout(box);
    bv->setContentsMargins(0,0,0,0); bv->setSpacing(0);

    auto* btn = new QToolButton;
    btn->setCheckable(true);
    btn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    btn->setArrowType(Qt::RightArrow);
    btn->setAutoRaise(true);
    btn->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);
    btn->setText(sum.count > 0
        ? QString("Puffs %1\xe2\x80\x93%2 \xc2\xb7 avg %3 \xc2\xb7 var %4").arg(sum.puffStart).arg(sum.puffEnd)
              .arg(sum.avgTpm,0,'f',2).arg(sum.varTpm,0,'f',2)
        : QString("Puffs %1\xe2\x80\x93%2 \xc2\xb7 (all excluded)").arg(sum.puffStart).arg(sum.puffEnd));
    btn->setStyleSheet(QStringLiteral("QToolButton{font-size:8pt;color:#555;border:none;text-align:left;padding:3px 2px;}"));
    bv->addWidget(btn);

    // Compact, full-width table with real gridlines. Columns stretch to fill
    // the panel; Smell/Clog are cell widgets, the rest are read-only items.
    // The OUTER panel scrolls — this table has BOTH scrollbars off and a fixed
    // height sized exactly to its content, so it never nests a scroll region.
    enum { ColPuff = 0, ColTpm, ColDraw, ColSmell, ColClog, ColCount };
    const int nRows = int(sum.rowIndices.size());
    auto* table = new QTableWidget(nRows, ColCount);
    table->setVisible(false);
    table->setHorizontalHeaderLabels({QStringLiteral("Puff"), QStringLiteral("TPM"),
        QStringLiteral("Draw"), QStringLiteral("Smell"), QStringLiteral("Clog")});
    table->verticalHeader()->setVisible(false);
    table->setShowGrid(true);
    table->setEditTriggers(QAbstractItemView::NoEditTriggers);
    table->setSelectionMode(QAbstractItemView::NoSelection);
    table->setFocusPolicy(Qt::NoFocus);
    table->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    table->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    table->setFrameShape(QFrame::StyledPanel);
    {
        QFont tf = table->font(); tf.setPointSize(8); table->setFont(tf);
    }
    {
        QFont hf = table->horizontalHeader()->font(); hf.setPointSize(8);
        table->horizontalHeader()->setFont(hf);
    }
    table->horizontalHeader()->setHighlightSections(false);
    // Numeric columns size to content; the editor columns stretch to fill the
    // remaining width so the whole table spans the panel with no right-side gap.
    table->horizontalHeader()->setSectionResizeMode(ColPuff,  QHeaderView::ResizeToContents);
    table->horizontalHeader()->setSectionResizeMode(ColTpm,   QHeaderView::ResizeToContents);
    table->horizontalHeader()->setSectionResizeMode(ColDraw,  QHeaderView::ResizeToContents);
    table->horizontalHeader()->setSectionResizeMode(ColSmell, QHeaderView::Stretch);
    table->horizontalHeader()->setSectionResizeMode(ColClog,  QHeaderView::Stretch);

    const int rowH = 22;
    for (int r = 0; r < nRows; ++r) {
        const int idx = sum.rowIndices[r];
        const DataRow& dr = s.rows[idx];
        const bool excl = m_excluded.contains(idx);
        table->setRowHeight(r, rowH);

        auto mkItem = [&](const QString& t){
            auto* it = new QTableWidgetItem(t);
            it->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
            it->setFlags(Qt::ItemIsEnabled);            // read-only, still enabled
            if (excl) it->setForeground(QColor(0xaa, 0xaa, 0xaa));
            return it;
        };
        table->setItem(r, ColPuff, mkItem(excl
            ? QString("%1 \xc2\xb7""excl").arg(int(dr.puffs))
            : QString::number(int(dr.puffs))));
        table->setItem(r, ColTpm,  mkItem(QString::number(dr.tpm,'f',2)));
        table->setItem(r, ColDraw, mkItem(QString::number(dr.drawPressure,'f',1)));

        auto* sm = new QLineEdit; sm->setMaxLength(1); sm->setAlignment(Qt::AlignCenter);
        sm->setValidator(new QIntValidator(0, 4, sm));
        sm->setText(QString::number(dr.smell.toInt()));   // value before connect
        sm->setStyleSheet(QStringLiteral("font-size:8pt;")); sm->setFocusPolicy(Qt::StrongFocus);
        connect(sm, &QLineEdit::editingFinished, this, [this, idx, sm]() {
            QString t = sm->text().trimmed(); if (t.isEmpty()) t = QStringLiteral("0");
            emit cellEdited(idx, Cols::SMELL, t);
        });
        table->setCellWidget(r, ColSmell, sm);

        auto* cl = new QLineEdit; cl->setMaxLength(1); cl->setAlignment(Qt::AlignCenter);
        cl->setValidator(new QRegularExpressionValidator(QRegularExpression(QStringLiteral("[YyNn]?")), cl));
        cl->setText(dr.clog.trimmed().toUpper() == QLatin1String("Y") ? QStringLiteral("Y") : QStringLiteral("N")); // before connect
        cl->setStyleSheet(QStringLiteral("font-size:8pt;"));
        connect(cl, &QLineEdit::editingFinished, this, [this, idx, cl]() {
            QString t = cl->text().trimmed().toUpper();
            if (t != QLatin1String("Y") && t != QLatin1String("N")) t = QStringLiteral("N");
            if (cl->text() != t) cl->setText(t);
            emit cellEdited(idx, Cols::CLOG, t);
        });
        table->setCellWidget(r, ColClog, cl);
    }

    // Fix the table's height to exactly fit header + all rows + frame so the
    // outer QScrollArea owns scrolling and no inner scrollbar appears and no
    // row is clipped.
    const int frame = 2 * table->frameWidth();
    table->setFixedHeight(table->horizontalHeader()->height() + nRows * rowH + frame);
    table->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);
    bv->addWidget(table);

    connect(btn, &QToolButton::toggled, table, [btn, table](bool on){
        table->setVisible(on);
        btn->setArrowType(on ? Qt::DownArrow : Qt::RightArrow);
    });
    return box;
}
void NotesStoryPanel::highlightRow(int dataRowIndex) {
    const auto it = m_cardByRow.constFind(dataRowIndex);
    if (it == m_cardByRow.constEnd() || !it.value()) return;   // not a note row → no-op
    QWidget* card = it.value();
    m_scroll->ensureWidgetVisible(card);
    // Flash an accent border for ~1.2s, then restore. Using `card` as the timer's
    // context object means the lambda is cancelled if the card is destroyed first.
    const QString base = card->styleSheet();
    card->setStyleSheet(base + QString("#storyCard{border:2px solid %1;}").arg(AppTheme::accent().name()));
    QTimer::singleShot(1200, card, [card, base]() { card->setStyleSheet(base); });
}

} // namespace DVE

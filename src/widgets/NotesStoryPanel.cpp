#include "NotesStoryPanel.h"
#include "FlowLayout.h"
#include "../utils/AppTheme.h"
#include "../pipeline/RegimeUtils.h"
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
#include <QToolTip>
#include <QTableWidget>
#include <QHeaderView>
#include <QEvent>
#include <QMouseEvent>

namespace DVE {

// Regime format/alias logic lives in RegimeUtils so the panel's regime editor
// validates consistently:
//   RegimeUtils::isStandardRegimeFormat() / RegimeUtils::canonicalRegime().

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

void NotesStoryPanel::showHint(const QString& message) {
    // Drop any sample data so a later clear()/setSample() rebuilds cleanly, then
    // replace the card stack with one centered, wrapping hint label.
    m_sample = SampleResult{};
    m_excluded.clear();
    m_cardByRow.clear();
    QLayoutItem* item;
    while ((item = m_vbox->takeAt(0)) != nullptr) {
        if (QWidget* w = item->widget()) w->deleteLater();
        delete item;
    }
    auto* hint = new QLabel(message);
    hint->setWordWrap(true);
    hint->setAlignment(Qt::AlignCenter);
    hint->setStyleSheet(QString("color:%1;padding:16px;font-size:9pt;")
        .arg(AppTheme::textSecondary().name()));
    m_vbox->addStretch(1);
    m_vbox->addWidget(hint);
    m_vbox->addStretch(1);
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

bool NotesStoryPanel::eventFilter(QObject* watched, QEvent* event) {
    if (event->type() == QEvent::MouseButtonPress) {
        const QVariant rowProp = watched->property("noteRow");
        if (rowProp.isValid()) {
            emit noteActivated(rowProp.toInt());
            // Don't consume the event — let the label paint its press state.
        }
    }
    return QWidget::eventFilter(watched, event);
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

    // Clicking the card's read-only chrome (header / context line) emphasises
    // this note's point on the plot (v1 note->plot linking). The editors below
    // consume their own clicks, so typing isn't hijacked.
    hdr->setProperty("noteRow", rowIndex);
    ctx->setProperty("noteRow", rowIndex);
    hdr->installEventFilter(this);
    ctx->installEventFilter(this);

    auto* note = new QPlainTextEdit(dr.notes);     // initial text set in ctor, before connect
    note->setPlaceholderText(QStringLiteral("Add a note…"));
    note->setMinimumHeight(48);
    note->setSizePolicy(QSizePolicy::Preferred, QSizePolicy::Minimum);
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
    clog->setMaxLength(1); clog->setMinimumWidth(34); clog->setAlignment(Qt::AlignCenter);
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
    smell->setMaxLength(1); smell->setMinimumWidth(34); smell->setAlignment(Qt::AlignCenter);
    smell->setValidator(new QIntValidator(0, 4, smell));
    smell->setText(QString::number(dr.smell.toInt()));   // before connect
    connect(smell, &QLineEdit::editingFinished, this, [this, rowIndex, smell]() {
        QString t = smell->text().trimmed();
        if (t.isEmpty()) t = QStringLiteral("0");
        emit cellEdited(rowIndex, Cols::SMELL, t);
    });
    edits->addWidget(bundle(QStringLiteral("Smell"), smell));

    if (perRowRegime) {
        // Free-text regime that ENFORCES the standard: on commit the value is
        // canonicalised (CORESTA -> 60mL/3s/30s) and saved only if it matches
        // the parametric format; anything else snaps back to the last accepted
        // value with a hint, so non-standard text (typos) can't enter the data.
        // A live red border previews the same check while typing.
        auto* regime = new QLineEdit;
        regime->setMinimumWidth(110);
        regime->setPlaceholderText(QStringLiteral("e.g. 60mL/3s/30s"));
        regime->setToolTip(QStringLiteral("Standard format, e.g. 60mL/3s/30s"));
        regime->setText(dr.puffingRegime);               // before connect
        auto styleRegime = [regime]() {
            const bool ok = RegimeUtils::isStandardRegimeFormat(
                RegimeUtils::canonicalRegime(regime->text()));
            regime->setStyleSheet(ok ? QString()
                : QStringLiteral("border:1px solid #CC0000; background:#FFF0F0;"));
        };
        styleRegime();
        connect(regime, &QLineEdit::textChanged, regime, [styleRegime](const QString&){ styleRegime(); });
        connect(regime, &QLineEdit::editingFinished, this,
            [this, rowIndex, regime, lastValid = dr.puffingRegime]() mutable {
                const QString canon = RegimeUtils::canonicalRegime(regime->text());
                if (RegimeUtils::isStandardRegimeFormat(canon)) {
                    if (regime->text() != canon) regime->setText(canon);   // show normalised form
                    lastValid = canon;
                    emit cellEdited(rowIndex, Cols::RESISTANCE, canon);
                } else {
                    regime->setText(lastValid);                            // visible snap-back
                    QToolTip::showText(regime->mapToGlobal(regime->rect().bottomLeft()),
                        QStringLiteral("Regime must be the standard format, e.g. 60mL/3s/30s"),
                        regime);
                }
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

    // Compact table with real gridlines. Columns size to content; Smell/Clog
    // are cell widgets, the rest are read-only items. The OUTER panel owns
    // vertical scroll (this table's height is fixed to fit every row); only a
    // horizontal scrollbar appears here, as needed, when the columns overflow.
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
    // Horizontal scroll AS-NEEDED so narrow panels scroll the columns instead
    // of squishing Smell/Clog until their headers/values clip. Vertical stays
    // off — the fixed height below fits every row, so the OUTER panel scrolls.
    table->setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    table->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    table->setFrameShape(QFrame::StyledPanel);
    // Tighter cell margins (esp. the Puff column) while keeping the gridlines.
    table->setStyleSheet(QStringLiteral(
        "QTableWidget{gridline-color:#D8D8D8;} QTableWidget::item{padding:0px 2px;}"));
    {
        QFont tf = table->font(); tf.setPointSize(8); table->setFont(tf);
    }
    {
        QFont hf = table->horizontalHeader()->font(); hf.setPointSize(8);
        table->horizontalHeader()->setFont(hf);
    }
    table->horizontalHeader()->setHighlightSections(false);
    // Every column sizes to its content/header (no Stretch) so nothing is
    // squished — when the total exceeds the panel width the horizontal scrollbar
    // appears instead of clipping Smell/Clog.
    for (int c = 0; c < ColCount; ++c)
        table->horizontalHeader()->setSectionResizeMode(c, QHeaderView::ResizeToContents);

    const int rowH = AppTheme::controlHeight();
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

    // Fix the table's height to fit header + all rows + frame, PLUS the
    // horizontal scrollbar's height so that when it appears (narrow panel) it
    // doesn't eat into / clip the last row. The outer QScrollArea owns vertical
    // scrolling; this table never nests a vertical scrollbar.
    const int frame = 2 * table->frameWidth();
    const int hbar  = table->horizontalScrollBar()->sizeHint().height();
    table->setFixedHeight(table->horizontalHeader()->height() + nRows * rowH + frame + hbar);
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

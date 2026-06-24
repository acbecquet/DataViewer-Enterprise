#include "NotesStoryPanel.h"
#include "../utils/AppTheme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QScrollArea>
#include <QScrollBar>
#include <QFrame>
#include <QLabel>
#include <QPlainTextEdit>
#include <QSpinBox>
#include <QComboBox>
#include <QTimer>
#include <QToolButton>
#include <QGridLayout>

namespace DVE {

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
    setMinimumWidth(280);
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
    auto* chips = new QHBoxLayout; chips->setSpacing(6);
    auto roChip = [](const QString& t){ auto* l=new QLabel(t); l->setStyleSheet(
        QStringLiteral("background:#F0F0F0;color:#555;border-radius:6px;padding:2px 6px;font-size:8pt;")); return l; };
    chips->addWidget(roChip(QString("TPM before %1").arg(neighbourTpm(-1))));
    chips->addWidget(roChip(QString("TPM after %1").arg(neighbourTpm(+1))));
    chips->addWidget(roChip(QString("Draw %1").arg(dr.drawPressure,0,'f',1)));
    chips->addStretch(1);
    v->addLayout(chips);

    auto* edits = new QHBoxLayout; edits->setSpacing(6);
    edits->addWidget(new QLabel(QStringLiteral("Clog")));
    auto* clog = new QComboBox; clog->addItems({QStringLiteral("N"), QStringLiteral("Y")});
    clog->setCurrentText(dr.clog.trimmed().toUpper() == QLatin1String("Y") ? QStringLiteral("Y") : QStringLiteral("N"));
    connect(clog, &QComboBox::currentTextChanged, this,
        [this, rowIndex](const QString& t){ emit cellEdited(rowIndex, Cols::CLOG, t); });
    edits->addWidget(clog);
    auto* smell = new QSpinBox; smell->setRange(0,4); smell->setValue(dr.smell.toInt());
    smell->setPrefix(QStringLiteral("Smell "));
    connect(smell, qOverload<int>(&QSpinBox::valueChanged), this,
        [this, rowIndex](int val){ emit cellEdited(rowIndex, Cols::SMELL, QString::number(val)); });
    edits->addWidget(smell);
    if (perRowRegime) {
        edits->addWidget(new QLabel(QStringLiteral("Regime")));
        auto* regime = new QComboBox; regime->setEditable(true);
        regime->setCurrentText(dr.puffingRegime);    // before connecting
        auto* regimeTimer = new QTimer(regime);
        regimeTimer->setSingleShot(true); regimeTimer->setInterval(700);
        connect(regime, &QComboBox::currentTextChanged, regimeTimer, qOverload<>(&QTimer::start));
        connect(regimeTimer, &QTimer::timeout, this, [this, rowIndex, regime]() {
            emit cellEdited(rowIndex, Cols::RESISTANCE, regime->currentText());
        });
        edits->addWidget(regime);
    }
    edits->addStretch(1);
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

    auto* gridHost = new QWidget;
    gridHost->setVisible(false);
    auto* grid = new QGridLayout(gridHost);
    grid->setContentsMargins(6,2,2,4); grid->setHorizontalSpacing(5); grid->setVerticalSpacing(3);
    const QStringList heads = {QStringLiteral("Puff"),QStringLiteral("TPM"),QStringLiteral("Draw"),QStringLiteral("Smell"),QStringLiteral("Clog")};
    for (int c = 0; c < heads.size(); ++c) {
        auto* h = new QLabel(heads[c]); h->setStyleSheet(QStringLiteral("color:#999;font-size:7pt;"));
        grid->addWidget(h, 0, c);
    }
    int gr = 1;
    for (int idx : sum.rowIndices) {
        const DataRow& dr = s.rows[idx];
        const bool excl = m_excluded.contains(idx);
        const QString lblStyle = excl ? QStringLiteral("color:#aaa;font-size:8pt;") : QStringLiteral("font-size:8pt;");
        auto mk = [&](const QString& t){ auto* l = new QLabel(t); l->setStyleSheet(lblStyle); return l; };
        grid->addWidget(mk(QString::number(int(dr.puffs))), gr, 0);
        grid->addWidget(mk(QString::number(dr.tpm,'f',2)),  gr, 1);
        grid->addWidget(mk(QString::number(dr.drawPressure,'f',1)), gr, 2);
        auto* sm = new QSpinBox; sm->setRange(0,4); sm->setValue(dr.smell.toInt());   // value before connect
        sm->setMaximumWidth(48); sm->setStyleSheet(QStringLiteral("font-size:8pt;"));
        connect(sm, qOverload<int>(&QSpinBox::valueChanged), this,
            [this, idx](int v){ emit cellEdited(idx, Cols::SMELL, QString::number(v)); });
        grid->addWidget(sm, gr, 3);
        auto* cl = new QComboBox; cl->addItems({QStringLiteral("N"), QStringLiteral("Y")});
        cl->setCurrentText(dr.clog.trimmed().toUpper() == QLatin1String("Y") ? QStringLiteral("Y") : QStringLiteral("N")); // before connect
        cl->setMaximumWidth(46); cl->setStyleSheet(QStringLiteral("font-size:8pt;"));
        connect(cl, &QComboBox::currentTextChanged, this,
            [this, idx](const QString& t){ emit cellEdited(idx, Cols::CLOG, t); });
        grid->addWidget(cl, gr, 4);
        ++gr;
    }
    grid->setColumnStretch(5, 1);
    bv->addWidget(gridHost);

    connect(btn, &QToolButton::toggled, gridHost, [btn, gridHost](bool on){
        gridHost->setVisible(on);
        btn->setArrowType(on ? Qt::DownArrow : Qt::RightArrow);
    });
    return box;
}
void NotesStoryPanel::highlightRow(int) {}

} // namespace DVE

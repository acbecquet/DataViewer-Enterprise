// Outline test harness — renders a QTextEdit with various border-style
// approaches and saves each variant as a PNG so we can see which one
// actually paints a visible 3 px border under the project's AppTheme.

#include <QApplication>
#include <QGroupBox>
#include <QVBoxLayout>
#include <QLabel>
#include <QTextEdit>
#include <QFrame>
#include <QPixmap>
#include <QDir>
#include <QCommandLineParser>
#include <QPainter>
#include <QStandardPaths>
#include <QDebug>

static QString kAppQss = QStringLiteral(
    "QWidget { background-color: #F5F6F8; color: #1F2A37; }"
    "QGroupBox { background-color: #FFFFFF; border: 1px solid #D4D8DC; "
    "border-radius: 6px; margin-top: 12px; padding: 8px 6px 6px 6px; "
    "font-size: 9pt; font-weight: 600; }"
    "QTextEdit, QPlainTextEdit { background-color: #FFFFFF; color: #1F2A37; "
    "border: 1px solid #D4D8DC; border-radius: 4px; padding: 4px; }"
    "QTextEdit:focus, QPlainTextEdit:focus { border-color: #0066CC; outline: none; }"
);

static QWidget* makeCard(const QString& label,
                         const QString& variantStyle,
                         bool useObjectName,
                         bool wrapInFrame,
                         bool stripQTextEditFrame)
{
    auto* card = new QGroupBox(label);
    card->setFixedWidth(245);
    auto* layout = new QVBoxLayout(card);
    layout->setSpacing(3);
    layout->setContentsMargins(6, 16, 6, 4);

    auto* nameLabel = new QLabel("Name: Sample 1");
    layout->addWidget(nameLabel);

    auto* puff = new QLabel("Puff length: 3.0 s");
    layout->addWidget(puff);

    layout->addWidget(new QLabel("Comments:"));
    layout->addSpacing(4);

    QTextEdit* edit = new QTextEdit;
    edit->setMinimumHeight(36);
    edit->setMaximumHeight(50);
    edit->setPlainText("Sample comment text…");
    if (useObjectName)
        edit->setObjectName("sensoryCommentsEdit");
    if (!variantStyle.isEmpty())
        edit->setStyleSheet(variantStyle);
    if (stripQTextEditFrame)
        edit->setFrameShape(QFrame::NoFrame);

    if (wrapInFrame) {
        auto* frame = new QFrame;
        frame->setObjectName("commentsFrame");
        frame->setStyleSheet(
            "QFrame#commentsFrame { border: 3px solid #A0A6AE; "
            "border-radius: 4px; background: white; }"
        );
        auto* fl = new QVBoxLayout(frame);
        fl->setContentsMargins(0, 0, 0, 0);
        fl->addWidget(edit);
        layout->addWidget(frame, 1);
    } else {
        layout->addWidget(edit, 1);
    }

    return card;
}

static void renderCard(QWidget* card, const QString& outPath)
{
    card->ensurePolished();
    card->adjustSize();
    card->resize(card->sizeHint().expandedTo(QSize(245, 240)));
    // Force a real paint into a buffer
    QPixmap pm(card->size());
    pm.fill(Qt::white);
    card->render(&pm);
    pm.save(outPath, "PNG");
    qInfo() << "wrote" << outPath << "size" << pm.size();
}

int main(int argc, char** argv)
{
    QApplication app(argc, argv);
    app.setStyleSheet(kAppQss);

    QString outDir = "outline_renders";
    QDir().mkpath(outDir);

    struct Case { QString name; QString css; bool useObjectName; bool wrap; bool stripFrame; };
    QVector<Case> cases = {
        // Variant 1: current code — type+ID selector, inline
        {"v1_id_selector",
         "QTextEdit#sensoryCommentsEdit { border: 3px solid #A0A6AE; "
         "border-radius: 4px; background: white; padding: 4px; }",
         true, false, false},

        // Variant 2: pure ID selector (no type)
        {"v2_pure_id",
         "#sensoryCommentsEdit { border: 3px solid #A0A6AE; "
         "border-radius: 4px; background: white; padding: 4px; }",
         true, false, false},

        // Variant 3: no selector at all (applies to widget itself)
        {"v3_no_selector",
         "border: 3px solid #A0A6AE; border-radius: 4px; "
         "background: white; padding: 4px;",
         false, false, false},

        // Variant 4: explicit per-edge borders (defeats border shorthand parsing bugs)
        {"v4_per_edge",
         "QTextEdit#sensoryCommentsEdit { "
         "border-top: 3px solid #A0A6AE; border-bottom: 3px solid #A0A6AE; "
         "border-left: 3px solid #A0A6AE; border-right: 3px solid #A0A6AE; "
         "border-radius: 4px; background: white; padding: 4px; }",
         true, false, false},

        // Variant 5: wrap QTextEdit in a QFrame, strip QTextEdit frame, draw on frame
        {"v5_frame_wrap", "", false, true, true},

        // Variant 6: QAbstractScrollArea selector (more specific to scroll widgets)
        {"v6_abstract_scroll",
         "QAbstractScrollArea#sensoryCommentsEdit { "
         "border: 3px solid #A0A6AE; border-radius: 4px; "
         "background: white; padding: 4px; }",
         true, false, false},
    };

    for (const Case& c : cases) {
        QWidget* card = makeCard("Sample 1 — " + c.name, c.css,
                                  c.useObjectName, c.wrap, c.stripFrame);
        renderCard(card, outDir + "/" + c.name + ".png");
        delete card;
    }

    qInfo() << "all variants rendered. inspect" << QDir(outDir).absolutePath();
    return 0;
}

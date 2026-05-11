#include <QtTest/QtTest>
#include <QStack>
#include <QSharedPointer>
#include "LayoutCommand.h"
#include "ReportLayout.h"

using DVE::ReportLayout;
using DVE::ContentSlideLayout;
using DVE::RectCommand;
using DVE::FontSizeCommand;
using DVE::SortCommand;

class TstLayoutCommand : public QObject {
    Q_OBJECT
private slots:
    void testRectCommandContentTableRoundTrip();
    void testRectCommandCoverTitleRoundTrip();
    void testRectCommandCumulativeRadarRoundTrip();
    void testFontSizeCommandContentTitleRoundTrip();
    void testFontSizeCommandDividerRoundTrip();
    void testSortCommandRoundTrip();
    void testApplyThenUndoRestoresOriginal();
    void testMultipleCommandsApplyInOrder();
    void testStackUndoRedoSequence();
};

// 1. RectCommand on content_<id> table rect — round trip apply/undo.
void TstLayoutCommand::testRectCommandContentTableRoundTrip() {
    ReportLayout l;
    ContentSlideLayout cs;
    cs.table = QRectF(0.0, 0.0, 1.0, 1.0);
    l.contentSlides["content_0"] = cs;

    const QRectF oldR = cs.table;
    const QRectF newR(2.0, 3.0, 5.0, 4.0);
    RectCommand cmd(QStringLiteral("content_0"),
                    QStringLiteral("table"),
                    oldR, newR);

    cmd.apply(l);
    QCOMPARE(l.contentSlides["content_0"].table, newR);
    cmd.undo(l);
    QCOMPARE(l.contentSlides["content_0"].table, oldR);
}

// 2. RectCommand on cover title — round trip.
void TstLayoutCommand::testRectCommandCoverTitleRoundTrip() {
    ReportLayout l;
    l.coverTitle = QRectF(0.1, 0.2, 12.0, 1.0);

    const QRectF oldR = l.coverTitle;
    const QRectF newR(0.5, 0.6, 11.0, 1.2);
    RectCommand cmd(QStringLiteral("cover"),
                    QStringLiteral("cover_title"),
                    oldR, newR);

    cmd.apply(l);
    QCOMPARE(l.coverTitle, newR);
    cmd.undo(l);
    QCOMPARE(l.coverTitle, oldR);
}

// 3. RectCommand on cumulative radar — round trip.
void TstLayoutCommand::testRectCommandCumulativeRadarRoundTrip() {
    ReportLayout l;
    l.cumulative.radar = QRectF(4.5, 2.5, 4.4, 4.4);

    const QRectF oldR = l.cumulative.radar;
    const QRectF newR(5.0, 3.0, 4.0, 4.0);
    RectCommand cmd(QStringLiteral("cumulative"),
                    QStringLiteral("radar"),
                    oldR, newR);

    cmd.apply(l);
    QCOMPARE(l.cumulative.radar, newR);
    cmd.undo(l);
    QCOMPARE(l.cumulative.radar, oldR);
}

// 4. FontSizeCommand on content_<id> title — round trip.
void TstLayoutCommand::testFontSizeCommandContentTitleRoundTrip() {
    ReportLayout l;
    ContentSlideLayout cs;
    cs.titleFontPt = 18;
    l.contentSlides["content_0"] = cs;

    FontSizeCommand cmd(QStringLiteral("content_0"),
                        QStringLiteral("title"),
                        18, 24);

    cmd.apply(l);
    QCOMPARE(l.contentSlides["content_0"].titleFontPt, 24);
    cmd.undo(l);
    QCOMPARE(l.contentSlides["content_0"].titleFontPt, 18);
}

// 5. FontSizeCommand on divider title — round trip.
//    Divider font pts live in ReportLayout::dividerTitleFontPts (QHash).
void TstLayoutCommand::testFontSizeCommandDividerRoundTrip() {
    ReportLayout l;
    l.dividerTitleFontPts["divider_7"] = 32;

    FontSizeCommand cmd(QStringLiteral("divider_7"),
                        QStringLiteral("divider_title"),
                        32, 40);

    cmd.apply(l);
    QCOMPARE(l.dividerTitleFontPts["divider_7"], 40);
    cmd.undo(l);
    QCOMPARE(l.dividerTitleFontPts["divider_7"], 32);
}

// 6. SortCommand — round trip apply/undo.
//    Note: SortCommand has 4 params (oldCol, oldOrd, newCol, newOrd).
//    TableSort has no 'enabled' field; an empty column means insertion order.
void TstLayoutCommand::testSortCommandRoundTrip() {
    ReportLayout l;
    l.tableSort.column = QStringLiteral("Overall Liking");
    l.tableSort.order  = Qt::DescendingOrder;

    SortCommand cmd(QStringLiteral("Overall Liking"), Qt::DescendingOrder,
                    QStringLiteral("Smoothness"),     Qt::AscendingOrder);

    cmd.apply(l);
    QCOMPARE(l.tableSort.column, QStringLiteral("Smoothness"));
    QCOMPARE(l.tableSort.order,  Qt::AscendingOrder);
    cmd.undo(l);
    QCOMPARE(l.tableSort.column, QStringLiteral("Overall Liking"));
    QCOMPARE(l.tableSort.order,  Qt::DescendingOrder);
}

// 7. Cross-type apply-then-undo restores the original layout exactly.
void TstLayoutCommand::testApplyThenUndoRestoresOriginal() {
    ReportLayout l;
    ContentSlideLayout cs;
    cs.table       = QRectF(0.0, 0.0, 10.0, 5.0);
    cs.titleFontPt = 18;
    l.contentSlides["content_0"] = cs;
    l.tableSort.column = QStringLiteral("Overall Liking");
    l.tableSort.order  = Qt::DescendingOrder;

    RectCommand rc(QStringLiteral("content_0"),
                   QStringLiteral("table"),
                   cs.table, QRectF(1.0, 1.0, 9.0, 4.0));
    FontSizeCommand fc(QStringLiteral("content_0"),
                       QStringLiteral("title"),
                       18, 28);
    SortCommand sc(QStringLiteral("Overall Liking"), Qt::DescendingOrder,
                   QStringLiteral("Smoothness"),     Qt::AscendingOrder);

    rc.apply(l); fc.apply(l); sc.apply(l);
    // Reverse order undo, as a stack would.
    sc.undo(l);  fc.undo(l);  rc.undo(l);

    QCOMPARE(l.contentSlides["content_0"].table,       cs.table);
    QCOMPARE(l.contentSlides["content_0"].titleFontPt, 18);
    QCOMPARE(l.tableSort.column, QStringLiteral("Overall Liking"));
    QCOMPARE(l.tableSort.order,  Qt::DescendingOrder);
}

// 8. Multiple commands applied in order each take effect.
void TstLayoutCommand::testMultipleCommandsApplyInOrder() {
    ReportLayout l;
    ContentSlideLayout cs;
    cs.table       = QRectF(0.0, 0.0, 1.0, 1.0);
    cs.titleFontPt = 12;
    l.contentSlides["content_0"] = cs;

    QStack<QSharedPointer<DVE::LayoutCommand>> stack;
    stack.push(QSharedPointer<DVE::LayoutCommand>(
        new RectCommand(QStringLiteral("content_0"),
                        QStringLiteral("table"),
                        QRectF(0.0, 0.0, 1.0, 1.0),
                        QRectF(2.0, 2.0, 3.0, 3.0))));
    stack.push(QSharedPointer<DVE::LayoutCommand>(
        new FontSizeCommand(QStringLiteral("content_0"),
                            QStringLiteral("title"),
                            12, 20)));
    stack.push(QSharedPointer<DVE::LayoutCommand>(
        new SortCommand(QString(), Qt::DescendingOrder,
                        QStringLiteral("Smoothness"), Qt::AscendingOrder)));

    for (const auto& c : stack) c->apply(l);

    QCOMPARE(l.contentSlides["content_0"].table,
             QRectF(2.0, 2.0, 3.0, 3.0));
    QCOMPARE(l.contentSlides["content_0"].titleFontPt, 20);
    QCOMPARE(l.tableSort.column, QStringLiteral("Smoothness"));
    QCOMPARE(l.tableSort.order,  Qt::AscendingOrder);
}

// 9. Simulate the dialog's QStack-based push/undo/redo flow: push 3
//    commands, undo 2, redo 1, verify state at each step matches the
//    expected snapshot.
void TstLayoutCommand::testStackUndoRedoSequence() {
    ReportLayout l;
    ContentSlideLayout cs;
    cs.table = QRectF(0, 0, 1, 1);
    l.contentSlides[QStringLiteral("content_0")] = cs;

    QStack<QSharedPointer<DVE::LayoutCommand>> undo;
    QStack<QSharedPointer<DVE::LayoutCommand>> redo;

    auto push = [&](QSharedPointer<DVE::LayoutCommand> cmd) {
        cmd->apply(l);
        undo.push(cmd);
        redo.clear();
    };
    auto doUndo = [&]() {
        if (undo.isEmpty()) return;
        auto cmd = undo.pop();
        cmd->undo(l);
        redo.push(cmd);
    };
    auto doRedo = [&]() {
        if (redo.isEmpty()) return;
        auto cmd = redo.pop();
        cmd->apply(l);
        undo.push(cmd);
    };

    push(QSharedPointer<DVE::LayoutCommand>(
        new RectCommand(QStringLiteral("content_0"),
                        QStringLiteral("table"),
                        QRectF(0, 0, 1, 1), QRectF(2, 2, 2, 2))));
    push(QSharedPointer<DVE::LayoutCommand>(
        new RectCommand(QStringLiteral("content_0"),
                        QStringLiteral("table"),
                        QRectF(2, 2, 2, 2), QRectF(3, 3, 3, 3))));
    push(QSharedPointer<DVE::LayoutCommand>(
        new RectCommand(QStringLiteral("content_0"),
                        QStringLiteral("table"),
                        QRectF(3, 3, 3, 3), QRectF(4, 4, 4, 4))));

    // After 3 pushes, table is at (4,4,4,4)
    QCOMPARE(l.contentSlides[QStringLiteral("content_0")].table,
             QRectF(4, 4, 4, 4));
    QCOMPARE(undo.size(), 3);
    QCOMPARE(redo.size(), 0);

    // Undo 1: back to (3,3,3,3)
    doUndo();
    QCOMPARE(l.contentSlides[QStringLiteral("content_0")].table,
             QRectF(3, 3, 3, 3));
    QCOMPARE(undo.size(), 2);
    QCOMPARE(redo.size(), 1);

    // Undo 2: back to (2,2,2,2)
    doUndo();
    QCOMPARE(l.contentSlides[QStringLiteral("content_0")].table,
             QRectF(2, 2, 2, 2));
    QCOMPARE(undo.size(), 1);
    QCOMPARE(redo.size(), 2);

    // Redo 1: forward to (3,3,3,3)
    doRedo();
    QCOMPARE(l.contentSlides[QStringLiteral("content_0")].table,
             QRectF(3, 3, 3, 3));
    QCOMPARE(undo.size(), 2);
    QCOMPARE(redo.size(), 1);

    // A new push clears the redo stack
    push(QSharedPointer<DVE::LayoutCommand>(
        new RectCommand(QStringLiteral("content_0"),
                        QStringLiteral("table"),
                        QRectF(3, 3, 3, 3), QRectF(9, 9, 9, 9))));
    QCOMPARE(l.contentSlides[QStringLiteral("content_0")].table,
             QRectF(9, 9, 9, 9));
    QCOMPARE(undo.size(), 3);
    QCOMPARE(redo.size(), 0);
}

QTEST_MAIN(TstLayoutCommand)
#include "tst_layoutcommand.moc"

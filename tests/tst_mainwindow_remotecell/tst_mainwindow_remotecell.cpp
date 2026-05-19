#include <QApplication>
#include <QtTest>
#include <QTableWidget>
#include <QTableWidgetItem>

#include "../../src/RemoteCellHelpers.h"

using namespace DVE;

// C6 / H6 — covers the small testable surface of the rowChanged
// consolidation work. The full MainWindow consolidation (single
// rowChanged subscriber, m_applyingRemote echo guard, [this] captures
// for M8) lands in the same commit but is verified by code review and
// the existing two-client e2e test.
class TstMainWindowRemoteCell : public QObject
{
    Q_OBJECT
private slots:
    void applyRemoteValueToCell_skipsDirtyCell();
    void applyRemoteValueToCell_appliesToCleanCell();
    void applyRemoteValueToCell_nullItemReturnsFalse();
};

void TstMainWindowRemoteCell::applyRemoteValueToCell_skipsDirtyCell()
{
    QTableWidget table(1, 1);
    auto* it = new QTableWidgetItem("old");
    it->setData(Qt::UserRole + 2, QStringLiteral("old"));   // baseline
    it->setData(Qt::UserRole + 3, true);                    // dirty
    table.setItem(0, 0, it);

    const bool applied = applyRemoteValueToCell(&table, 0, 0, QStringLiteral("remote"));
    QVERIFY(!applied);                                      // skipped
    QCOMPARE(it->text(), QStringLiteral("old"));            // unchanged
    QCOMPARE(it->data(Qt::UserRole + 2).toString(),
             QStringLiteral("old"));                        // baseline unchanged
    QCOMPARE(it->data(Qt::UserRole + 3).toBool(), true);    // still dirty
}

void TstMainWindowRemoteCell::applyRemoteValueToCell_appliesToCleanCell()
{
    QTableWidget table(1, 1);
    auto* it = new QTableWidgetItem("old");
    it->setData(Qt::UserRole + 2, QStringLiteral("old"));
    it->setData(Qt::UserRole + 3, false);                   // clean
    table.setItem(0, 0, it);

    const bool applied = applyRemoteValueToCell(&table, 0, 0, QStringLiteral("remote"));
    QVERIFY(applied);
    QCOMPARE(it->text(), QStringLiteral("remote"));
    QCOMPARE(it->data(Qt::UserRole + 2).toString(),
             QStringLiteral("remote"));                     // baseline updated
    QCOMPARE(it->data(Qt::UserRole + 3).toBool(), false);   // still clean
}

void TstMainWindowRemoteCell::applyRemoteValueToCell_nullItemReturnsFalse()
{
    QTableWidget table(1, 1);
    // no setItem — table.item(0, 0) returns nullptr.
    const bool applied = applyRemoteValueToCell(&table, 0, 0, QStringLiteral("remote"));
    QVERIFY(!applied);
}

QTEST_MAIN(TstMainWindowRemoteCell)
#include "tst_mainwindow_remotecell.moc"

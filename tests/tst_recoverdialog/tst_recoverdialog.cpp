#include <QtTest/QtTest>

#include <QListWidget>
#include <QListWidgetItem>

#include "ui/RecoverDialog.h"
#include "utils/RecoveryManager.h"

using namespace DVE;

// Unit test for RecoverDialog::selected() (Plan C C9 selective-reload dialog).
//
// RecoverDialog shows one checkable QListWidget row per RecoveryEntry, all
// checked by default; selected() returns the entries whose row is still checked,
// in their original order. This guards the "restore the RIGHT session" invariant:
// the checked-row -> entry mapping must stay 1:1 and order-preserving, so a user
// who unticks one item gets exactly the rest back -- never a shifted/wrong item.
//
// The dialog keeps its QListWidget private (m_list), so the test reaches it via
// findChild<QListWidget*>() and toggles item check states directly -- the same
// surface a real click drives. No modal exec(); selected() is callable without
// show()/exec() because the rows exist as soon as the dialog is constructed.
class TstRecoverDialog : public QObject
{
    Q_OBJECT

private:
    // Three distinct entries (distinct kind / id / displayName / payload) so a
    // wrong mapping is unambiguous in a failure diff.
    static QVector<RecoveryEntry> makeThreeEntries()
    {
        RecoveryEntry a;
        a.kind        = RecoveryKind::Tpm;
        a.id          = QStringLiteral("C:/data/Alpha.xlsx");
        a.displayName = QStringLiteral("Alpha.xlsx");
        a.sourcePath  = QStringLiteral("C:/data/Alpha.xlsx");
        a.dirty       = true;
        a.blobFile    = QStringLiteral("tpm_alpha.json");
        a.payload     = QJsonObject{ { "marker", QStringLiteral("ALPHA") } };

        RecoveryEntry b;
        b.kind        = RecoveryKind::Sensory;
        b.id          = QStringLiteral("session-beta");
        b.displayName = QStringLiteral("Beta Panel");
        b.dirty       = false;
        b.blobFile    = QStringLiteral("sensory_beta.json");
        b.payload     = QJsonObject{ { "marker", QStringLiteral("BETA") } };

        RecoveryEntry c;
        c.kind        = RecoveryKind::Detailed;
        c.id          = QStringLiteral("ds-gamma");
        c.displayName = QStringLiteral("Gamma Detailed");
        c.dirty       = true;
        c.blobFile    = QStringLiteral("detailed_gamma.json");
        c.payload     = QJsonObject{ { "marker", QStringLiteral("GAMMA") } };

        return { a, b, c };
    }

    // Locate the dialog's internal checkable list.
    static QListWidget* listOf(RecoverDialog& dlg)
    {
        QListWidget* list = dlg.findChild<QListWidget*>();
        return list;
    }

    // Compare two entries by the identity fields that make them distinct.
    static void compareEntry(const RecoveryEntry& got, const RecoveryEntry& want)
    {
        QCOMPARE(static_cast<int>(got.kind), static_cast<int>(want.kind));
        QCOMPARE(got.id,          want.id);
        QCOMPARE(got.displayName, want.displayName);
        QCOMPARE(got.payload,     want.payload);
    }

private slots:
    // Uncheck the MIDDLE row (index 1) -> selected() returns [0] and [2], in
    // order, and NOT [1]. This is the core "right session" guard.
    void uncheckMiddleReturnsOuterTwoInOrder()
    {
        const QVector<RecoveryEntry> entries = makeThreeEntries();
        RecoverDialog dlg(entries);

        QListWidget* list = listOf(dlg);
        QVERIFY(list != nullptr);
        QCOMPARE(list->count(), 3);

        // All rows start checked (the dialog's documented default).
        QCOMPARE(list->item(0)->checkState(), Qt::Checked);
        QCOMPARE(list->item(1)->checkState(), Qt::Checked);
        QCOMPARE(list->item(2)->checkState(), Qt::Checked);

        // Uncheck the middle one.
        list->item(1)->setCheckState(Qt::Unchecked);

        const QVector<RecoveryEntry> sel = dlg.selected();
        QCOMPARE(sel.size(), 2);
        compareEntry(sel[0], entries[0]);   // Alpha, in original position
        compareEntry(sel[1], entries[2]);   // Gamma, after the gap -> order preserved

        // The unchecked middle entry must NOT be present.
        for (const RecoveryEntry& e : sel)
            QVERIFY(e.id != entries[1].id);
    }

    // All rows checked (the default) -> selected() returns all three, in order.
    void allCheckedReturnsAllInOrder()
    {
        const QVector<RecoveryEntry> entries = makeThreeEntries();
        RecoverDialog dlg(entries);

        QListWidget* list = listOf(dlg);
        QVERIFY(list != nullptr);
        QCOMPARE(list->count(), 3);
        // No toggling: rows are checked by default.

        const QVector<RecoveryEntry> sel = dlg.selected();
        QCOMPARE(sel.size(), 3);
        compareEntry(sel[0], entries[0]);
        compareEntry(sel[1], entries[1]);
        compareEntry(sel[2], entries[2]);
    }

    // All rows unchecked -> selected() returns an empty vector.
    void allUncheckedReturnsEmpty()
    {
        const QVector<RecoveryEntry> entries = makeThreeEntries();
        RecoverDialog dlg(entries);

        QListWidget* list = listOf(dlg);
        QVERIFY(list != nullptr);
        QCOMPARE(list->count(), 3);

        for (int i = 0; i < list->count(); ++i)
            list->item(i)->setCheckState(Qt::Unchecked);

        const QVector<RecoveryEntry> sel = dlg.selected();
        QVERIFY(sel.isEmpty());
    }
};

QTEST_MAIN(TstRecoverDialog)
#include "tst_recoverdialog.moc"

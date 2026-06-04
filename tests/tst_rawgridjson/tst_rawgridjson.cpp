#include <QtTest>
#include "database/RawGridJson.h"
using namespace DVE;

class TestRawGridJson : public QObject {
    Q_OBJECT
private slots:
    void roundTrip_headersAndRows();
    void empty_serializesToEmptyString();
    void emptyString_parsesToEmpty();
    void specialChars_survive();
    void ragged_rows_survive();
};

void TestRawGridJson::roundTrip_headersAndRows() {
    const QStringList h{"A","B","C"};
    const QVector<QStringList> r{{"1","2","3"},{"x","y","z"}};
    const QString j = rawGridToJson(h, r);
    QVERIFY(!j.isEmpty());
    QStringList h2; QVector<QStringList> r2;
    rawGridFromJson(j, h2, r2);
    QCOMPARE(h2, h);
    QCOMPARE(r2, r);
}
void TestRawGridJson::empty_serializesToEmptyString() {
    QCOMPARE(rawGridToJson({}, {}), QString());
}
void TestRawGridJson::emptyString_parsesToEmpty() {
    QStringList h{"stale"}; QVector<QStringList> r{{"stale"}};
    rawGridFromJson(QString(), h, r);
    QVERIFY(h.isEmpty());
    QVERIFY(r.isEmpty());
}
void TestRawGridJson::specialChars_survive() {
    const QStringList h{"He said \"hi\""};
    const QVector<QStringList> r{{"a,b","unïcode","line\nbreak"}};
    QStringList h2; QVector<QStringList> r2;
    rawGridFromJson(rawGridToJson(h, r), h2, r2);
    QCOMPARE(h2, h);
    QCOMPARE(r2, r);
}
void TestRawGridJson::ragged_rows_survive() {
    const QStringList h{"A","B"};
    const QVector<QStringList> r{{"1"},{"1","2","3"}};
    QStringList h2; QVector<QStringList> r2;
    rawGridFromJson(rawGridToJson(h, r), h2, r2);
    QCOMPARE(r2, r);
}
QTEST_MAIN(TestRawGridJson)
#include "tst_rawgridjson.moc"

#include <QtTest>
#include "XmlBuilder.h"
#include "TestHelpers.h"

class TestXmlBuilder : public QObject
{
    Q_OBJECT

private slots:
    void testDeclaration();
    void testOpenClose();
    void testLeafWithText();
    void testSelfClose();
    void testAttributes();
    void testEscaping();
    void testNesting();
    void testRaw();
    void testClear();
    void testToUtf8();
    void escapeXml_stripsIllegalControlChars();
};

void TestXmlBuilder::testDeclaration()
{
    XmlBuilder b;
    QString xml = b.declaration().build();
    QVERIFY(xml.startsWith("<?xml"));
    QVERIFY(xml.contains("version=\"1.0\""));
    QVERIFY(xml.contains("encoding=\"UTF-8\""));
    QVERIFY(xml.contains("standalone=\"yes\""));
}

void TestXmlBuilder::testOpenClose()
{
    XmlBuilder b;
    QString xml = b.open("root").close("root").build();
    QVERIFY(xml.contains("<root>"));
    QVERIFY(xml.contains("</root>"));
}

void TestXmlBuilder::testLeafWithText()
{
    XmlBuilder b;
    QString xml = b.leaf("name", "hello").build();
    QVERIFY(xml.contains("<name>hello</name>"));
}

void TestXmlBuilder::testSelfClose()
{
    XmlBuilder b;
    QString xml = b.selfClose("br").build();
    QVERIFY(xml.contains("<br/>"));
}

void TestXmlBuilder::testAttributes()
{
    XmlBuilder b;
    QMap<QString,QString> attrs;
    attrs["href"] = "url";
    attrs["class"] = "c";
    QString xml = b.open("a", attrs).close("a").build();
    QVERIFY(xml.contains("href=\"url\""));
    QVERIFY(xml.contains("class=\"c\""));
    QVERIFY(xml.contains("<a "));
}

void TestXmlBuilder::testEscaping()
{
    XmlBuilder b;
    QString xml = b.leaf("text", "<>&\"").build();
    QVERIFY(xml.contains("&lt;"));
    QVERIFY(xml.contains("&gt;"));
    QVERIFY(xml.contains("&amp;"));
    QVERIFY(xml.contains("&quot;"));
    // Original chars should NOT appear unescaped in the text content
    // (they will appear in the tags themselves, so just check the content portion)
}

void TestXmlBuilder::testNesting()
{
    XmlBuilder b;
    QString xml = b.open("a").open("b").close("b").close("a").build();
    QVERIFY(xml.contains("<a>"));
    QVERIFY(xml.contains("<b>"));
    QVERIFY(xml.contains("</b>"));
    QVERIFY(xml.contains("</a>"));

    // Verify ordering: <a> before <b>, </b> before </a>
    int openA  = xml.indexOf("<a>");
    int openB  = xml.indexOf("<b>");
    int closeB = xml.indexOf("</b>");
    int closeA = xml.indexOf("</a>");
    QVERIFY(openA < openB);
    QVERIFY(openB < closeB);
    QVERIFY(closeB < closeA);
}

void TestXmlBuilder::testRaw()
{
    XmlBuilder b;
    QString xml = b.raw("<raw/>").build();
    QVERIFY(xml.contains("<raw/>"));
}

void TestXmlBuilder::testClear()
{
    XmlBuilder b;
    b.open("root").close("root");
    b.clear();
    QString xml = b.build();
    QVERIFY(xml.isEmpty());
}

void TestXmlBuilder::testToUtf8()
{
    XmlBuilder b;
    b.leaf("tag", "value");
    QByteArray bytes = b.toUtf8();
    QVERIFY(!bytes.isEmpty());
    // Should be valid UTF-8 — verify it round-trips
    QString roundTrip = QString::fromUtf8(bytes);
    QVERIFY(roundTrip.contains("<tag>value</tag>"));
}

// W4 (2026-08-27 smoke): C0 control chars in tester-entered text (sensory
// notes) reached the emitted PPTX verbatim - XML 1.0 forbids them even
// escaped, so the slide part became fatally unparseable and PowerPoint's
// repairer crashed on it. escapeXml must strip what no parser may accept.
void TestXmlBuilder::escapeXml_stripsIllegalControlChars()
{
    const QString in = QStringLiteral("a") + QChar(0x01)
                     + QStringLiteral("b\t\n c") + QChar(0x0B) + QChar(0x1F)
                     + QStringLiteral("d");
    const QString out = XmlBuilder::escapeXml(in);
    QCOMPARE(out, QStringLiteral("ab\t\n cd"));  // tab/newline survive, C0 controls do not
}

QTEST_APPLESS_MAIN(TestXmlBuilder)
#include "tst_xmlbuilder.moc"

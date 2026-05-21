#pragma once
#include <QString>
#include <QStringList>
#include <QMap>
#include <QByteArray>

// ---------------------------------------------------------------------------
// XmlBuilder – fast XML string builder using string concatenation.
// Intentionally avoids Qt's DOM/SAX API for performance when generating
// large amounts of OOXML.
//
// Typical usage:
//   XmlBuilder b;
//   b.declaration()
//    .open("root", {{"xmlns", "http://..."}})
//      .leaf("child", "text", {{"attr", "val"}})
//    .close("root");
//   QByteArray bytes = b.toUtf8();
// ---------------------------------------------------------------------------
class XmlBuilder {
public:
    // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    XmlBuilder& declaration();

    // <tag attr1="v1" ...>
    XmlBuilder& open(const QString& tag,
                     const QMap<QString, QString>& attrs = {});

    // </tag>
    XmlBuilder& close(const QString& tag);

    // <tag attr1="v1" ...>text</tag>  (text is XML-escaped)
    XmlBuilder& leaf(const QString& tag,
                     const QString& text,
                     const QMap<QString, QString>& attrs = {});

    // <tag attr1="v1" .../>  (self-closing, no text)
    XmlBuilder& selfClose(const QString& tag,
                          const QMap<QString, QString>& attrs = {});

    // Inject raw pre-built XML verbatim (no escaping).
    XmlBuilder& raw(const QString& xml);

    QString    build()   const;
    QByteArray toUtf8()  const;
    void       clear();

    // Public so external XML emitters (e.g. PptxWriter shape builders that
    // don't use this builder) can share the canonical 5-entity escape.
    static QString escapeXml(const QString& s);

private:
    QStringList m_parts;

    static QString attrsToString(const QMap<QString, QString>& attrs);
};

#include "XmlBuilder.h"

// ---------------------------------------------------------------------------
XmlBuilder& XmlBuilder::declaration()
{
    m_parts.append(
        QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"));
    return *this;
}

// ---------------------------------------------------------------------------
XmlBuilder& XmlBuilder::open(const QString& tag,
                              const QMap<QString, QString>& attrs)
{
    m_parts.append(QLatin1Char('<') + tag + attrsToString(attrs) + QLatin1Char('>'));
    return *this;
}

// ---------------------------------------------------------------------------
XmlBuilder& XmlBuilder::close(const QString& tag)
{
    m_parts.append(QStringLiteral("</") + tag + QLatin1Char('>'));
    return *this;
}

// ---------------------------------------------------------------------------
XmlBuilder& XmlBuilder::leaf(const QString& tag,
                              const QString& text,
                              const QMap<QString, QString>& attrs)
{
    m_parts.append(QLatin1Char('<') + tag + attrsToString(attrs) + QLatin1Char('>') +
                   escapeXml(text) +
                   QStringLiteral("</") + tag + QLatin1Char('>'));
    return *this;
}

// ---------------------------------------------------------------------------
XmlBuilder& XmlBuilder::selfClose(const QString& tag,
                                  const QMap<QString, QString>& attrs)
{
    m_parts.append(QLatin1Char('<') + tag + attrsToString(attrs) +
                   QStringLiteral("/>"));
    return *this;
}

// ---------------------------------------------------------------------------
XmlBuilder& XmlBuilder::raw(const QString& xml)
{
    m_parts.append(xml);
    return *this;
}

// ---------------------------------------------------------------------------
QString XmlBuilder::build() const
{
    return m_parts.join(QString());
}

// ---------------------------------------------------------------------------
QByteArray XmlBuilder::toUtf8() const
{
    return build().toUtf8();
}

// ---------------------------------------------------------------------------
void XmlBuilder::clear()
{
    m_parts.clear();
}

// ---------------------------------------------------------------------------
QString XmlBuilder::escapeXml(const QString& s)
{
    QString out;
    out.reserve(s.size() + 8);
    for (const QChar c : s) {
        switch (c.unicode()) {
        case '&':  out += QStringLiteral("&amp;");  break;
        case '<':  out += QStringLiteral("&lt;");   break;
        case '>':  out += QStringLiteral("&gt;");   break;
        case '"':  out += QStringLiteral("&quot;"); break;
        case '\'': out += QStringLiteral("&apos;"); break;
        default:
            // W4 (2026-08-27): XML 1.0 legal chars are #x9 #xA #xD and
            // #x20-#xFFFD (minus #xFFFE/#xFFFF). Anything else - C0 controls
            // pasted into notes text, most commonly - is fatal to EVERY
            // conforming parser even when "escaped", and PowerPoint's repairer
            // crashes on the resulting slide part. Strip rather than emit.
            // Surrogate halves fall inside the range on purpose: QString
            // stores non-BMP text as adjacent valid pairs, and dropping a
            // half would corrupt them.
            if (const char16_t u = c.unicode();
                u == 0x9 || u == 0xA || u == 0xD || (u >= 0x20 && u <= 0xFFFD)) {
                out += c;
            }
            break;
        }
    }
    return out;
}

// ---------------------------------------------------------------------------
QString XmlBuilder::attrsToString(const QMap<QString, QString>& attrs)
{
    if (attrs.isEmpty())
        return QString();

    QString out;
    // QMap iterates in key-sorted order – deterministic output.
    for (auto it = attrs.cbegin(); it != attrs.cend(); ++it) {
        out += QLatin1Char(' ') + it.key() +
               QStringLiteral("=\"") + escapeXml(it.value()) + QLatin1Char('"');
    }
    return out;
}

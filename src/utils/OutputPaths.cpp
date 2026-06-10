#include "utils/OutputPaths.h"

#include <QSettings>
#include <QStandardPaths>
#include <QDir>

namespace DVE {

QString OutputPaths::settingsKey(ReportMode mode)
{
    switch (mode) {
    case ReportMode::Tpm:             return QStringLiteral("output/tpmDir");
    case ReportMode::Sensory:         return QStringLiteral("output/sensoryDir");
    case ReportMode::DetailedSensory: return QStringLiteral("output/detailedDir");
    }
    return QStringLiteral("output/tpmDir");
}

QString OutputPaths::configuredDir(ReportMode mode)
{
    QSettings s(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"));
    return s.value(settingsKey(mode)).toString();
}

void OutputPaths::setConfiguredDir(ReportMode mode, const QString& dir)
{
    QSettings s(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"));
    s.setValue(settingsKey(mode), dir);
}

QString OutputPaths::documentsDir()
{
    const QString docs = QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation);
    if (!docs.isEmpty() && QDir(docs).exists())
        return docs;
    return QDir::homePath() + QStringLiteral("/Documents");
}

QString OutputPaths::firstExistingDir(const QStringList& candidates, const QString& fallback)
{
    for (const QString& c : candidates) {
        if (!c.isEmpty() && QDir(c).exists())
            return c;
    }
    return fallback;
}

QString OutputPaths::resolveDir(ReportMode mode, const QString& lastUsedDir)
{
    return firstExistingDir({ configuredDir(mode), lastUsedDir }, documentsDir());
}

QString OutputPaths::sanitize(const QString& raw)
{
    static const QString illegal = QStringLiteral("\\/:*?\"<>|");
    QString out;
    out.reserve(raw.size());
    for (const QChar ch : raw) {
        if (illegal.contains(ch)) continue;
        out.append(ch.isSpace() ? QChar('_') : ch);
    }
    while (out.contains(QStringLiteral("__")))
        out.replace(QStringLiteral("__"), QStringLiteral("_"));
    while (out.startsWith(QLatin1Char('_'))) out.remove(0, 1);
    while (out.endsWith(QLatin1Char('_')))   out.chop(1);
    return out;
}

QString OutputPaths::nextSuffixedName(const QString& name)
{
    // Find a trailing "_<digits>" run. Walk back from the end over digits; the
    // run is an iterator only if it is non-empty AND immediately preceded by an
    // underscore. "Vape_Test" has no trailing digit run, so it is left intact.
    int i = name.size();
    while (i > 0 && name.at(i - 1).isDigit())
        --i;
    const bool hasDigits = (i < name.size());
    const bool precededByUnderscore = (i > 0 && name.at(i - 1) == QLatin1Char('_'));
    if (hasDigits && precededByUnderscore) {
        const QString stem  = name.left(i - 1);   // drop the "_" too
        bool ok = false;
        // toLongLong tolerates leading zeros ("03" -> 3); the re-emitted
        // iterator is unpadded by design (documented in the header).
        const qlonglong n = name.mid(i).toLongLong(&ok);
        if (ok)
            return stem + QLatin1Char('_') + QString::number(n + 1);
    }
    return name + QStringLiteral("_1");
}

QString OutputPaths::reportFileName(const QString& base, const QString& sheet)
{
    QString b = base;
    if (b.endsWith(QStringLiteral(".pptx"), Qt::CaseInsensitive))
        b.chop(5);
    QString stem = sanitize(b);
    if (stem.isEmpty()) stem = QStringLiteral("untitled");
    if (!sheet.isEmpty()) {
        const QString s = sanitize(sheet);
        if (!s.isEmpty()) stem += QStringLiteral("_") + s;
    }
    return stem + QStringLiteral("_report.pptx");
}

QString OutputPaths::autoSavePath(ReportMode mode, const QString& sessionLabel,
                                  const QString& lastUsedDir, const QString& ext)
{
    QString dir = resolveDir(mode, lastUsedDir);
    while (dir.endsWith(QLatin1Char('/')) || dir.endsWith(QLatin1Char('\\')))
        dir.chop(1);                                  // normalize: no trailing slash
    QString base = sanitize(sessionLabel);
    if (base.trimmed().isEmpty()) base = QStringLiteral("untitled");
    QString e = ext;
    if (!e.isEmpty() && !e.startsWith(QLatin1Char('.'))) e.prepend(QLatin1Char('.'));
    return dir + QLatin1Char('/') + base + e;
}

} // namespace DVE

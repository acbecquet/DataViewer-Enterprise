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

QString OutputPaths::resolveReportDir(ReportMode mode, const QString& lastUsedDir)
{
    return firstExistingDir({ configuredDir(mode), lastUsedDir }, documentsDir());
}

QString OutputPaths::resolveSaveDir(const QString& lastUsedDir)
{
    return firstExistingDir({ lastUsedDir }, documentsDir());
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

} // namespace DVE

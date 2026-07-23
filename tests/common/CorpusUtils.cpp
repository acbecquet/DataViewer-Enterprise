#include "CorpusUtils.h"
#include <QDir>
#include <QDirIterator>
#include <QFileInfo>

namespace DVE { namespace testutil {

static QString fixturesDir()
{
    // SRCDIR is <repo>/tests/tst_<suite>; fixtures live at <repo>/tests/data.
    return QFileInfo(QStringLiteral(SRCDIR) + "/../data").absoluteFilePath();
}

QStringList corpusFiles()
{
    QStringList out;
    QDirIterator fix(fixturesDir(), {QStringLiteral("*.xlsx")}, QDir::Files);
    while (fix.hasNext()) out << fix.next();

    const QByteArray env = qgetenv("DVE_TEST_CORPUS_DIR");
    if (!env.isEmpty() && QDir(QString::fromUtf8(env)).exists()) {
        QDirIterator it(QString::fromUtf8(env), {QStringLiteral("*.xlsx")},
                        QDir::Files, QDirIterator::Subdirectories);
        while (it.hasNext()) out << it.next();
    }
    out.removeDuplicates();
    out.sort();
    return out;
}

QString corpusDirDescription()
{
    const QByteArray env = qgetenv("DVE_TEST_CORPUS_DIR");
    if (env.isEmpty() || !QDir(QString::fromUtf8(env)).exists())
        return QStringLiteral("fixtures only");
    return QString::fromUtf8(env);
}

}} // namespace DVE::testutil

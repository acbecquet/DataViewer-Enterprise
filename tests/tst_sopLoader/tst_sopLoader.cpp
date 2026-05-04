#include <QtTest>
#include <QFileInfo>
#include <QDir>
#include "SopLoader.h"

class TestSopLoader : public QObject
{
    Q_OBJECT

private slots:
    void loadsKnownTemplate();
    void missingFileReturnsEmpty();

private:
    QString findTemplateDir() const
    {
        // Walk up from executable directory looking for resources/templates
        QDir d(QCoreApplication::applicationDirPath());
        for (int i = 0; i < 6; ++i) {
            QString candidate = d.absoluteFilePath("resources/templates");
            if (QDir(candidate).exists()) return candidate;
            candidate = d.absoluteFilePath("../resources/templates");
            if (QDir(candidate).exists()) return candidate;
            d.cdUp();
        }
        return QString();
    }
};

void TestSopLoader::loadsKnownTemplate()
{
    const QString templateDir = findTemplateDir();
    QVERIFY2(!templateDir.isEmpty(), "Could not find resources/templates directory");

    const QString xlsx = templateDir + "/Standardized Test Template - December 2025.xlsx";
    QVERIFY2(QFileInfo::exists(xlsx), qPrintable("Template not found: " + xlsx));

    const QVector<DVE::SopEntry> rows = DVE::SopLoader::load(xlsx);
    QVERIFY2(rows.size() >= 3, qPrintable(QString("expected >= 3 SOP rows, got %1").arg(rows.size())));

    bool hasLifetime = false;
    for (const auto& r : rows) {
        if (r.test.compare("Lifetime Test", Qt::CaseInsensitive) == 0) {
            hasLifetime = true;
            QVERIFY2(!r.objective.isEmpty(), "Lifetime Test row has empty Objective");
            QVERIFY2(!r.passCriteria.isEmpty(), "Lifetime Test row has empty Pass Criteria");
        }
    }
    QVERIFY2(hasLifetime, "no row matched 'Lifetime Test'");
}

void TestSopLoader::missingFileReturnsEmpty()
{
    const QVector<DVE::SopEntry> rows = DVE::SopLoader::load("Z:/does/not/exist.xlsx");
    QCOMPARE(rows.size(), 0);
}

QTEST_GUILESS_MAIN(TestSopLoader)
#include "tst_sopLoader.moc"

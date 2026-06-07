#include <QtTest>
#include <QTemporaryDir>
#include <QStandardPaths>
#include <QSettings>
#include "utils/OutputPaths.h"

using namespace DVE;

class TestOutputPaths : public QObject
{
    Q_OBJECT
private slots:
    void initTestCase();
    void cleanup();
    void sanitize_strips_illegal();
    void sanitize_collapses_and_trims();
    void reportFileName_base_only();
    void reportFileName_with_sheet();
    void reportFileName_strips_existing_ext();
    void reportFileName_empty_base_fallback();
    void firstExistingDir_picks_first();
    void firstExistingDir_skips_missing_and_empty();
    void firstExistingDir_returns_fallback();
    void resolveDir_prefers_configured();
    void resolveDir_falls_to_lastused_when_unset();
    void resolveDir_no_configured_uses_lastused_then_documents();
    void documentsDir_nonempty();
    void autoSavePath_joinsResolvedDirSanitizedLabelAndExt();
};

void TestOutputPaths::initTestCase()
{
    QStandardPaths::setTestModeEnabled(true);
    QSettings::setDefaultFormat(QSettings::IniFormat);
}

void TestOutputPaths::cleanup()
{
    OutputPaths::setConfiguredDir(ReportMode::Tpm, QString());
    OutputPaths::setConfiguredDir(ReportMode::Sensory, QString());
    OutputPaths::setConfiguredDir(ReportMode::DetailedSensory, QString());
}

void TestOutputPaths::sanitize_strips_illegal()
{
    QCOMPARE(OutputPaths::sanitize(QStringLiteral("a/b\\c:d*e?f\"g<h>i|j")),
             QStringLiteral("abcdefghij"));
}

void TestOutputPaths::sanitize_collapses_and_trims()
{
    QCOMPARE(OutputPaths::sanitize(QStringLiteral("  Foo   Bar  ")), QStringLiteral("Foo_Bar"));
    QCOMPARE(OutputPaths::sanitize(QStringLiteral("__x__")), QStringLiteral("x"));
}

void TestOutputPaths::reportFileName_base_only()
{
    QCOMPARE(OutputPaths::reportFileName(QStringLiteral("Foo")), QStringLiteral("Foo_report.pptx"));
}

void TestOutputPaths::reportFileName_with_sheet()
{
    QCOMPARE(OutputPaths::reportFileName(QStringLiteral("Foo"), QStringLiteral("Sheet 1")),
             QStringLiteral("Foo_Sheet_1_report.pptx"));
}

void TestOutputPaths::reportFileName_strips_existing_ext()
{
    QCOMPARE(OutputPaths::reportFileName(QStringLiteral("Foo.pptx")), QStringLiteral("Foo_report.pptx"));
}

void TestOutputPaths::reportFileName_empty_base_fallback()
{
    QCOMPARE(OutputPaths::reportFileName(QString()), QStringLiteral("untitled_report.pptx"));
}

void TestOutputPaths::firstExistingDir_picks_first()
{
    QTemporaryDir a, b;
    QCOMPARE(OutputPaths::firstExistingDir({a.path(), b.path()}, QStringLiteral("/fallback")),
             a.path());
}

void TestOutputPaths::firstExistingDir_skips_missing_and_empty()
{
    QTemporaryDir real;
    const QString missing = real.path() + QStringLiteral("/nope");
    QCOMPARE(OutputPaths::firstExistingDir({QString(), missing, real.path()},
                                           QStringLiteral("/fallback")),
             real.path());
}

void TestOutputPaths::firstExistingDir_returns_fallback()
{
    QCOMPARE(OutputPaths::firstExistingDir({QString(), QStringLiteral("/nope")},
                                           QStringLiteral("/fallback")),
             QStringLiteral("/fallback"));
}

void TestOutputPaths::resolveDir_prefers_configured()
{
    QTemporaryDir cfg, last;
    OutputPaths::setConfiguredDir(ReportMode::Tpm, cfg.path());
    QCOMPARE(OutputPaths::resolveDir(ReportMode::Tpm, last.path()), cfg.path());
}

void TestOutputPaths::resolveDir_falls_to_lastused_when_unset()
{
    QTemporaryDir last;
    OutputPaths::setConfiguredDir(ReportMode::Sensory, QString());
    QCOMPARE(OutputPaths::resolveDir(ReportMode::Sensory, last.path()), last.path());
}

void TestOutputPaths::resolveDir_no_configured_uses_lastused_then_documents()
{
    QTemporaryDir last;
    OutputPaths::setConfiguredDir(ReportMode::DetailedSensory, QString());
    QCOMPARE(OutputPaths::resolveDir(ReportMode::DetailedSensory, last.path()), last.path());
    QCOMPARE(OutputPaths::resolveDir(ReportMode::DetailedSensory, QString()), OutputPaths::documentsDir());
}

void TestOutputPaths::documentsDir_nonempty()
{
    QVERIFY(!OutputPaths::documentsDir().isEmpty());
}

void TestOutputPaths::autoSavePath_joinsResolvedDirSanitizedLabelAndExt()
{
    using namespace DVE;
    const QString dir = OutputPaths::resolveDir(ReportMode::Sensory, QString());
    const QString p = OutputPaths::autoSavePath(ReportMode::Sensory,
                                                QStringLiteral("My Test - Alice - 1"),
                                                QString(), QStringLiteral(".xlsx"));
    QVERIFY(p.startsWith(dir));
    QVERIFY(p.endsWith(QStringLiteral(".xlsx")));
    const QString expectedBase = OutputPaths::sanitize(QStringLiteral("My Test - Alice - 1"));
    QVERIFY(p.contains(expectedBase));
    // ext normalization: "xlsx" (no dot) == ".xlsx"
    QCOMPARE(OutputPaths::autoSavePath(ReportMode::Sensory, QStringLiteral("X"),
                                       QString(), QStringLiteral("xlsx")),
             OutputPaths::autoSavePath(ReportMode::Sensory, QStringLiteral("X"),
                                       QString(), QStringLiteral(".xlsx")));
    // empty/whitespace label -> non-empty base (never ".../.xlsx" or ".../ .xlsx")
    const QString empty = OutputPaths::autoSavePath(ReportMode::Sensory, QStringLiteral("   "),
                                                    QString(), QStringLiteral(".xlsx"));
    QVERIFY(!empty.endsWith(QStringLiteral("/.xlsx")));
    QVERIFY(empty.contains(QStringLiteral("untitled")));
}

QTEST_MAIN(TestOutputPaths)
#include "tst_outputpaths.moc"

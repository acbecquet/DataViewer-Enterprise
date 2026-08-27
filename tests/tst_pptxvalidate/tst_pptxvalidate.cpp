// ── DataViewer Enterprise — PPTX semantic validation harness ────────────────
//
// W3/W4 (2026-08-27 smoke): every generated report triggered PowerPoint's
// "repair" prompt, and the sensory deck crashed the repairer. The decks were
// STRUCTURALLY clean (zip, rels, content types, well-formedness all pass on a
// real production deck), so the defect class is SEMANTIC: xsd:sequence child
// order, shape-id uniqueness, table geometry, XML-illegal control characters.
// validate_pptx.py (same directory) encodes those rules for the element
// vocabulary PptxWriter emits; this suite generates representative decks and
// fails with the validator's own VIOLATION lines, which name part + element +
// rule precisely.
//
// Python resolution: plain "python" from PATH, same policy as the corpus
// harnesses — QSKIP when it cannot be started.
// ────────────────────────────────────────────────────────────────────────────
#include <QtTest>
#include <QTemporaryDir>
#include <QProcess>
#include <QDir>
#include "TestHelpers.h"
#include "ReportData.h"
#include "DataProcessor.h"
#include "ReportGenerator.h"
#include "PptxWriter.h"

class tst_PptxValidate : public QObject
{
    Q_OBJECT

private:
    QTemporaryDir m_tmp;
    QString m_validator;      // resolved validate_pptx.py path
    QString m_resourceDir;    // repo resources/ dir (logos), may be empty
    bool m_pythonOk = false;

    // Walk up from the executable dir looking for resources/images/.
    static QString findResourceDir()
    {
        QDir d(QCoreApplication::applicationDirPath());
        for (int i = 0; i < 6; ++i) {
            const QString candidate = d.absoluteFilePath("resources");
            if (QFile::exists(candidate + "/images/Cover_Page_Logo.jpg"))
                return candidate;
            d.cdUp();
        }
        return QString();
    }

    // Run the validator over the given decks; QFAIL with its output on nonzero
    // exit. Returns false when python is unavailable (caller QSKIPs).
    bool runValidator(const QStringList& decks, QString* failMsg)
    {
        QProcess p;
        p.setProcessChannelMode(QProcess::MergedChannels);
        p.start(QStringLiteral("python"), QStringList{m_validator} + decks);
        if (!p.waitForStarted(10000))
            return false;
        p.waitForFinished(120000);
        const QString out = QString::fromUtf8(p.readAll());
        if (p.exitStatus() != QProcess::NormalExit || p.exitCode() != 0) {
            *failMsg = QStringLiteral("validator exit %1:\n%2")
                           .arg(p.exitCode()).arg(out);
            return true;   // started fine; the FAILURE carries the violations
        }
        failMsg->clear();
        return true;
    }

private slots:
    void initTestCase()
    {
        QVERIFY(m_tmp.isValid());
        m_validator = QFINDTESTDATA("validate_pptx.py");
        QVERIFY2(!m_validator.isEmpty(), "validate_pptx.py not found next to the test source");
        m_resourceDir = findResourceDir();

        QProcess probe;
        probe.start(QStringLiteral("python"), {QStringLiteral("--version")});
        m_pythonOk = probe.waitForStarted(10000) && probe.waitForFinished(10000);
    }

    // Deck built straight through PptxWriter: cover slide (real logos when the
    // repo resources are found) + a content slide whose table carries
    // multi-line text, all five XML-special characters, and a C0 control char
    // (the sensory-crash class — tester notes can contain anything).
    void pptxWriterDeck_passesSemanticValidation()
    {
        if (!m_pythonOk) QSKIP("python not on PATH");

        DVE::PptxWriter w;
        if (!m_resourceDir.isEmpty()) {
            w.setResourcePath(m_resourceDir);
            w.addCoverSlide(QStringLiteral("Validator Deck"),
                            QStringLiteral("2026-08-27"));
        }

        DVE::SlideTable t;
        t.headers = { QStringLiteral("Puffs"),
                      QStringLiteral("Notes & <specials>") };
        t.rows = {
            { QStringLiteral("10"),
              QStringLiteral("line1\nline2 \"quoted\" '&'") },
            { QStringLiteral("20"),
              QStringLiteral("ctrl:") + QChar(0x01) + QStringLiteral(":char") },
        };
        w.addContentSlide(QStringLiteral("Sheet A"), t, {});

        const QString deck = m_tmp.path() + QStringLiteral("/pptxwriter_deck.pptx");
        QVERIFY2(w.save(deck), qPrintable(w.lastError()));

        QString fail;
        if (!runValidator({deck}, &fail)) QSKIP("python not on PATH");
        QVERIFY2(fail.isEmpty(), qPrintable(fail));
    }

    // Integration-level deck through the real generator + fixture pipeline.
    void generatorFullReport_passesSemanticValidation()
    {
        if (!m_pythonOk) QSKIP("python not on PATH");

        DVE::DataProcessor dp;
        DVE::FileResult formatE = dp.processFile(testDataFile("format_e.xlsx"));
        if (formatE.filePath.isEmpty())
            QSKIP("Python/openpyxl unavailable for the fixture parse");

        DVE::ReportGenerator gen;
        if (!m_resourceDir.isEmpty())
            gen.setResourcePath(m_resourceDir);
        DVE::ReportConfig config;
        config.outputPath    = m_tmp.path() + QStringLiteral("/full_report.pptx");
        config.reportTitle   = QStringLiteral("Validator Full Report");
        config.includePlots  = false;   // plot rendering deadlocks headless; tst_plotengine owns it
        config.includeImages = false;
        QVERIFY2(gen.generateFullReport(formatE, config), qPrintable(gen.lastError()));

        QString fail;
        if (!runValidator({config.outputPath}, &fail)) QSKIP("python not on PATH");
        QVERIFY2(fail.isEmpty(), qPrintable(fail));
    }
};

QTEST_MAIN(tst_PptxValidate)
#include "tst_pptxvalidate.moc"

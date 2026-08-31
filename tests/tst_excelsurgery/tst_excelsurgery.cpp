// ── DataViewer Enterprise — Excel write-back surgery invariants ─────────────
//
// W1 (2026-08-27 smoke): the openpyxl load+save write-back rewrote the ENTIRE
// workbook package, stripping every cached formula value (openpyxl never
// computes) and dropping parts it does not model (xl/metadata.xml,
// webextensions, calcChain). Excel then demanded repair, and the app's own
// data_only reader saw None for every formula cell on the next open — the
// SheetProcessors heuristic fabricated puff chains from the wreck and the
// garbage reached the database.
//
// These tests drive the REAL scripts (DVE::excelWriteCellsScript /
// excelDeleteRowScript, the exact bytes production runs) against a synthetic
// formula-rich fixture with real cached values, then assert the invariants the
// surgical rewrite must hold:
//   * untouched formula cells keep BOTH <f> and their cached <v>
//   * foreign parts survive byte-identical
//   * calcChain is dropped WITH its rels + content-type references
//   * fullCalcOnLoad is set (dependent caches are stale after an edit)
//   * a bad input file is never half-written (atomic tmp+replace tail)
//
// The fixture is assembled from raw zip parts (make_formula_fixture.py):
// openpyxl cannot WRITE cached values, which is the entire bug.
// ────────────────────────────────────────────────────────────────────────────
#include <QtTest>
#include <QProcess>
#include <QScopeGuard>
#include <QTemporaryDir>
#include <QFile>
#include "utils/ExcelWritePayload.h"
#include "ExcelReader.h"

using DVE::ExcelCellWrite;

class tst_ExcelSurgery : public QObject
{
    Q_OBJECT

private:
    QTemporaryDir m_tmp;
    QString m_fixtureScript;
    QString m_checkScript;
    bool m_pythonOk = false;
    int m_seq = 0;

    struct PyResult { int exit = -1; QString out; };

    PyResult runPy(const QStringList& args)
    {
        QProcess p;
        p.setProcessChannelMode(QProcess::MergedChannels);
        p.start(QStringLiteral("python"), args);
        if (!p.waitForStarted(10000)) return {};
        p.waitForFinished(60000);
        return { p.exitCode(), QString::fromUtf8(p.readAll()) };
    }

    // Materialize a production script's text into a temp .py and run it with
    // the given argv (the contract is "argv AFTER the script path").
    PyResult runScript(const char* scriptText, const QStringList& args)
    {
        const QString file = m_tmp.path() + QStringLiteral("/script%1.py").arg(++m_seq);
        QFile f(file);
        if (!f.open(QIODevice::WriteOnly)) return {};
        f.write(scriptText);
        f.close();
        return runPy(QStringList{file} + args);
    }

    QString makeFixture(bool stripped = false)
    {
        const QString path = m_tmp.path() + QStringLiteral("/fx%1.xlsx").arg(++m_seq);
        QStringList args{m_fixtureScript, path};
        if (stripped) args << QStringLiteral("--strip");
        const PyResult r = runPy(args);
        return (r.exit == 0 && r.out.contains(QLatin1String("OK"))) ? path : QString();
    }

private slots:
    void initTestCase()
    {
        QVERIFY(m_tmp.isValid());
        m_fixtureScript = QFINDTESTDATA("make_formula_fixture.py");
        m_checkScript   = QFINDTESTDATA("check_surgery.py");
        QVERIFY2(!m_fixtureScript.isEmpty() && !m_checkScript.isEmpty(),
                 "fixture/checker scripts not found next to the test source");
        const PyResult r = runPy({QStringLiteral("--version")});
        m_pythonOk = (r.exit == 0);
    }

    void writeCells_preservesCachesAndForeignParts()
    {
        if (!m_pythonOk) QSKIP("python not on PATH");
        const QString fx = makeFixture();
        QVERIFY2(!fx.isEmpty(), "fixture generation failed");

        const PyResult w = runScript(
            DVE::excelWriteCellsScript(),
            DVE::buildWriteCellsArgs(fx, QStringLiteral("Data"),
                                     {ExcelCellWrite{1, 1, QStringLiteral("99")}}));
        QVERIFY2(w.exit == 0 && w.out.contains(QLatin1String("OK")), qPrintable(w.out));

        const PyResult c = runPy({m_checkScript, fx});
        QVERIFY2(c.exit == 0, qPrintable(c.out));
    }

    void writeCells_textAndClearSemantics()
    {
        if (!m_pythonOk) QSKIP("python not on PATH");
        const QString fx = makeFixture();
        QVERIFY2(!fx.isEmpty(), "fixture generation failed");

        const PyResult w = runScript(
            DVE::excelWriteCellsScript(),
            DVE::buildWriteCellsArgs(fx, QStringLiteral("Data"),
                                     {ExcelCellWrite{3, 3, QStringLiteral("hello world")},
                                      ExcelCellWrite{1, 1, QString()}}));
        QVERIFY2(w.exit == 0 && w.out.contains(QLatin1String("OK")), qPrintable(w.out));

        const PyResult c = runPy({m_checkScript, fx, QStringLiteral("--textclear")});
        QVERIFY2(c.exit == 0, qPrintable(c.out));
    }

    void deleteRow_renumbersAndPreservesCaches()
    {
        if (!m_pythonOk) QSKIP("python not on PATH");
        const QString fx = makeFixture();
        QVERIFY2(!fx.isEmpty(), "fixture generation failed");

        const PyResult w = runScript(
            DVE::excelDeleteRowScript(),
            DVE::buildDeleteRowArgs(fx, QStringLiteral("Data"), 2));
        QVERIFY2(w.exit == 0 && w.out.contains(QLatin1String("OK")), qPrintable(w.out));

        const PyResult c = runPy({m_checkScript, fx, QStringLiteral("--after-delete")});
        QVERIFY2(c.exit == 0, qPrintable(c.out));
    }

    void writeCells_atomicOnBadInput()
    {
        if (!m_pythonOk) QSKIP("python not on PATH");
        const QString bad = m_tmp.path() + QStringLiteral("/bad.xlsx");
        {
            QFile f(bad);
            QVERIFY(f.open(QIODevice::WriteOnly));
            f.write("NOTAZIP789");   // 10 bytes of not-a-workbook
        }
        const PyResult w = runScript(
            DVE::excelWriteCellsScript(),
            DVE::buildWriteCellsArgs(bad, QStringLiteral("Data"),
                                     {ExcelCellWrite{1, 1, QStringLiteral("99")}}));
        QVERIFY2(w.exit != 0, "script must fail loudly on a corrupt source");
        QFile f(bad);
        QVERIFY(f.open(QIODevice::ReadOnly));
        QCOMPARE(f.readAll(), QByteArray("NOTAZIP789"));   // untouched
    }

    // Integration: the SHIPPED reader still sees correct values after a
    // surgical edit. (On Office machines the COM path computes live; on
    // openpyxl-fallback machines this additionally depends on the caches the
    // checker slots prove survive.)
    void readerSeesValuesAfterSurgicalEdit()
    {
        if (!m_pythonOk) QSKIP("python not on PATH");
        const QString fx = makeFixture();
        QVERIFY2(!fx.isEmpty(), "fixture generation failed");

        ExcelReader before;
        if (!before.loadFile(fx))
            QSKIP("reader unavailable (python/COM)");
        QVERIFY(before.selectSheet(QStringLiteral("Data")));
        QCOMPARE(before.currentSheetCells()[1][0].toDouble(), 20.0);   // A2 cache

        const PyResult w = runScript(
            DVE::excelWriteCellsScript(),
            DVE::buildWriteCellsArgs(fx, QStringLiteral("Data"),
                                     {ExcelCellWrite{1, 1, QStringLiteral("99")}}));
        QVERIFY2(w.exit == 0, qPrintable(w.out));

        ExcelReader after;
        QVERIFY(after.loadFile(fx));
        QVERIFY(after.selectSheet(QStringLiteral("Data")));
        const auto cells = after.currentSheetCells();
        QCOMPARE(cells[0][0].toDouble(), 99.0);   // the edit landed
        // A2's value must still be READABLE. COM recomputes (=99+10=109);
        // the openpyxl fallback reads the preserved cache (20). Either way it
        // must never be the empty/None the old write-back caused.
        QVERIFY2(!cells[1][0].toString().isEmpty() || cells[1][0].isValid(),
                 "A2 became unreadable after the edit - caches destroyed");
        const double a2 = cells[1][0].toDouble();
        QVERIFY2(a2 == 109.0 || a2 == 20.0,
                 qPrintable(QStringLiteral("A2 read as %1").arg(a2)));
    }

    // W1 poka-yoke: the openpyxl fallback DETECTS a cache-stripped workbook
    // (formula cells whose cached value is gone) and surfaces the count, so
    // the pipeline can refuse to fabricate values from the wreck. COM is
    // disabled via DVE_READER_NO_COM to make the fallback path deterministic
    // on Office machines.
    void reader_detectsStrippedCaches()
    {
        if (!m_pythonOk) QSKIP("python not on PATH");
        qputenv("DVE_READER_NO_COM", "1");
        const auto restore = qScopeGuard([] { qunsetenv("DVE_READER_NO_COM"); });

        const QString clean = makeFixture();
        QVERIFY2(!clean.isEmpty(), "fixture generation failed");
        ExcelReader healthy;
        if (!healthy.loadFile(clean))
            QSKIP("openpyxl unavailable");
        QVERIFY(healthy.selectSheet(QStringLiteral("Data")));
        // W3b: D2 is a healthy Excel-cached EMPTY-STRING result (t="str" with
        // an empty <v/>) - openpyxl data_only reads it as None, but the
        // workbook carries other NON-EMPTY caches (A2/A3/B2), which proves an
        // Excel save. It must NOT be counted as stripped (the 2026-08-31
        // false positive: every unfilled IF(...,"",...) helper column in the
        // owner's real workbooks flagged the file and blocked DB saves).
        QCOMPARE(healthy.currentSheetStrippedFormulas(), 0);
        QCOMPARE(healthy.currentSheetCells()[1][0].toDouble(), 20.0);  // cache read

        const QString wrecked = makeFixture(/*stripped=*/true);
        QVERIFY2(!wrecked.isEmpty(), "stripped fixture generation failed (openpyxl?)");
        ExcelReader poisoned;
        QVERIFY(poisoned.loadFile(wrecked));
        QVERIFY(poisoned.selectSheet(QStringLiteral("Data")));
        // The fixture holds 4 formula cells (A2, A3, B2, D2); openpyxl's save
        // stripped every cache, so NO formula anywhere carries a non-empty
        // value and all 4 must be flagged.
        QCOMPARE(poisoned.currentSheetStrippedFormulas(), 4);
    }
};

QTEST_MAIN(tst_ExcelSurgery)
#include "tst_excelsurgery.moc"

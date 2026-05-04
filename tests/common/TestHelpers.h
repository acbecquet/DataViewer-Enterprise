#pragma once
// ── DataViewer Enterprise — Test Helpers ──────────────────────────────
// Shared macros and utilities for all test classes.
// ──────────────────────────────────────────────────────────────────────

#include <QString>
#include <QDir>
#include <QCoreApplication>
#include <cmath>

// Resolve the test data directory relative to the test executable.
// The build puts executables in tests/tst_*/release/ or tests/tst_*/.
// Walks up the directory tree looking for tests/data/ or data/.
inline QString testDataDir()
{
    // Walk up from executable directory looking for the data folder
    QDir d(QCoreApplication::applicationDirPath());
    for (int i = 0; i < 6; ++i) {
        QString candidate = d.absoluteFilePath("tests/data");
        if (QDir(candidate).exists()) return candidate;
        candidate = d.absoluteFilePath("data");
        if (QDir(candidate).exists()) return candidate;
        // Also check ../data (common when exe is in tst_foo/release/)
        candidate = d.absoluteFilePath("../data");
        if (QDir(candidate).exists()) return candidate;
        d.cdUp();
    }
    return QStringLiteral("data");
}

inline QString testDataFile(const QString& name)
{
    return testDataDir() + "/" + name;
}

// Fuzzy double comparison with configurable tolerance.
inline bool fuzzyEqual(double a, double b, double eps = 1e-6)
{
    if (std::isnan(a) && std::isnan(b)) return true;
    if (std::isinf(a) && std::isinf(b)) return (a > 0) == (b > 0);
    return std::fabs(a - b) <= eps;
}

// QVERIFY with message for fuzzy comparison (2 or 3 args).
#define QFUZZY_COMPARE_3(actual, expected, eps) \
    QVERIFY2(fuzzyEqual((actual), (expected), (eps)), \
        qPrintable(QString("Fuzzy compare failed: actual=%1, expected=%2, eps=%3") \
            .arg(actual).arg(expected).arg(eps)))

#define QFUZZY_COMPARE_2(actual, expected) \
    QFUZZY_COMPARE_3(actual, expected, 1e-4)

// Dispatch macro: 2 args uses default eps=1e-4, 3 args uses explicit eps
#define QFUZZY_GET_MACRO(_1, _2, _3, NAME, ...) NAME
#define QFUZZY_COMPARE(...) QFUZZY_GET_MACRO(__VA_ARGS__, QFUZZY_COMPARE_3, QFUZZY_COMPARE_2)(__VA_ARGS__)

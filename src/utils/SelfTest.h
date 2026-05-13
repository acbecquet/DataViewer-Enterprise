#pragma once

#include <QString>

namespace DVE {

/// Run the deployment self-test and return a process exit code.
///
/// Exercises every code path that depends on install-tree layout, OS
/// permissions, or external resources: registry write, SQLite open, the
/// bundled Python interpreter, openpyxl Excel round-trip, the
/// Synology update-folder probe, PostgreSQL connection, and the offline
/// snapshot path/readability checks.
///
/// If outputPath is non-empty, a JSON report is written there. Returns
/// 0 if every test passed, 1 otherwise.
int runSelfTest(const QString& outputPath);

} // namespace DVE

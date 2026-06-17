#pragma once

#include <QString>

// MipFallback — shared MIP/AIP-resilience helper for the local durability
// stores (RecoveryManager JSON store, OfflineSnapshot snapshot.sqlite +
// pending_edits.sqlite).
//
// On this deployment, Microsoft Information Protection (MIP / AIP) sensitivity
// labels are applied to files at rest by the user's Windows account. A labelled
// file's raw bytes are ciphertext beginning with the magic marker
// "%TSD-Header-###%"; any plain QFile / SQLite open of such a file reads
// garbage and fails. The risk this closes: a single MIP-encrypted local file
// can silently collapse crash-recovery, offline-read, and offline-write all at
// once, because each store today treats an unreadable file as "nothing here".
//
// The bundled Python interpreter ({app}/python/python.exe) is on the MIP
// allowlist (exactly as ExcelReader already relies on to read encrypted
// workbooks): it sees PLAINTEXT. So when a store detects the marker, it shells
// out to bundled python to copy the decrypted bytes to a temp file the store
// can then open normally. If python is unavailable or the copy fails, the
// caller surfaces a LOUD warning rather than silently returning empty.
namespace DVE {

// True iff the first bytes of `path` are the MIP/AIP ciphertext marker
// "%TSD-Header-###%". Cheap: peeks ~32 bytes. Returns false on a missing /
// unreadable / plaintext file.
bool looksEncrypted(const QString& path);

// Resolve a python interpreter the same way ExcelReader / MainWindow do:
// prefer the bundled {app}/python/python.exe (which is MIP-allowlisted), then
// fall back to a system python on PATH. Returns QString() when none is found,
// in which case decryptToTempViaPython() will fail loudly. Cached internally
// after the first probe (the result can't change during a run).
QString resolveBundledPython();

// Decrypt `path` to a readable PLAINTEXT temp copy by running the
// MIP-allowlisted bundled python (`pythonExe`) to read the source bytes and
// write them to a fresh temp file. Returns the temp path on success (the caller
// then opens THAT), or QString() on failure (errOut set; caller warns loudly).
//
// The temp file is created with a stable, collision-resistant name under the
// system temp location and is NOT auto-removed (the caller keeps the QSQLITE /
// QFile handle open against it); callers may QFile::remove() it when done.
// Mirrors ExcelReader::runPythonReader's QProcess + bounded-timeout shape.
QString decryptToTempViaPython(const QString& path,
                               const QString& pythonExe,
                               QString& errOut);

} // namespace DVE

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

// True iff the first ~32 bytes of `path` contain the MIP/AIP ciphertext marker
// "%TSD-Header-###%". The scan tolerates a leading UTF-8 BOM and ASCII
// whitespace before the marker (some labellers prepend those) but a valid
// JSON / SQLite file is NEVER misdetected, so the happy path never spawns a
// spurious python decrypt. Returns false on a missing / unreadable / plaintext
// file.
bool looksEncrypted(const QString& path);

// Resolve a python interpreter the same way ExcelReader / MainWindow do:
// prefer the bundled {app}/python/python.exe (which is MIP-allowlisted), then
// fall back to a system python on PATH. Returns QString() when none is found,
// in which case the read/decrypt helpers fail loudly. Cached internally after
// the first probe (the result can't change during a run).
QString resolveBundledPython();

// CRITICAL (R5): read `path`'s PLAINTEXT bytes by running the MIP-allowlisted
// bundled python and capturing its STDOUT, base64-decoding IN MEMORY. This is
// the SAME proven shape as ExcelReader::runPythonReader: the ALLOWLISTED python
// does ALL the file I/O and hands the data back over the pipe, so the
// non-allowlisted parent never re-opens a file a MIP minifilter could re-label
// between the child write and the parent read. Use this for the small JSON
// recovery store (the highest-value durability store).
//
// Returns true + fills `out` on success; on failure (no python / process error
// / timeout / non-base64 / empty output) returns false and sets `errOut`
// (caller warns loudly). The python prints ONLY the base64 of the file bytes to
// stdout; any stderr is folded into errOut on failure.
bool readBytesViaPython(const QString& path,
                        const QString& pythonExe,
                        QByteArray& out,
                        QString& errOut);

// Decrypt `path` to a readable PLAINTEXT temp copy by running the
// MIP-allowlisted bundled python (`pythonExe`) to read the source bytes and
// write them to a fresh temp file. Returns the temp path on success (the caller
// then opens THAT), or QString() on failure (errOut set; caller warns loudly).
//
// Used ONLY for the SQLite stores (snapshot.sqlite / pending_edits.sqlite),
// where QSqlDatabase requires a real file path so a temp is unavoidable. The
// temp is python-written, which by this project's MIP convention is UNLABELED
// (the same reliance the whole MIP strategy rests on). If work-machine testing
// ever shows the SQLite temp re-labels, escalate to a python-side row export;
// do NOT build that speculatively now. The JSON store does NOT use this -- see
// readBytesViaPython above.
//
// The temp file is created with a stable, collision-resistant name under the
// system temp location and is NOT auto-removed (the caller keeps the QSQLITE
// handle open against it); callers QFile::remove() it when done. Mirrors
// ExcelReader::runPythonReader's QProcess + bounded-timeout shape.
QString decryptToTempViaPython(const QString& path,
                               const QString& pythonExe,
                               QString& errOut);

} // namespace DVE

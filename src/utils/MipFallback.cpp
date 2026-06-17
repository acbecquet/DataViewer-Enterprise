#include "MipFallback.h"

#include <QCoreApplication>
#include <QFile>
#include <QFileInfo>
#include <QDir>
#include <QStandardPaths>
#include <QTemporaryFile>
#include <QProcess>
#include <QProcessEnvironment>
#include <QDebug>

namespace DVE {

// The MIP/AIP ciphertext marker. A labelled file's raw on-disk bytes begin
// with this literal (see CLAUDE.md "MIP file encryption"). 16 bytes.
static const char kMipMarker[] = "%TSD-Header-###%";
static constexpr int kMipMarkerLen = 16;  // strlen, excludes the NUL

bool looksEncrypted(const QString& path)
{
    QFile f(path);
    if (!f.open(QIODevice::ReadOnly))
        return false;
    // R5 minor: scan the first ~32 bytes for the marker rather than matching at
    // exact offset 0. Some labellers prepend a UTF-8 BOM or stray ASCII
    // whitespace ahead of the magic; we tolerate that. We compare raw bytes (no
    // text decoding) so SQLite/JSON files are handled uniformly. Crucially a
    // valid JSON ("[", "{", whitespace) or SQLite ("SQLite format 3\0") file
    // does NOT contain "%TSD-Header-###%" in its first 32 bytes, so the happy
    // path is never misdetected and never spawns a spurious python decrypt.
    const QByteArray head = f.read(32);
    f.close();
    if (head.size() < kMipMarkerLen)
        return false;
    const QByteArray marker = QByteArray::fromRawData(kMipMarker, kMipMarkerLen);
    return head.contains(marker);
}

QString resolveBundledPython()
{
    // Cache: the interpreter location can't change during a single run, and a
    // PATH probe spawns processes -- only do it once.
    static QString cached;
    static bool probed = false;
    if (probed) return cached;
    probed = true;

    // 1. Prefer bundled python shipped next to the exe (MIP-allowlisted, the
    //    same one ExcelReader::findPython resolves first).
    const QString bundled =
        QCoreApplication::applicationDirPath() + QStringLiteral("/python/python.exe");
    if (QFile::exists(bundled)) {
        cached = bundled;
        return cached;
    }

    // 2. Fall back to a system python on PATH. (In tests there is no bundled
    //    python next to the test exe; this lets the detection + fallback wiring
    //    be exercised against a system interpreter.)
    for (const QString& exe : { QStringLiteral("python"),
                                QStringLiteral("python3"),
                                QStringLiteral("py") }) {
        QProcess p;
        p.start(exe, { QStringLiteral("--version") });
        if (p.waitForFinished(5000) && p.exitCode() == 0) {
            cached = exe;
            return cached;
        }
    }
    return cached;  // empty -> caller fails loudly
}

bool readBytesViaPython(const QString& path,
                        const QString& pythonExe,
                        QByteArray& out,
                        QString& errOut)
{
    out.clear();
    errOut.clear();

    if (pythonExe.isEmpty()) {
        errOut = QStringLiteral("readBytesViaPython: no python interpreter "
                                "available to read MIP-labelled file: ") + path;
        return false;
    }
    if (!QFile::exists(path)) {
        errOut = QStringLiteral("readBytesViaPython: source does not exist: ")
                 + path;
        return false;
    }

    // Tiny script: the ALLOWLISTED python reads the (MIP-decrypted-in-session)
    // bytes and writes their base64 to stdout. The parent never re-opens the
    // file, so a MIP minifilter can't hand it ciphertext. Path is passed as
    // argv (not interpolated) so backslashes / quotes can't break it. We write
    // base64 to STDOUT via the raw buffer to avoid any text-mode newline
    // translation, and emit nothing else on stdout.
    static const char kScript[] =
        "import sys, base64\n"
        "with open(sys.argv[1], 'rb') as f:\n"
        "    data = f.read()\n"
        "sys.stdout.buffer.write(base64.b64encode(data))\n"
        "sys.stdout.buffer.flush()\n";

    const QString tmpDir = QStandardPaths::writableLocation(
        QStandardPaths::TempLocation);
    QTemporaryFile scriptTmp(tmpDir + QStringLiteral("/dve_mip_read_XXXXXX.py"));
    scriptTmp.setAutoRemove(true);
    if (!scriptTmp.open()) {
        errOut = QStringLiteral("readBytesViaPython: cannot write script: ")
                 + scriptTmp.errorString();
        return false;
    }
    scriptTmp.write(kScript);
    scriptTmp.flush();
    const QString scriptPath = scriptTmp.fileName();
    scriptTmp.close();  // close handle so QProcess can read it on Windows

    QProcess proc;
    QProcessEnvironment env = QProcessEnvironment::systemEnvironment();
    env.insert(QStringLiteral("PYTHONIOENCODING"), QStringLiteral("utf-8"));
    env.insert(QStringLiteral("PYTHONUTF8"), QStringLiteral("1"));
    proc.setProcessEnvironment(env);
    proc.start(pythonExe, { scriptPath, path });
    // 30 s budget mirrors ExcelReader/runPython: a local read is fast; the
    // budget only guards a hung interpreter.
    if (!proc.waitForFinished(30000)) {
        proc.kill();
        proc.waitForFinished(2000);
        errOut = QStringLiteral("readBytesViaPython: python timed out reading ")
                 + path;
        return false;
    }

    const QByteArray errBytes = proc.readAllStandardError();
    if (proc.exitStatus() != QProcess::NormalExit || proc.exitCode() != 0) {
        errOut = QStringLiteral("readBytesViaPython: python failed (exit %1): %2")
                     .arg(proc.exitCode())
                     .arg(QString::fromUtf8(errBytes).trimmed());
        return false;
    }

    const QByteArray b64 = proc.readAllStandardOutput().trimmed();
    if (b64.isEmpty()) {
        errOut = QStringLiteral("readBytesViaPython: python produced no output "
                                "for ") + path;
        return false;
    }

    // base64-decode IN MEMORY. AbortOnBase64DecodingErrors makes a non-base64
    // stdout (e.g. an interpreter banner leaking through) a hard failure rather
    // than silently yielding partial garbage the caller would mis-parse.
    const QByteArray::FromBase64Result decoded = QByteArray::fromBase64Encoding(
        b64, QByteArray::Base64Encoding | QByteArray::AbortOnBase64DecodingErrors);
    if (!decoded) {   // FromBase64Result::operator bool: false on decode error
        errOut = QStringLiteral("readBytesViaPython: stdout was not valid base64 "
                                "for ") + path
                 + (errBytes.isEmpty()
                        ? QString()
                        : (QStringLiteral(" (stderr: ")
                           + QString::fromUtf8(errBytes).trimmed()
                           + QLatin1Char(')')));
        return false;
    }
    // An empty source file decodes to empty bytes; treat that as a (degenerate)
    // success -- the caller's JSON parse will reject empty content loudly.
    out = decoded.decoded;
    return true;
}

QString decryptToTempViaPython(const QString& path,
                               const QString& pythonExe,
                               QString& errOut)
{
    errOut.clear();

    if (pythonExe.isEmpty()) {
        errOut = QStringLiteral("decryptToTempViaPython: no python interpreter "
                                "available to read MIP-labelled file: ") + path;
        return QString();
    }
    if (!QFile::exists(path)) {
        errOut = QStringLiteral("decryptToTempViaPython: source does not exist: ")
                 + path;
        return QString();
    }

    // Choose a destination temp path. We reserve a unique name via
    // QTemporaryFile then close+remove it so python owns the create — the
    // suffix preserves the source extension so a SQLite open of the temp picks
    // the right journal mode siblings if any (there should be none here).
    const QString suffix = QFileInfo(path).suffix();
    const QString tmpDir = QStandardPaths::writableLocation(
        QStandardPaths::TempLocation);
    QString dstPath;
    {
        const QString templ = tmpDir + QStringLiteral("/dve_mip_XXXXXX")
            + (suffix.isEmpty() ? QString() : (QLatin1Char('.') + suffix));
        QTemporaryFile reserve(templ);
        reserve.setAutoRemove(false);
        if (!reserve.open()) {
            errOut = QStringLiteral("decryptToTempViaPython: cannot reserve temp "
                                    "path: ") + reserve.errorString();
            return QString();
        }
        dstPath = reserve.fileName();
        reserve.close();
        // Remove the empty reservation so python's open(dst,'wb') is the creator
        // (avoids any lingering handle / partial-file ambiguity on Windows).
        QFile::remove(dstPath);
    }

    // Tiny script: read the (MIP-decrypted, because bundled python is on the
    // allowlist) source bytes and write them verbatim to dst. Paths are passed
    // as argv, not interpolated, so backslashes / quotes can't break it.
    static const char kScript[] =
        "import sys\n"
        "src, dst = sys.argv[1], sys.argv[2]\n"
        "with open(src, 'rb') as f:\n"
        "    data = f.read()\n"
        "with open(dst, 'wb') as f:\n"
        "    f.write(data)\n"
        "sys.stdout.write(str(len(data)))\n";

    QTemporaryFile scriptTmp(tmpDir + QStringLiteral("/dve_mip_decrypt_XXXXXX.py"));
    scriptTmp.setAutoRemove(true);
    if (!scriptTmp.open()) {
        errOut = QStringLiteral("decryptToTempViaPython: cannot write script: ")
                 + scriptTmp.errorString();
        QFile::remove(dstPath);
        return QString();
    }
    scriptTmp.write(kScript);
    scriptTmp.flush();
    const QString scriptPath = scriptTmp.fileName();
    scriptTmp.close();  // close handle so QProcess can read it on Windows

    QProcess proc;
    QProcessEnvironment env = QProcessEnvironment::systemEnvironment();
    env.insert(QStringLiteral("PYTHONIOENCODING"), QStringLiteral("utf-8"));
    env.insert(QStringLiteral("PYTHONUTF8"), QStringLiteral("1"));
    proc.setProcessEnvironment(env);
    proc.start(pythonExe, { scriptPath, path, dstPath });
    // 30 s budget: mirrors ExcelReader/runPython. A local file copy is fast;
    // the budget only guards against a hung interpreter.
    if (!proc.waitForFinished(30000)) {
        proc.kill();
        proc.waitForFinished(2000);
        errOut = QStringLiteral("decryptToTempViaPython: python timed out "
                                "decrypting ") + path;
        QFile::remove(dstPath);
        return QString();
    }

    const QByteArray errBytes = proc.readAllStandardError();
    if (proc.exitStatus() != QProcess::NormalExit || proc.exitCode() != 0) {
        errOut = QStringLiteral("decryptToTempViaPython: python failed (exit %1): %2")
                     .arg(proc.exitCode())
                     .arg(QString::fromUtf8(errBytes).trimmed());
        QFile::remove(dstPath);
        return QString();
    }

    // Sanity: the temp must exist and be non-empty AND must NOT itself carry
    // the marker (i.e. python actually decrypted). If the copy is still
    // ciphertext, python was NOT allowlisted — fail loudly rather than hand the
    // caller another unreadable file.
    QFileInfo dstInfo(dstPath);
    if (!dstInfo.exists() || dstInfo.size() == 0) {
        errOut = QStringLiteral("decryptToTempViaPython: temp copy missing/empty "
                                "after python ran for ") + path;
        QFile::remove(dstPath);
        return QString();
    }
    if (looksEncrypted(dstPath)) {
        errOut = QStringLiteral("decryptToTempViaPython: temp copy still encrypted "
                                "(bundled python is not MIP-allowlisted?) for ")
                 + path;
        QFile::remove(dstPath);
        return QString();
    }

    return dstPath;
}

} // namespace DVE

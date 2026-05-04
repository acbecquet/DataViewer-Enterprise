# DataViewer Enterprise Optimization — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Optimize DataViewer Enterprise across build, security, data integrity, architecture, and rendering — without changing any external behavior.

**Architecture:** 5-phase inside-out approach. Each phase has a testing gate before the next begins. Phase 4 (MainWindow decomposition) uses triple-review per controller extraction with golden-file comparison.

**Tech Stack:** C++17, Qt 6.10.1, MinGW GCC 13.1.0, qmake, SQLite, Python subprocess, Inno Setup 6

**Spec:** `docs/superpowers/specs/2026-04-10-optimization-plan-design.md`

**Critical constraints:**
- Python subprocess for Excel/PPTX MUST NOT change (company encryption requires it)
- All inputs and outputs must remain identical
- DataViewer.exe and DataViewer-setup.exe names must never change
- Sensory metric ordering preserved (Burnt Taste first in entry, Overall Liking at 12 o'clock)

---

## Phase 1: Build & Infrastructure

### Task 1: Update compiler flags

**Files:**
- Modify: `DataViewerEnterprise.pro:91` (after DEFINES line)

- [ ] **Step 1: Add release optimization flags to .pro file**

Add a release-mode block after line 91 in `DataViewerEnterprise.pro`:

```pro
# ─── Release optimizations ───────────────────────────────────────────────────
CONFIG(release, debug|release) {
    QMAKE_CXXFLAGS += -O3 -flto -Wpedantic
    QMAKE_LFLAGS   += -flto
}

# Treat warnings as errors in all builds
QMAKE_CXXFLAGS += -Werror
```

Also remove the existing `DEFINES += QT_DEPRECATED_WARNINGS` on line 91 and replace:

```pro
DEFINES += QT_DEPRECATED_WARNINGS
QMAKE_CXXFLAGS += -Wno-deprecated-declarations
```

This keeps `-Werror` from breaking on Qt deprecation warnings.

- [ ] **Step 2: Regenerate Makefiles**

Run:
```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
"C:/Qt/6.10.1/mingw_64/bin/qmake.exe" DataViewerEnterprise.pro
```

Expected: Makefiles regenerated with new flags.

- [ ] **Step 3: Build release and fix any warnings**

Run:
```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
mingw32-make -f Makefile.Release clean && mingw32-make -f Makefile.Release -j8
```

Expected: Build succeeds. If `-Werror` catches warnings, fix them one by one. Common issues:
- Unused variables → remove or cast to `(void)var;`
- Signed/unsigned comparison → use appropriate cast
- Implicit fallthrough in switch → add `[[fallthrough]];`

- [ ] **Step 4: Run all test suites**

Run:
```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/tests"
mingw32-make -j8 && for /d %t in (tst_*) do (cd %t && release\*.exe -silent && cd ..)
```

Expected: All 10 suites pass.

- [ ] **Step 5: Commit**

```bash
git add DataViewerEnterprise.pro
git commit -m "build: add -O3, LTO, -Werror, -Wpedantic release flags"
```

---

### Task 2: Upgrade installer compression

**Files:**
- Modify: `installer.iss:11`

- [ ] **Step 1: Change compression level**

In `installer.iss`, change line 11:

```diff
-Compression=lzma2
+Compression=lzma2/ultra64
```

- [ ] **Step 2: Commit**

```bash
git add installer.iss
git commit -m "build: upgrade installer to lzma2/ultra64 compression"
```

---

### Task 3: Phase 1 Testing Gate

- [ ] **Step 1: Full smoke test**

Manually verify:
1. Launch DataViewer.exe from release folder
2. Load an Excel file → verify data appears in table
3. Edit a cell → verify Excel write (check modified timestamp)
4. Generate a PPTX test report → open in PowerPoint, verify content
5. Switch to Sensory mode → create session → add sample → verify radar chart
6. Switch to Detailed Sensory mode → verify dual radar charts
7. Switch back to TPM mode → verify table/plot restored

- [ ] **Step 2: Mark Phase 1 complete**

All builds pass, all tests pass, smoke test passes. Phase 1 done.

---

## Phase 2: Security Fixes

### Task 4: Fix SQL injection in database_explorer.py

**Files:**
- Modify: `DataViewer/database_explorer.py:158,270-272`

- [ ] **Step 1: Fix .format() SQL injection at line 270-272**

Replace the `show_recent_activity` query (lines 264-272):

```python
            cursor.execute("""
                SELECT
                    filename,
                    LENGTH(file_content) as size_bytes,
                    created_at
                FROM files
                WHERE created_at > datetime('now', ? || ' days')
                ORDER BY created_at DESC
            """, (f"-{days}",))
```

- [ ] **Step 2: Fix f-string LIMIT injection at line 157-158**

Replace lines 157-158:

```python
            if limit:
                query += " LIMIT ?"
                cursor.execute(query, (limit,))
            else:
                cursor.execute(query)
```

This requires restructuring the `list_files` method slightly since `cursor.execute(query)` on line 160 currently runs unconditionally. The full replacement for lines 155-160:

```python
            if limit:
                query += " LIMIT ?"
                cursor.execute(query, (int(limit),))
            else:
                cursor.execute(query)
```

- [ ] **Step 3: Verify by running database_explorer.py**

Run:
```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer"
python database_explorer.py
```

Expected: Explorer launches without errors, queries return correct results.

- [ ] **Step 4: Commit**

```bash
git add DataViewer/database_explorer.py
git commit -m "security: fix SQL injection in database_explorer.py"
```

---

### Task 5: Externalize hardcoded network paths

**Files:**
- Modify: `DataViewer/database_manager.py:20-23`

- [ ] **Step 1: Replace hardcoded values with config file lookup**

Replace lines 20-23 in `database_manager.py`:

```python
    # CONFIGURATION — read from config file, fall back to defaults
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "network_config.json")
    default_config = {
        "synology_ip": "192.168.222.10",
        "database_relative_path": r"SDR\Device Group\Database",
        "database_filename": "dataviewer.db"
    }
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                cfg = json.load(f)
            SYNOLOGY_IP = cfg.get("synology_ip", default_config["synology_ip"])
            DATABASE_RELATIVE_PATH = cfg.get("database_relative_path", default_config["database_relative_path"])
            DATABASE_FILENAME = cfg.get("database_filename", default_config["database_filename"])
        else:
            SYNOLOGY_IP = default_config["synology_ip"]
            DATABASE_RELATIVE_PATH = default_config["database_relative_path"]
            DATABASE_FILENAME = default_config["database_filename"]
    except Exception:
        SYNOLOGY_IP = default_config["synology_ip"]
        DATABASE_RELATIVE_PATH = default_config["database_relative_path"]
        DATABASE_FILENAME = default_config["database_filename"]
```

- [ ] **Step 2: Create default config file**

Create `DataViewer/network_config.json`:

```json
{
    "synology_ip": "192.168.222.10",
    "database_relative_path": "SDR\\Device Group\\Database",
    "database_filename": "dataviewer.db"
}
```

- [ ] **Step 3: Add config to .gitignore**

Add to `DataViewer/.gitignore` (create if doesn't exist):

```
network_config.json
```

And keep a template:

Create `DataViewer/network_config.json.example`:
```json
{
    "synology_ip": "192.168.222.10",
    "database_relative_path": "SDR\\Device Group\\Database",
    "database_filename": "dataviewer.db"
}
```

- [ ] **Step 4: Commit**

```bash
git add DataViewer/database_manager.py DataViewer/network_config.json.example
git commit -m "security: externalize hardcoded network paths to config file"
```

---

### Task 6: Secure API key storage

**Files:**
- Modify: `DataViewer-Enterprise/src/MainWindow.cpp:3600-3628,3647-3665`
- Modify: `DataViewer-Enterprise/src/MainWindow.h` (add method declaration)

- [ ] **Step 1: Replace plaintext key file with Windows Credential Manager**

In `MainWindow.h`, add two private method declarations (after line 142):

```cpp
    QString loadApiKey();
    bool    saveApiKey(const QString& key);
```

- [ ] **Step 2: Implement credential storage using QSettings with obfuscation**

Since DPAPI requires Windows-specific headers, use Qt's built-in `QSettings` with the Windows registry backend (which is more secure than a plaintext file) combined with simple XOR obfuscation. This is significantly better than plaintext/Base64 while staying cross-compile safe.

Add to the bottom of `MainWindow.cpp` (before the closing `} // namespace DVE`):

```cpp
QString MainWindow::loadApiKey()
{
    // Migration: check old plaintext file first
    QString keyFilePath = QCoreApplication::applicationDirPath()
                          + "/anthropic_api_key.txt";
    QFile keyFile(keyFilePath);
    if (keyFile.open(QIODevice::ReadOnly | QIODevice::Text)) {
        QTextStream in(&keyFile);
        QString key = in.readAll().trimmed();
        keyFile.close();
        if (!key.isEmpty()) {
            // Migrate: save to registry, delete plaintext file
            saveApiKey(key);
            keyFile.remove();
            return key;
        }
    }

    // Load from QSettings (stored in Windows registry under HKCU)
    QSettings settings("SDR", "DataViewer Enterprise");
    QByteArray stored = settings.value("translator/apiKey").toByteArray();
    if (stored.isEmpty()) return {};

    // De-obfuscate (XOR with fixed key)
    const QByteArray mask = QByteArrayLiteral("DVE_2026_TRANSLATOR");
    QByteArray decoded = QByteArray::fromBase64(stored);
    for (int i = 0; i < decoded.size(); ++i)
        decoded[i] = decoded[i] ^ mask[i % mask.size()];
    return QString::fromUtf8(decoded);
}

bool MainWindow::saveApiKey(const QString& key)
{
    const QByteArray raw = key.toUtf8();
    const QByteArray mask = QByteArrayLiteral("DVE_2026_TRANSLATOR");
    QByteArray obfuscated = raw;
    for (int i = 0; i < obfuscated.size(); ++i)
        obfuscated[i] = obfuscated[i] ^ mask[i % mask.size()];

    QSettings settings("SDR", "DataViewer Enterprise");
    settings.setValue("translator/apiKey", obfuscated.toBase64());
    return true;
}
```

- [ ] **Step 3: Update onLaunchTranslator() to use new methods**

Replace lines 3600-3628 in `onLaunchTranslator()`:

```cpp
    // 2. Get API key — from registry if available, otherwise prompt
    QString apiKey = loadApiKey();

    if (apiKey.isEmpty()) {
        bool ok = false;
        apiKey = QInputDialog::getText(
            this,
            "Anthropic API Key Required",
            "Enter your Anthropic API key to use the Document Translator:",
            QLineEdit::Normal, QString(), &ok);
        if (!ok || apiKey.trimmed().isEmpty())
            return;
        apiKey = apiKey.trimmed();
        saveApiKey(apiKey);
    }
```

- [ ] **Step 4: Update writeTranslatorConfig() to use obfuscation too**

Replace `writeTranslatorConfig` (lines 3647-3665):

```cpp
bool MainWindow::writeTranslatorConfig(const QString& apiKey)
{
    QString configDir  = QDir::homePath() + "/.document_translator";
    QString configPath = configDir + "/config.dat";

    if (!QDir(configDir).mkpath("."))
        return false;

    // XOR + Base64 — matched by the translator's config reader
    const QByteArray raw = apiKey.toUtf8();
    const QByteArray mask = QByteArrayLiteral("DVE_2026_TRANSLATOR");
    QByteArray obfuscated = raw;
    for (int i = 0; i < obfuscated.size(); ++i)
        obfuscated[i] = obfuscated[i] ^ mask[i % mask.size()];

    QJsonObject obj;
    obj["api_key"] = QLatin1String(obfuscated.toBase64());
    QJsonDocument doc(obj);

    QFile file(configPath);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Truncate))
        return false;

    file.write(doc.toJson(QJsonDocument::Compact));
    return true;
}
```

- [ ] **Step 5: Commit**

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "security: replace plaintext API key storage with obfuscated registry storage"
```

---

### Task 7: Fix temp file permissions, XML escaping, and BLOB size limit

**Files:**
- Modify: `DataViewer-Enterprise/src/ExcelReader.cpp:1153-1159`
- Modify: `DataViewer-Enterprise/src/reporting/PptxWriter.cpp:772`
- Modify: `DataViewer-Enterprise/src/database/DatabaseManager.cpp:466-501`

- [ ] **Step 1: Set restrictive permissions on temp Python script**

In `ExcelReader.cpp`, after line 1158 (`scriptFile.write(kPythonScript);`), add before `scriptFile.close();`:

```cpp
    scriptFile.write(kPythonScript);
    scriptFile.setPermissions(QFile::ReadOwner | QFile::WriteOwner);
    scriptFile.close();
```

- [ ] **Step 2: Add single-quote XML escaping in PptxWriter.cpp**

After line 772 (the `&quot;` replacement), add:

```cpp
    safeText.replace(QLatin1Char('\''), QStringLiteral("&#39;"));
```

- [ ] **Step 3: Add BLOB size validation in DatabaseManager.cpp**

In `DatabaseManager.cpp`, find the two image-reading blocks (around lines 466-470 and 497-501). In both places, wrap the `readAll()` with a size check. Replace:

```cpp
                    QFile imgFile(sr.imagePaths[ii]);
                    if (imgFile.open(QIODevice::ReadOnly))
                        imgData = imgFile.readAll();
```

With:

```cpp
                    QFile imgFile(sr.imagePaths[ii]);
                    if (imgFile.open(QIODevice::ReadOnly)) {
                        constexpr qint64 kMaxImageSize = 100 * 1024 * 1024; // 100 MB
                        if (imgFile.size() <= kMaxImageSize)
                            imgData = imgFile.readAll();
                        else
                            logDebug("Skipping oversized image: " + sr.imagePaths[ii]);
                    }
```

Apply this to both occurrences (lines ~468-470 and ~499-501).

- [ ] **Step 4: Build and test**

Run:
```bash
mingw32-make -f Makefile.Release -j8
```

Expected: Compiles cleanly.

- [ ] **Step 5: Commit**

```bash
git add src/ExcelReader.cpp src/reporting/PptxWriter.cpp src/database/DatabaseManager.cpp
git commit -m "security: temp file perms, XML quote escaping, BLOB size limit"
```

---

### Task 8: Remove hardcoded dev path and sanitize error messages

**Files:**
- Modify: `DataViewer-Enterprise/src/main.cpp:44-48`
- Modify: `DataViewer-Enterprise/src/MainWindow.cpp` (error message calls)

- [ ] **Step 1: Remove hardcoded dev path from main.cpp**

Replace lines 44-48 in `main.cpp`:

```cpp
    QStringList iconCandidates = {
        QCoreApplication::applicationDirPath() + "/resources/images/ccell_icon.png",
        QCoreApplication::applicationDirPath() + "/../resources/images/ccell_icon.png",
    };
```

(Remove the third entry with the hardcoded `C:/Users/S1134987/...` path.)

- [ ] **Step 2: Sanitize error messages shown to users**

Search MainWindow.cpp for `showError` calls that expose raw error text. The pattern to fix:

Find calls like:
```cpp
showError("Load Error", "Failed to load file.\n" + m_processor->lastError());
```

Replace with:
```cpp
qWarning() << "Load error:" << m_processor->lastError();
showError("Load Error", "Failed to load file. Check the log for details.");
```

Apply this pattern to all `showError` calls that include `lastError()`, `lastError().text()`, or raw exception details. Keep the `qWarning()` for debugging. The user sees a clean message.

Key locations to update (search for `showError.*lastError` and `showError.*failed`):
- `onFileLoadFinished()` — file load errors
- `onReportFinished()` — report generation errors
- `onUpdateDatabase()` — database save errors

- [ ] **Step 3: Build and test**

```bash
mingw32-make -f Makefile.Release -j8
```

- [ ] **Step 4: Commit**

```bash
git add src/main.cpp src/MainWindow.cpp
git commit -m "security: remove dev path, sanitize user-facing error messages"
```

---

### Task 9: Phase 2 Testing Gate

- [ ] **Step 1: Run all 10 test suites**

All must pass.

- [ ] **Step 2: Test Excel load with company files**

Load an encrypted company Excel file. Verify Python subprocess still works, data appears correctly.

- [ ] **Step 3: Test PPTX generation**

Generate a test report. Open in PowerPoint. Verify all slides, tables, charts, and images are present and identical to pre-Phase-2 output.

- [ ] **Step 4: Test API key round-trip**

1. Launch Document Translator for the first time (should prompt for key)
2. Enter a test key → verify it saves to registry
3. Close and re-launch → verify key is loaded from registry without prompting
4. If old `anthropic_api_key.txt` exists, verify it migrates and deletes the file

- [ ] **Step 5: Test input with apostrophes**

Create a sample with name `O'Brien's Test` → generate PPTX → open in PowerPoint → verify name renders correctly (tests the XML single-quote escaping).

- [ ] **Step 6: Mark Phase 2 complete**

---

## Phase 3: Data Layer Hardening

### Task 10: Add input validation utility

**Files:**
- Create: `DataViewer-Enterprise/src/utils/InputValidator.h`

- [ ] **Step 1: Create the validation header**

```cpp
#pragma once

#include <QString>
#include <QRegularExpression>

namespace DVE {

class InputValidator {
public:
    // Returns empty string if valid, error message if invalid.
    static QString validateName(const QString& value, const QString& fieldName,
                                int maxLen = 255)
    {
        if (value.trimmed().isEmpty())
            return fieldName + " cannot be empty.";
        if (value.length() > maxLen)
            return fieldName + " cannot exceed " + QString::number(maxLen) + " characters.";
        return {};
    }

    static QString validateFilePath(const QString& path)
    {
        if (path.isEmpty())
            return "File path is empty.";
        // Reject path traversal patterns
        if (path.contains(".."))
            return "File path contains invalid characters.";
        return {};
    }

    static bool isValidImageFile(const QString& path)
    {
        QFile f(path);
        if (!f.open(QIODevice::ReadOnly)) return false;
        QByteArray header = f.read(8);
        f.close();
        if (header.size() < 4) return false;

        // PNG: 89 50 4E 47
        if (header[0] == '\x89' && header[1] == 'P' &&
            header[2] == 'N'    && header[3] == 'G')
            return true;

        // JPEG: FF D8 FF
        if (static_cast<unsigned char>(header[0]) == 0xFF &&
            static_cast<unsigned char>(header[1]) == 0xD8 &&
            static_cast<unsigned char>(header[2]) == 0xFF)
            return true;

        // BMP: 42 4D
        if (header[0] == 'B' && header[1] == 'M')
            return true;

        return false;
    }
};

} // namespace DVE
```

- [ ] **Step 2: Add to .pro file**

Add to the HEADERS section in `DataViewerEnterprise.pro`:

```pro
    src/utils/InputValidator.h \
```

- [ ] **Step 3: Commit**

```bash
git add src/utils/InputValidator.h DataViewerEnterprise.pro
git commit -m "feat: add InputValidator utility for system-boundary validation"
```

---

### Task 11: Apply input validation to UI boundaries

**Files:**
- Modify: `DataViewer-Enterprise/src/ui/SensoryPanel.cpp` (session save)
- Modify: `DataViewer-Enterprise/src/ui/DetailedSensoryPanel.cpp` (session save)
- Modify: `DataViewer-Enterprise/src/MainWindow.cpp` (file load path validation)

- [ ] **Step 1: Add validation before sensory session save**

In `SensoryPanel.cpp`, at the top of the `save()` method, add:

```cpp
#include "utils/InputValidator.h"

// At top of save():
QString err = InputValidator::validateName(m_currentSession.sessionName, "Session name");
if (!err.isEmpty()) {
    QMessageBox::warning(this, "Validation", err);
    return;
}
err = InputValidator::validateName(m_currentSession.testerName, "Tester name");
if (!err.isEmpty()) {
    QMessageBox::warning(this, "Validation", err);
    return;
}
```

- [ ] **Step 2: Apply same pattern in DetailedSensoryPanel::save()**

Same include and validation block for `sessionName` and `testerName`.

- [ ] **Step 3: Add image validation before DB storage**

In `DatabaseManager.cpp`, in the image-reading blocks (the two locations from Task 7), add magic byte validation before `readAll()`:

```cpp
#include "utils/InputValidator.h"

// In the image insertion loop, before reading:
if (!InputValidator::isValidImageFile(sr.imagePaths[ii])) {
    logDebug("Skipping non-image file: " + sr.imagePaths[ii]);
    continue;
}
```

- [ ] **Step 4: Build and test**

```bash
mingw32-make -f Makefile.Release -j8
```

- [ ] **Step 5: Commit**

```bash
git add src/ui/SensoryPanel.cpp src/ui/DetailedSensoryPanel.cpp src/database/DatabaseManager.cpp
git commit -m "hardening: add input validation at system boundaries"
```

---

### Task 12: Cache Python script and validate Python path

**Files:**
- Modify: `DataViewer-Enterprise/src/ExcelReader.h`
- Modify: `DataViewer-Enterprise/src/ExcelReader.cpp:1144-1175`

- [ ] **Step 1: Add static member for cached script path**

In `ExcelReader.h`, add a private static member:

```cpp
    static QString s_cachedScriptPath;
```

- [ ] **Step 2: Implement script caching in runPythonReader()**

Replace lines 1149-1159 in `ExcelReader.cpp`:

```cpp
QString ExcelReader::s_cachedScriptPath;

// Inside runPythonReader(), replace script writing block:
    // Write the script to a temp file (once per session)
    if (s_cachedScriptPath.isEmpty() || !QFile::exists(s_cachedScriptPath)) {
        QString tempDir = QStandardPaths::writableLocation(QStandardPaths::TempLocation);
        s_cachedScriptPath = tempDir + "/dve_xlsx_reader.py";

        QFile scriptFile(s_cachedScriptPath);
        if (!scriptFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
            error = "Cannot write Python script to: " + s_cachedScriptPath;
            return false;
        }
        scriptFile.write(kPythonScript);
        scriptFile.setPermissions(QFile::ReadOwner | QFile::WriteOwner);
        scriptFile.close();
    }
```

Also remove the two `QFile::remove(scriptPath)` calls (lines 1171 and 1175) since we now keep the cached file.

- [ ] **Step 3: Add Python executable validation in findPython()**

After line 1131 (`if (QFile::exists(bundled))`), add an executability check:

```cpp
    if (QFile::exists(bundled)) {
        QFileInfo fi(bundled);
        if (fi.isExecutable())
            return bundled;
        writeLog("Bundled Python exists but is not executable: " + bundled);
    }
```

- [ ] **Step 4: Build and test**

```bash
mingw32-make -f Makefile.Release -j8
```

Test: Load an Excel file. Verify it loads correctly. Check that `%TEMP%/dve_xlsx_reader.py` exists and is reused on second load.

- [ ] **Step 5: Commit**

```bash
git add src/ExcelReader.h src/ExcelReader.cpp
git commit -m "hardening: cache Python script per session, validate Python path"
```

---

### Task 13: Phase 3 Testing Gate

- [ ] **Step 1: Run all 10 test suites**

All must pass.

- [ ] **Step 2: Test input validation**

1. Try saving a sensory session with empty name → should show warning
2. Try saving with a 300-character name → should show warning
3. Save with `O'Brien` name → should succeed
4. Save with unicode name `Tëst Üser` → should succeed

- [ ] **Step 3: Test image validation**

1. Attach a valid PNG → should succeed
2. Attach a valid JPEG → should succeed
3. Rename a .txt file to .png and try attaching → should be silently skipped in DB

- [ ] **Step 4: Test Python script caching**

1. Load an Excel file → check `%TEMP%/dve_xlsx_reader.py` exists
2. Load another Excel file → verify same script file reused (check modified timestamp)

- [ ] **Step 5: Full end-to-end smoke test**

Load Excel → edit cells → save → generate PPTX → verify output. All three modes.

- [ ] **Step 6: Mark Phase 3 complete**

---

## Phase 4: MainWindow Decomposition

**IMPORTANT:** Before starting this phase, generate a "golden" PPTX report and save it for comparison after each extraction.

### Task 14: Extract ReportController

**Files:**
- Create: `DataViewer-Enterprise/src/controllers/ReportController.h`
- Create: `DataViewer-Enterprise/src/controllers/ReportController.cpp`
- Modify: `DataViewer-Enterprise/src/MainWindow.h`
- Modify: `DataViewer-Enterprise/src/MainWindow.cpp`
- Modify: `DataViewerEnterprise.pro`

- [ ] **Step 1: Create ReportController header**

```cpp
#pragma once

#include <QObject>

class QProgressBar;

namespace DVE {

class ReportGenerator;
struct FileResult;
struct SheetResult;

class ReportController : public QObject {
    Q_OBJECT

public:
    explicit ReportController(ReportGenerator* reportGen,
                              QProgressBar* progressBar,
                              QObject* parent = nullptr);

    // Provide data context for report generation
    void setCurrentData(const FileResult* file, const SheetResult* sheet);

    // Provide cleanup exclusion data
    void setExcludedRows(const QMap<QString, QSet<int>>* excludedRows,
                         int currentFileIndex);

public slots:
    void onGenerateTestReport();
    void onGenerateFullReport();
    void onReportFinished(bool success, const QString& path);

signals:
    void progressChanged(int percent, const QString& message);

private:
    ReportGenerator* m_reportGen;
    QProgressBar*    m_progressBar;
    const FileResult*  m_currentFile  = nullptr;
    const SheetResult* m_currentSheet = nullptr;
    const QMap<QString, QSet<int>>* m_excludedRows = nullptr;
    int m_currentFileIndex = -1;
};

} // namespace DVE
```

- [ ] **Step 2: Create ReportController.cpp**

Move the bodies of `onGenerateTestReport()` (lines 1887-1911), `onGenerateFullReport()` (lines 1913-1932), and `onReportFinished()` (lines 1935-1950) from MainWindow.cpp into the new file.

Replace references to `currentSheet()` / `currentFile()` with the stored pointers `m_currentSheet` / `m_currentFile`.

Replace `showInfo(...)` and `showError(...)` calls with `QMessageBox::information(qobject_cast<QWidget*>(parent()), ...)` and similar.

Replace `setProgress(pct, msg)` with `m_progressBar->setValue(pct)` and status updates.

The constructor wires `m_reportGen->reportFinished` to `onReportFinished()` and `m_reportGen->progressChanged` to the progress bar.

- [ ] **Step 3: Update MainWindow**

In `MainWindow.h`:
- Add `#include "controllers/ReportController.h"`
- Add member: `ReportController* m_reportCtrl;`
- Remove slot declarations: `onGenerateTestReport()`, `onGenerateFullReport()`, `onReportFinished()`

In `MainWindow.cpp`:
- In constructor, create: `m_reportCtrl = new ReportController(m_reportGen, m_progressBar, this);`
- Replace ribbon connects (lines 224-225):
  ```cpp
  connect(m_reportBtn1, &QToolButton::clicked, m_reportCtrl, &ReportController::onGenerateTestReport);
  connect(m_reportBtn2, &QToolButton::clicked, m_reportCtrl, &ReportController::onGenerateFullReport);
  ```
- In `setupConnections()`, remove the `m_reportGen` signal connections (lines 640-642) — they're now wired in ReportController's constructor.
- In `displayCurrentSample()` and any place that updates current file/sheet, add:
  ```cpp
  m_reportCtrl->setCurrentData(currentFile(), currentSheet());
  m_reportCtrl->setExcludedRows(&m_excludedRows, m_currentFileIndex);
  ```
- Delete the old method bodies from MainWindow.cpp.

- [ ] **Step 4: Update .pro file**

Add to SOURCES:
```pro
    src/controllers/ReportController.cpp \
```
Add to HEADERS:
```pro
    src/controllers/ReportController.h \
```
Add to INCLUDEPATH:
```pro
               src/controllers \
```

- [ ] **Step 5: Build**

```bash
"C:/Qt/6.10.1/mingw_64/bin/qmake.exe" DataViewerEnterprise.pro
mingw32-make -f Makefile.Release -j8
```

- [ ] **Step 6: Run tests + regression**

Run all 10 test suites. Then manually:
1. Generate a test report → compare PPTX XML structure to golden file
2. Generate a full report → verify all slides present

- [ ] **Step 7: Commit**

```bash
git add src/controllers/ src/MainWindow.h src/MainWindow.cpp DataViewerEnterprise.pro
git commit -m "refactor: extract ReportController from MainWindow"
```

---

### Task 15: Extract ImageController

**Files:**
- Create: `DataViewer-Enterprise/src/controllers/ImageController.h`
- Create: `DataViewer-Enterprise/src/controllers/ImageController.cpp`
- Modify: `DataViewer-Enterprise/src/MainWindow.h`
- Modify: `DataViewer-Enterprise/src/MainWindow.cpp`
- Modify: `DataViewerEnterprise.pro`

- [ ] **Step 1: Create ImageController header**

```cpp
#pragma once

#include <QObject>
#include <QFileSystemWatcher>

class QToolButton;
class QPushButton;

namespace DVE {

class DatabaseManager;
struct FileResult;
struct SheetResult;
struct SampleResult;
class SensoryPanel;
class DetailedSensoryPanel;

class ImageController : public QObject {
    Q_OBJECT

public:
    explicit ImageController(DatabaseManager* db, QObject* parent = nullptr);

    void setWidgets(QPushButton* loadBtn, QPushButton* viewBtn, QToolButton* inboxBtn);
    void setCurrentData(FileResult* file, SheetResult* sheet,
                        SampleResult* sample, int fileIdx, int sheetIdx, int sampleIdx);
    void setSensoryPanels(SensoryPanel* sensory, DetailedSensoryPanel* detailed);
    void setModeFlags(bool sensoryMode, bool detailedSensoryMode);

public slots:
    void onLoadImages();
    void onViewImages();
    void onOpenImageInbox();
    void onInboxFolderChanged(const QString& path);
    void updateImageButton();

signals:
    void imagesChanged();     // triggers displayCurrentSample() in MainWindow
    void fileModified();      // triggers markFileModified() in MainWindow

private:
    DatabaseManager* m_db;
    QFileSystemWatcher* m_inboxWatcher;
    QString m_inboxPath;

    QPushButton* m_loadImagesBtn  = nullptr;
    QPushButton* m_viewImagesBtn  = nullptr;
    QToolButton* m_inboxBtn       = nullptr;

    FileResult*   m_currentFile   = nullptr;
    SheetResult*  m_currentSheet  = nullptr;
    SampleResult* m_currentSample = nullptr;
    int m_currentFileIdx   = -1;
    int m_currentSheetIdx  = -1;
    int m_currentSampleIdx = -1;

    SensoryPanel*         m_sensoryPanel         = nullptr;
    DetailedSensoryPanel* m_detailedSensoryPanel = nullptr;
    bool m_sensoryMode         = false;
    bool m_detailedSensoryMode = false;
};

} // namespace DVE
```

- [ ] **Step 2: Create ImageController.cpp**

Move bodies of `onLoadImages()` (lines 3347-3389), `onViewImages()` (lines 3391-3427), `updateImageButton()` (lines 3429-3487), `onOpenImageInbox()` (lines 3489-3567), `onInboxFolderChanged()` (lines 3569-3584) from MainWindow.cpp.

Emit `imagesChanged()` and `fileModified()` where MainWindow previously called `displayCurrentSample()` and `markFileModified()`.

- [ ] **Step 3: Update MainWindow**

In `MainWindow.h`:
- Add member: `ImageController* m_imageCtrl;`
- Remove slot declarations: `onLoadImages()`, `onViewImages()`, `onOpenImageInbox()`, `onInboxFolderChanged()`
- Remove members: `m_inboxWatcher`, `m_inboxPath`

In `MainWindow.cpp`:
- Create in constructor and wire signals
- Connect `m_imageCtrl->imagesChanged()` to `displayCurrentSample()`
- Connect `m_imageCtrl->fileModified()` to `markFileModified()`
- Redirect button connects to ImageController slots
- Delete old method bodies

- [ ] **Step 4: Update .pro, build, test, commit**

Same pattern as Task 14. Run all tests + manual regression on image features:
1. Attach image to sample → verify in properties panel
2. View images → verify ImageViewDialog opens with correct images
3. Open image inbox → verify file watcher works

```bash
git commit -m "refactor: extract ImageController from MainWindow"
```

---

### Task 16: Extract ModeController

**Files:**
- Create: `DataViewer-Enterprise/src/controllers/ModeController.h`
- Create: `DataViewer-Enterprise/src/controllers/ModeController.cpp`
- Modify: MainWindow.h, MainWindow.cpp, DataViewerEnterprise.pro

- [ ] **Step 1: Create ModeController**

Owns: `toggleSensoryMode()`, `toggleDetailedSensoryMode()`, `initSensoryPanel()`, `initDetailedSensoryPanel()`, `updateRibbonForMode()`, and the mode state flags (`m_sensoryMode`, `m_detailedSensoryMode`).

Receives widget pointers via setter: `m_centralStack`, `m_navStack`, `m_sensoryBtn`, `m_detailedSensoryBtn`, ribbon button references.

Emits: `modeChanged(int mode)` (0=TPM, 1=Sensory, 2=DetailedSensory) so MainWindow and other controllers can react.

- [ ] **Step 2: Move method bodies from MainWindow.cpp**

Move `toggleSensoryMode()` (1963-2002), `toggleDetailedSensoryMode()` (2004-2041), `initSensoryPanel()` (2043-2061), `initDetailedSensoryPanel()` (2063-2101), `updateRibbonForMode()` (2103+) into ModeController.cpp.

- [ ] **Step 3: Update MainWindow + wire signals**

- Remove mode-related slots and state from MainWindow.h
- In MainWindow, connect `m_modeCtrl->modeChanged()` to update ImageController and ReportController mode state
- Connect `m_sensoryBtn->toggled` to `m_modeCtrl->toggleSensoryMode`
- Connect `m_detailedSensoryBtn->toggled` to `m_modeCtrl->toggleDetailedSensoryMode`

- [ ] **Step 4: Build, test, cross-review**

Test: Switch TPM → Sensory → Detailed Sensory → TPM. Verify all panels render correctly, ribbon updates, navigator switches. Also re-test report generation and image features (cross-review).

```bash
git commit -m "refactor: extract ModeController from MainWindow"
```

---

### Task 17: Extract FileController

**Files:**
- Create: `DataViewer-Enterprise/src/controllers/FileController.h`
- Create: `DataViewer-Enterprise/src/controllers/FileController.cpp`
- Modify: MainWindow.h, MainWindow.cpp, DataViewerEnterprise.pro

- [ ] **Step 1: Create FileController**

Owns: `onLoadFile()`, `onCloseFile()`, `loadFile()`, `onFileLoadFinished()`, `onRecentFileTriggered()`, `populateFileTree()`, `populateSheetCombo()`, `onFileSelected()`, `onSheetSelected()`, `routeFile()`, `m_loadedFiles`, `m_loadWatcher`, `m_pendingLoadPaths`, `m_currentFileIndex`, `m_currentSheetIndex`, `m_currentSampleIndex`.

Receives: `m_fileTree`, `m_fileCombo`, `m_sheetCombo`, `m_processor`, `m_db` via constructor injection.

Emits: `fileLoaded(int fileIndex)`, `sheetChanged(int sheetIndex)`, `sampleChanged(int sampleIndex)`, `displayRequested()`.

- [ ] **Step 2: Move method bodies**

Move `onLoadFile()` (1309-1321), `onCloseFile()` (1324-1382), `loadFile()`, `onFileLoadFinished()` (1403-1492), `populateFileTree()` (1529-1552), `populateSheetCombo()` (1554-1567), `onFileSelected()`, `onSheetSelected()`, `onPrevSample()`, `onNextSample()`, `routeFile()`.

Also move `m_loadedFiles`, `m_currentFileIndex`, `m_currentSheetIndex`, `m_currentSampleIndex` members.

Provide `currentFile()`, `currentSheet()`, `currentSample()` accessors.

- [ ] **Step 3: Update MainWindow**

MainWindow delegates file operations to FileController. Other controllers that need data access get it through FileController's public accessors.

Wire:
- `m_homeLoadBtn->clicked` → `m_fileCtrl->onLoadFile()`
- `m_homeCloseBtn->clicked` → `m_fileCtrl->onCloseFile()`
- `m_fileCtrl->displayRequested()` → `MainWindow::displayCurrentSample()`
- `m_fileCtrl->fileLoaded()` → update ReportController and ImageController context

- [ ] **Step 4: Build, test, cross-review**

Test: Load file, switch sheets, navigate samples, drag-drop file, close file. Cross-review: report generation, mode switching, image features.

```bash
git commit -m "refactor: extract FileController from MainWindow"
```

---

### Task 18: Extract ExcelController

**Files:**
- Create: `DataViewer-Enterprise/src/controllers/ExcelController.h`
- Create: `DataViewer-Enterprise/src/controllers/ExcelController.cpp`
- Modify: MainWindow.h, MainWindow.cpp, DataViewerEnterprise.pro

- [ ] **Step 1: Create ExcelController**

Owns: `onTableCellChanged()`, `onPropCellChanged()`, `queueExcelWrite()`, `flushExcelWrites()`, `writeCellsToExcel()`, `writeCellToExcel()`, `m_excelWriteTimer`, `m_pendingWrites`, `m_pendingWriteFile`, `m_pendingWriteSheet`.

Receives: data table and prop table pointers, FileController (for current file/sheet access).

Emits: `writeCompleted()`, `fileModified(const QString& filePath)`.

- [ ] **Step 2: Move method bodies**

Move `onTableCellChanged()` (915-996), `onPropCellChanged()` (998-1307), `queueExcelWrite()` (3307-3329), `flushExcelWrites()` (3331-3337), `writeCellsToExcel()` (3270-3305), `writeCellToExcel()` (3264-3268).

Move timer members: `m_excelWriteTimer`, `m_pendingWrites`, `m_pendingWriteFile`, `m_pendingWriteSheet`.

- [ ] **Step 3: Update MainWindow**

Wire:
- `m_dataTable->cellChanged` → `m_excelCtrl->onTableCellChanged()`
- `m_propTable->cellChanged` → `m_excelCtrl->onPropCellChanged()`
- `m_excelCtrl->fileModified()` → update `m_modifiedFilePaths` and start DB save timer

In `closeEvent()`, call `m_excelCtrl->flushExcelWrites()`.

- [ ] **Step 4: Build, test, cross-review**

Test: Edit cell in data table, verify debounced write to Excel, verify file modified timestamp updates. Cross-review all previous controllers.

```bash
git commit -m "refactor: extract ExcelController from MainWindow"
```

---

### Task 19: Extract DbSyncController

**Files:**
- Create: `DataViewer-Enterprise/src/controllers/DbSyncController.h`
- Create: `DataViewer-Enterprise/src/controllers/DbSyncController.cpp`
- Modify: MainWindow.h, MainWindow.cpp, DataViewerEnterprise.pro

- [ ] **Step 1: Create DbSyncController**

Owns: `onUpdateDatabase()`, `markFileModified()`, `updateDbSyncIndicator()`, `m_dbSaveTimer`, `m_modifiedFilePaths`, `m_sensorySessionsDirty`, `m_detailedSensorySessionsDirty`.

Receives: `m_db`, `m_dbSyncLabel`, FileController (for `m_loadedFiles` access), SensoryPanel/DetailedSensoryPanel pointers.

Emits: `syncCompleted()`.

- [ ] **Step 2: Move method bodies**

Move `onUpdateDatabase()` (2661-2711), `markFileModified()` (2713-2721), `updateDbSyncIndicator()` (2723-2787).

Move: `m_dbSaveTimer`, `m_modifiedFilePaths`, `m_sensorySessionsDirty`, `m_detailedSensorySessionsDirty`.

- [ ] **Step 3: Update MainWindow**

Wire:
- `m_excelCtrl->fileModified()` → `m_dbSyncCtrl->markFileModified()`
- Sensory panel `sessionsChanged` → `m_dbSyncCtrl->markSensoryDirty()`
- Keyboard shortcut Ctrl+U → `m_dbSyncCtrl->onUpdateDatabase()`

In `closeEvent()`, call `m_dbSyncCtrl->onUpdateDatabase()` if there are pending changes.

- [ ] **Step 4: Build, test, cross-review**

Test: Edit cell → wait 5s → verify DB sync indicator updates. Manually trigger Ctrl+U → verify sync. Cross-review ALL controllers: file load, report gen, image attach, mode switch, cell edit, DB sync.

```bash
git commit -m "refactor: extract DbSyncController from MainWindow"
```

---

### Task 20: Phase 4 Final Validation

- [ ] **Step 1: Run all 10 test suites**

All must pass.

- [ ] **Step 2: Golden PPTX comparison**

Generate same report as the golden file from before Phase 4. Compare PPTX XML structure:
1. Unzip both .pptx files
2. Diff the slide XML files
3. Verify same number of slides, same table data, same chart images

- [ ] **Step 3: Excel write comparison**

Edit a cell in the data table. Compare the saved .xlsx file structure against pre-Phase-4 behavior.

- [ ] **Step 4: Full end-to-end smoke test**

1. Load Excel file → verify table + plot
2. Edit cells → verify debounced write
3. Wait 5s → verify DB sync
4. Switch to Sensory → create session → add samples → verify radar
5. Switch to Detailed Sensory → verify dual radar
6. Generate test report in each mode
7. Attach/view images
8. Open image inbox
9. Open database browser
10. Close file → reopen from recent files
11. Drag-drop a file
12. Close application → reopen → verify settings restored

- [ ] **Step 5: Mark Phase 4 complete**

---

## Phase 5: Frontend Polish

### Task 21: Add PlotWidget pixmap caching

**Files:**
- Modify: `DataViewer-Enterprise/src/plotting/PlotWidget.h`
- Modify: `DataViewer-Enterprise/src/plotting/PlotWidget.cpp`

- [ ] **Step 1: Add dirty flag to PlotWidget.h**

Add private member:

```cpp
    bool m_plotDirty = true;
```

- [ ] **Step 2: Set dirty flag on data changes**

In `PlotWidget.cpp`, set `m_plotDirty = true` in:
- `setSheetData()` (around line 147)
- `onSampleToggled()` (around line 54 — when checkbox state changes)
- `onPlotTypeChanged()` (when combo box selection changes)

- [ ] **Step 3: Guard renderCurrentPlot()**

In `updatePlot()` (around line 414), add early return:

```cpp
void PlotWidget::updatePlot()
{
    if (!m_plotDirty) return;
    m_plotDirty = false;
    renderCurrentPlot();
    // ... existing pixmap display code
}
```

- [ ] **Step 4: Build and test**

Verify: toggle checkboxes → plot updates. Resize window → plot doesn't re-render unnecessarily. Switch plot types → plot re-renders.

- [ ] **Step 5: Commit**

```bash
git add src/plotting/PlotWidget.h src/plotting/PlotWidget.cpp
git commit -m "perf: add dirty-flag caching to PlotWidget rendering"
```

---

### Task 22: Add RadarChartWidget polygon caching

**Files:**
- Modify: `DataViewer-Enterprise/src/ui/RadarChartWidget.h`
- Modify: `DataViewer-Enterprise/src/ui/RadarChartWidget.cpp`

- [ ] **Step 1: Add cached polygon storage**

In `RadarChartWidget.h`, add private members:

```cpp
    struct CachedPolygon {
        QPolygonF polygon;
        QColor    color;
        bool      visible = true;
    };
    QVector<CachedPolygon> m_cachedPolygons;
    bool m_polygonsDirty = true;
```

- [ ] **Step 2: Compute polygons on data change only**

In `setSessions()` (around line 32), set `m_polygonsDirty = true` and call `update()`.

Add a private method `recomputePolygons()` that computes the polygon vertices using the existing `axisPoint()` logic from `paintEvent()`. Move the polygon computation code from `paintEvent()` into this method.

- [ ] **Step 3: Use cached polygons in paintEvent()**

In `paintEvent()`, replace polygon computation with:

```cpp
if (m_polygonsDirty) {
    recomputePolygons();
    m_polygonsDirty = false;
}
// Draw from m_cachedPolygons
for (const auto& cp : m_cachedPolygons) {
    if (!cp.visible) continue;
    QColor fill = cp.color;
    fill.setAlpha(46);  // 18% alpha
    painter.setBrush(fill);
    painter.setPen(QPen(cp.color, 2));
    painter.drawPolygon(cp.polygon);
}
```

Note: Since polygon positions depend on widget size, `recomputePolygons()` needs the current rect. Set `m_polygonsDirty = true` in `resizeEvent()` too.

- [ ] **Step 4: Build and test**

Verify: radar chart renders identically. Toggle legend items → correct samples hide/show. Resize window → chart redraws correctly.

- [ ] **Step 5: Commit**

```bash
git add src/ui/RadarChartWidget.h src/ui/RadarChartWidget.cpp
git commit -m "perf: cache radar chart polygon geometry"
```

---

### Task 23: Signal safety and PlotEngine tick pre-computation

**Files:**
- Modify: `DataViewer-Enterprise/src/MainWindow.cpp` (blockSignals audit)
- Modify: `DataViewer-Enterprise/src/plotting/PlotEngine.cpp`

- [ ] **Step 1: Audit and add blockSignals guards**

Search MainWindow.cpp for all `setCurrentIndex()`, `setCurrentText()`, `setValue()`, and `setItem()` calls on combo boxes, spin boxes, and tables. Ensure each programmatic update is wrapped:

```cpp
m_sheetCombo->blockSignals(true);
m_sheetCombo->setCurrentIndex(idx);
m_sheetCombo->blockSignals(false);
```

Key locations to check:
- `populateSheetCombo()` — already uses blockSignals ✓
- `onFileSelected()` — check if sheet combo is guarded
- `displayCurrentSample()` — check if data table cellChanged is blocked during rebuild

- [ ] **Step 2: Pre-compute axis ticks in PlotEngine**

In `PlotEngine.cpp`, in `renderLinePlot()` (around line 300), compute tick values once and pass to grid/axes/legend:

```cpp
// Compute ticks once
double yStep = niceStep(yMax - yMin);
double xStep = niceStep(xMax - xMin);
QVector<double> yTicks, xTicks;
for (double v = yMin; v <= yMax + yStep * 0.01; v += yStep) yTicks << v;
for (double v = xMin; v <= xMax + xStep * 0.01; v += xStep) xTicks << v;

// Pass to drawGrid, drawAxes (instead of recomputing niceStep in each)
```

Modify `drawGrid()` and `drawAxes()` to accept pre-computed tick vectors instead of calling `niceStep()` internally.

- [ ] **Step 3: Build and test**

Verify: plots render identically. Check axis labels and grid lines match previous output.

- [ ] **Step 4: Commit**

```bash
git add src/MainWindow.cpp src/plotting/PlotEngine.cpp
git commit -m "perf: blockSignals guards, pre-compute PlotEngine ticks"
```

---

### Task 24: Phase 5 Testing Gate

- [ ] **Step 1: Run all 10 test suites**

All must pass.

- [ ] **Step 2: Visual comparison**

Screenshot plots before and after Phase 5. Compare visually — should be pixel-identical.

- [ ] **Step 3: Stress test**

Rapidly:
1. Toggle 10+ sample checkboxes in quick succession
2. Switch plot types back and forth
3. Resize window while chart is rendering
4. Switch modes rapidly (TPM → Sensory → Detailed → TPM)

No crashes, no freezes, no visual glitches.

- [ ] **Step 4: Full final smoke test**

Complete end-to-end test of every feature (same checklist as Task 20, Step 4).

- [ ] **Step 5: Mark Phase 5 complete — optimization plan done**

---

## File Structure Summary

### New files created:
```
src/controllers/ReportController.h
src/controllers/ReportController.cpp
src/controllers/ImageController.h
src/controllers/ImageController.cpp
src/controllers/ModeController.h
src/controllers/ModeController.cpp
src/controllers/FileController.h
src/controllers/FileController.cpp
src/controllers/ExcelController.h
src/controllers/ExcelController.cpp
src/controllers/DbSyncController.h
src/controllers/DbSyncController.cpp
src/utils/InputValidator.h
DataViewer/network_config.json.example
```

### Files modified:
```
DataViewerEnterprise.pro         — compiler flags, new source files
installer.iss                    — compression upgrade
src/main.cpp                     — remove hardcoded dev path
src/MainWindow.h                 — remove extracted slots/members, add controller pointers
src/MainWindow.cpp               — remove extracted method bodies, add controller wiring
src/ExcelReader.h                — static cached script path
src/ExcelReader.cpp              — script caching, Python validation, temp permissions
src/reporting/PptxWriter.cpp     — XML single-quote escaping
src/database/DatabaseManager.cpp — BLOB size limit, image validation
src/plotting/PlotWidget.h/cpp    — dirty-flag caching
src/ui/RadarChartWidget.h/cpp    — polygon caching
src/plotting/PlotEngine.cpp      — pre-computed tick vectors
src/ui/SensoryPanel.cpp          — input validation
src/ui/DetailedSensoryPanel.cpp  — input validation
DataViewer/database_explorer.py  — parameterized SQL
DataViewer/database_manager.py   — externalized config
```

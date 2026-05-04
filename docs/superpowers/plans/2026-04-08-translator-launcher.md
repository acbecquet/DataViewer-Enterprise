# Document Translator Launcher — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Translator" ribbon button to the Tools tab that launches the bundled Document Translator app, pre-populating its Anthropic API key from `anthropic_api_key.txt` if present.

**Architecture:** A new ribbon button in `buildToolsTab()` connects to a new private slot `onLaunchTranslator()`. The slot reads the API key from a file, writes the translator's `config.dat`, and fires off `QProcess::startDetached()`. The translator directory is copied into the project and shipped as an optional Inno Setup component.

**Tech Stack:** C++17, Qt 6.10, QProcess, QJsonDocument, QInputDialog, Inno Setup

---

## File Map

| File | Change |
|------|--------|
| `dataviewer_translator/` | New dir — full copy of translator `dev v2` |
| `src/MainWindow.h` | Add `onLaunchTranslator()` private slot + `writeTranslatorConfig()` private static |
| `src/MainWindow.cpp` | Add ribbon button in `buildToolsTab()`; implement slot and helper |
| `installer.iss` | Add `[Types]`/`[Components]`; tag all existing `[Files]` entries; add translator file entries |

---

## Task 1: Copy the Translator Directory

**Files:**
- Create: `dataviewer_translator/` (from `C:\Users\S1134987\Documents\Python\translator\dev v2\`)

- [ ] **Step 1: Copy the directory**

Run from the DataViewer-Enterprise project root:
```bash
cp -r "C:/Users/S1134987/Documents/Python/translator/dev v2/." dataviewer_translator/
```

- [ ] **Step 2: Verify the exe is present**

```bash
ls dataviewer_translator/dist/DocumentTranslator.exe
```
Expected: file listed with non-zero size.

- [ ] **Step 3: Commit**

```bash
git add dataviewer_translator/
git commit -m "chore: bundle Document Translator as dataviewer_translator"
```

---

## Task 2: Declare Slot and Helper in MainWindow.h

**Files:**
- Modify: `src/MainWindow.h`

- [ ] **Step 1: Add the slot declaration**

In `src/MainWindow.h`, locate the `// ── Image Inbox ──` block (around line 119). Add the new slot in the same `private slots:` section directly after `onOpenImageInbox`:

```cpp
    // ── Document Translator ──
    void onLaunchTranslator();
```

- [ ] **Step 2: Add the static helper declaration**

In `src/MainWindow.h`, locate the `private:` section (around line 122). Add this just before the `// ── File type detection` comment:

```cpp
    static bool writeTranslatorConfig(const QString& apiKey);
```

- [ ] **Step 3: Verify the file compiles**

```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
mingw32-make -f Makefile.Release 2>&1 | head -20
```
Expected: no errors related to MainWindow.h.

- [ ] **Step 4: Commit**

```bash
git add src/MainWindow.h
git commit -m "feat: declare onLaunchTranslator slot and writeTranslatorConfig helper"
```

---

## Task 3: Implement writeTranslatorConfig

**Files:**
- Modify: `src/MainWindow.cpp`

- [ ] **Step 1: Add required includes at the top of MainWindow.cpp**

Locate the existing `#include` block at the top of `src/MainWindow.cpp`. Add these if not already present:

```cpp
#include <QJsonDocument>
#include <QJsonObject>
#include <QDir>
#include <QFile>
#include <QTextStream>
#include <QStandardPaths>
```

- [ ] **Step 2: Implement the helper**

Find the end of the file (or a logical grouping near the other tool-related slots). Add this static method:

```cpp
bool MainWindow::writeTranslatorConfig(const QString& apiKey)
{
    QString configDir  = QDir::homePath() + "/.document_translator";
    QString configPath = configDir + "/config.dat";

    if (!QDir().mkpath(configDir))
        return false;

    QByteArray encoded = apiKey.trimmed().toUtf8().toBase64();
    QJsonObject obj;
    obj["api_key"] = QString::fromUtf8(encoded);
    QJsonDocument doc(obj);

    QFile file(configPath);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Truncate | QIODevice::Text))
        return false;

    file.write(doc.toJson(QJsonDocument::Compact));
    return true;
}
```

- [ ] **Step 3: Build and verify no errors**

```bash
mingw32-make -f Makefile.Release 2>&1 | grep -i error
```
Expected: no output (no errors).

- [ ] **Step 4: Manually verify the output format**

Open Qt Creator or any C++ REPL and confirm that for input `"sk-ant-test123"`, `toBase64()` produces `"c2stYW50LXRlc3QxMjM="`, and the JSON would be:
```json
{"api_key":"c2stYW50LXRlc3QxMjM="}
```
Cross-check using Python:
```python
import base64, json
key = "sk-ant-test123"
print(json.dumps({"api_key": base64.b64encode(key.encode()).decode()}))
# {"api_key": "c2stYW50LXRlc3QxMjM="}
```

- [ ] **Step 5: Commit**

```bash
git add src/MainWindow.cpp
git commit -m "feat: implement writeTranslatorConfig static helper"
```

---

## Task 4: Add Ribbon Button in buildToolsTab

**Files:**
- Modify: `src/MainWindow.cpp` (around line 276, end of `buildToolsTab`)

- [ ] **Step 1: Add the new group and button**

In `src/MainWindow.cpp`, locate `MainWindow::buildToolsTab` (line 252). The function currently ends at line 276 with the closing `}`. Insert the new group **before** that closing brace:

```cpp
    auto* extGrp = tab->addGroup("External Tools");
    auto* translatorBtn = extGrp->addLargeButton("Translator",
        style()->standardIcon(QStyle::SP_FileDialogDetailedView),
        "Open Document Translator");
    connect(translatorBtn, &QToolButton::clicked, this, &MainWindow::onLaunchTranslator);
```

The complete end of `buildToolsTab` should now read:

```cpp
    auto* imgGrp = tab->addGroup("Images");
    m_inboxBtn = imgGrp->addLargeButton("Images",
        style()->standardIcon(QStyle::SP_DirOpenIcon),
        "Open Image Inbox to assign photos to samples");
    connect(m_inboxBtn, &QToolButton::clicked, this, &MainWindow::onOpenImageInbox);

    auto* extGrp = tab->addGroup("External Tools");
    auto* translatorBtn = extGrp->addLargeButton("Translator",
        style()->standardIcon(QStyle::SP_FileDialogDetailedView),
        "Open Document Translator");
    connect(translatorBtn, &QToolButton::clicked, this, &MainWindow::onLaunchTranslator);
}
```

- [ ] **Step 2: Build**

```bash
mingw32-make -f Makefile.Release 2>&1 | grep -i error
```
Expected: no errors.

- [ ] **Step 3: Run the app and verify the button appears**

Launch `release/DataViewer.exe`. Navigate to the **Tools** ribbon tab. Confirm a new "External Tools" group appears with a "Translator" button. Clicking it will crash or do nothing yet (slot not implemented) — that is expected.

- [ ] **Step 4: Commit**

```bash
git add src/MainWindow.cpp
git commit -m "feat: add Translator button to Tools ribbon External Tools group"
```

---

## Task 5: Implement onLaunchTranslator Slot

**Files:**
- Modify: `src/MainWindow.cpp`

- [ ] **Step 1: Add required includes (if not already added in Task 3)**

Ensure `src/MainWindow.cpp` includes:
```cpp
#include <QProcess>
#include <QInputDialog>
#include <QMessageBox>
#include <QCoreApplication>
```

- [ ] **Step 2: Implement the slot**

Add the following method to `src/MainWindow.cpp` (near the other tool slots, e.g. after `onOpenImageInbox`):

```cpp
void MainWindow::onLaunchTranslator()
{
    // 1. Resolve the exe path relative to DataViewer.exe location
    QString exePath = QCoreApplication::applicationDirPath()
                      + "/dataviewer_translator/dist/DocumentTranslator.exe";

    if (!QFile::exists(exePath)) {
        QMessageBox::information(this, "Document Translator Not Installed",
            "The Document Translator is not installed.\n\n"
            "Re-run the DataViewer installer and select the "
            "\"Document Translator\" component.");
        return;
    }

    // 2. Get API key — from file if available, otherwise prompt
    QString apiKey;
    QString keyFilePath = QCoreApplication::applicationDirPath()
                          + "/anthropic_api_key.txt";
    QFile keyFile(keyFilePath);
    if (keyFile.open(QIODevice::ReadOnly | QIODevice::Text)) {
        apiKey = QTextStream(&keyFile).readAll().trimmed();
    }

    if (apiKey.isEmpty()) {
        bool ok = false;
        apiKey = QInputDialog::getText(
            this,
            "Anthropic API Key Required",
            "Enter your Anthropic API key to use the Document Translator:",
            QLineEdit::Normal, QString(), &ok);
        if (!ok || apiKey.trimmed().isEmpty())
            return;
    }

    // 3. Write translator config
    writeTranslatorConfig(apiKey);

    // 4. Launch — fire and forget
    if (!QProcess::startDetached(exePath)) {
        QMessageBox::warning(this, "Launch Failed",
            "Could not launch Document Translator.\n"
            "Path: " + exePath);
    }
}
```

- [ ] **Step 3: Build**

```bash
mingw32-make -f Makefile.Release 2>&1 | grep -i error
```
Expected: no errors.

- [ ] **Step 4: End-to-end test**

a. Ensure `dataviewer_translator/dist/DocumentTranslator.exe` exists in the build output directory OR temporarily copy it to `release/dataviewer_translator/dist/DocumentTranslator.exe` for testing.

b. Place a file `release/anthropic_api_key.txt` containing your real API key (plain text, one line).

c. Launch `release/DataViewer.exe`, go to Tools → Translator button.

d. Confirm:
- Document Translator window opens in the foreground
- No API key dialog appears (it was pre-filled from the file)
- `~/.document_translator/config.dat` exists and contains valid JSON

e. Remove `release/anthropic_api_key.txt` and repeat:
- Confirm the "Enter API key" dialog appears
- After entry, the translator still launches

f. Test the "not installed" path: rename `DocumentTranslator.exe` temporarily and click the button — confirm the information dialog appears.

- [ ] **Step 5: Commit**

```bash
git add src/MainWindow.cpp
git commit -m "feat: implement onLaunchTranslator — key injection and detached launch"
```

---

## Task 6: Update installer.iss

**Files:**
- Modify: `installer.iss`

- [ ] **Step 1: Add [Types] and [Components] sections**

Open `installer.iss`. After the `[Languages]` section (line 20), insert:

```ini
[Types]
Name: "full";    Description: "Full installation"
Name: "compact"; Description: "Compact installation"
Name: "custom";  Description: "Custom installation"; Flags: iscustom

[Components]
Name: "main";       Description: "DataViewer Enterprise (required)"; Types: full compact custom; Flags: fixed
Name: "translator"; Description: "Document Translator";              Types: full
```

- [ ] **Step 2: Tag all existing [Files] entries with Components: main**

Every existing line in the `[Files]` section needs `Components: main` appended. The complete updated `[Files]` section should read:

```ini
[Files]
; Main executable
Source: "release\DataViewer.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; Qt DLLs
Source: "release\Qt6Core.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Gui.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Widgets.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Sql.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Network.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Svg.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; MinGW runtime
Source: "release\libgcc_s_seh-1.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\libstdc++-6.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\libwinpthread-1.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; Qt plugins
Source: "release\platforms\*"; DestDir: "{app}\platforms"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\sqldrivers\qsqlite.dll"; DestDir: "{app}\sqldrivers"; Flags: ignoreversion; Components: main
Source: "release\imageformats\*"; DestDir: "{app}\imageformats"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\styles\*"; DestDir: "{app}\styles"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\iconengines\*"; DestDir: "{app}\iconengines"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\tls\*"; DestDir: "{app}\tls"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\networkinformation\*"; DestDir: "{app}\networkinformation"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\generic\*"; DestDir: "{app}\generic"; Flags: ignoreversion recursesubdirs; Components: main

; Bundled Python
Source: "release\python_bundle.zip"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; Resources
Source: "resources\templates\*"; DestDir: "{app}\resources\templates"; Flags: ignoreversion recursesubdirs; Components: main
Source: "resources\images\*"; DestDir: "{app}\resources\images"; Flags: ignoreversion recursesubdirs; Components: main
Source: "resources\sops.xlsx"; DestDir: "{app}\resources"; Flags: ignoreversion; Components: main

; Document Translator (optional component)
Source: "dataviewer_translator\dist\DocumentTranslator.exe"; DestDir: "{app}\dataviewer_translator\dist"; Flags: ignoreversion; Components: translator
Source: "dataviewer_translator\resources\*"; DestDir: "{app}\dataviewer_translator\resources"; Flags: ignoreversion recursesubdirs; Components: translator
```

- [ ] **Step 3: Verify the installer compiles**

Open Inno Setup Compiler and compile `installer.iss` (or run from command line if ISCC is on PATH):
```bash
"C:/Program Files (x86)/Inno Setup 6/ISCC.exe" installer.iss
```
Expected: `Successful compile (0 error(s))` and `dist/DataViewer-setup.exe` generated.

- [ ] **Step 4: Test the installer**

Run `dist/DataViewer-setup.exe`. On the component selection screen, confirm:
- "DataViewer Enterprise (required)" is checked and fixed
- "Document Translator" is checked by default (it's in the `full` type)
- Unchecking Document Translator and installing confirms `dataviewer_translator\` is absent from the install directory
- A full install confirms `dataviewer_translator\dist\DocumentTranslator.exe` is present

- [ ] **Step 5: Commit**

```bash
git add installer.iss
git commit -m "feat: add Document Translator as optional Inno Setup component"
```

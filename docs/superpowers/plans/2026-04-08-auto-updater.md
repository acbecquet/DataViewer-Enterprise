# Auto-Updater Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Synology Drive-based auto-updater that checks for new versions on startup and hourly, shows a dismissible dialog when one is found, and persists suppression choices across sessions.

**Architecture:** A standalone `UpdateChecker` QObject owns a `QTimer`, scans the Synology `Current/` folder for version subfolders, compares against the running version via `QApplication::applicationVersion()`, and shows a `QDialog` with Update/Remind/Later options. Suppression is stored in QSettings. MainWindow constructs it and calls `start()` after `setupUI()`.

**Tech Stack:** C++17, Qt 6.10 (QTimer, QVersionNumber, QSettings, QDialog, QProcess)

---

## File Map

| File | Change |
|------|--------|
| `src/main.cpp` | Update `setApplicationVersion` from `"0.7.1"` to `"0.8.0"` |
| `src/utils/UpdateChecker.h` | New — class declaration |
| `src/utils/UpdateChecker.cpp` | New — full implementation |
| `src/MainWindow.h` | Add `#include`, add `UpdateChecker* m_updateChecker` member |
| `src/MainWindow.cpp` | Construct and `start()` UpdateChecker after `setupUI()` |
| `DataViewerEnterprise.pro` | Add UpdateChecker to SOURCES + HEADERS |

---

## Task 1: Fix Application Version in main.cpp

**Files:**
- Modify: `src/main.cpp:17`

- [ ] **Step 1: Update the version string**

In `src/main.cpp` line 17, change:
```cpp
app.setApplicationVersion("0.7.1");
```
to:
```cpp
app.setApplicationVersion("0.8.0");
```

- [ ] **Step 2: Build to verify no errors**

```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
"C:/Qt/Tools/mingw1310_64/bin/mingw32-make.exe" -f Makefile.Release CXX="C:/Qt/Tools/mingw1310_64/bin/g++.exe" LINKER="C:/Qt/Tools/mingw1310_64/bin/g++.exe" 2>&1 | grep -i error
```
Expected: no output (no errors).

- [ ] **Step 3: Commit**

```bash
git add src/main.cpp
git commit -m "chore: bump applicationVersion to 0.8.0"
```

---

## Task 2: Create UpdateChecker.h

**Files:**
- Create: `src/utils/UpdateChecker.h`

- [ ] **Step 1: Create the header file**

Create `src/utils/UpdateChecker.h` with this exact content:

```cpp
#pragma once

#include <QObject>
#include <QTimer>
#include <QVersionNumber>

namespace DVE {

class UpdateChecker : public QObject
{
    Q_OBJECT

public:
    explicit UpdateChecker(QObject* parent = nullptr);

    /// Call once after the main window is shown. Runs an immediate check
    /// then starts the hourly timer.
    void start();

private slots:
    void check();

private:
    static QString      updateRoot();
    static QVersionNumber latestAvailable(QString* installerPathOut = nullptr);
    bool isSuppressed() const;
    void showDialog(const QVersionNumber& latest, const QString& installerPath);

    QTimer* m_timer;
    bool    m_suppressedThisSession = false;
};

} // namespace DVE
```

- [ ] **Step 2: Commit**

```bash
git add src/utils/UpdateChecker.h
git commit -m "feat: add UpdateChecker header"
```

---

## Task 3: Implement UpdateChecker.cpp

**Files:**
- Create: `src/utils/UpdateChecker.cpp`

- [ ] **Step 1: Create the implementation file**

Create `src/utils/UpdateChecker.cpp` with this exact content:

```cpp
#include "UpdateChecker.h"

#include <QApplication>
#include <QComboBox>
#include <QDate>
#include <QDialog>
#include <QDir>
#include <QFile>
#include <QHBoxLayout>
#include <QLabel>
#include <QMessageBox>
#include <QProcess>
#include <QPushButton>
#include <QSettings>
#include <QVBoxLayout>

namespace DVE {

static const int   kCheckIntervalMs = 60 * 60 * 1000; // 1 hour
static const char  kSettingsKey[]   = "update/suppressedUntil";

// ── Construction ─────────────────────────────────────────────────────────────

UpdateChecker::UpdateChecker(QObject* parent)
    : QObject(parent)
    , m_timer(new QTimer(this))
{
    m_timer->setInterval(kCheckIntervalMs);
    connect(m_timer, &QTimer::timeout, this, &UpdateChecker::check);
}

void UpdateChecker::start()
{
    check();          // immediate check on startup
    m_timer->start(); // then every hour
}

// ── Version discovery ─────────────────────────────────────────────────────────

QString UpdateChecker::updateRoot()
{
    return QDir::homePath()
           + "/SynologyDrive/SDR/Device Group/Software Release/Current";
}

QVersionNumber UpdateChecker::latestAvailable(QString* installerPathOut)
{
    QDir root(updateRoot());
    if (!root.exists())
        return {};

    QVersionNumber best;
    QString        bestInstaller;

    for (const QString& entry : root.entryList(QDir::Dirs | QDir::NoDotAndDotDot)) {
        QVersionNumber v = QVersionNumber::fromString(entry);
        if (v.isNull())
            continue;

        QString installer = root.filePath(entry + "/DataViewer-setup.exe");
        if (!QFile::exists(installer))
            continue;

        if (v > best) {
            best         = v;
            bestInstaller = installer;
        }
    }

    if (installerPathOut)
        *installerPathOut = bestInstaller;
    return best;
}

// ── Suppression ───────────────────────────────────────────────────────────────

bool UpdateChecker::isSuppressed() const
{
    QSettings s;
    const QString val = s.value(kSettingsKey).toString();
    if (val.isEmpty())
        return false;

    const QDate until = QDate::fromString(val, Qt::ISODate);
    if (!until.isValid())
        return false;

    if (QDate::currentDate() <= until)
        return true;

    // Past date — clean up stale entry
    QSettings().remove(kSettingsKey);
    return false;
}

// ── Check logic ───────────────────────────────────────────────────────────────

void UpdateChecker::check()
{
    if (m_suppressedThisSession)
        return;

    if (!QDir(updateRoot()).exists())
        return;

    QString        installerPath;
    QVersionNumber latest = latestAvailable(&installerPath);
    if (latest.isNull() || installerPath.isEmpty())
        return;

    const QVersionNumber current =
        QVersionNumber::fromString(QApplication::applicationVersion());
    if (QVersionNumber::compare(latest, current) <= 0)
        return;

    if (isSuppressed())
        return;

    showDialog(latest, installerPath);
}

// ── Dialog ────────────────────────────────────────────────────────────────────

void UpdateChecker::showDialog(const QVersionNumber& latest,
                               const QString&         installerPath)
{
    QDialog dlg;
    dlg.setWindowTitle("Update Available");
    dlg.setWindowFlags(dlg.windowFlags() & ~Qt::WindowContextHelpButtonHint);
    dlg.setMinimumWidth(420);

    QVBoxLayout* vl = new QVBoxLayout(&dlg);
    vl->setSpacing(8);
    vl->setContentsMargins(16, 16, 16, 12);

    // Message
    QLabel* msg = new QLabel(
        QString("DataViewer Enterprise <b>v%1</b> is available.<br>"
                "You are running v%2.")
            .arg(latest.toString(), QApplication::applicationVersion()));
    msg->setTextFormat(Qt::RichText);
    vl->addWidget(msg);
    vl->addSpacing(4);

    // Remind-in-X-days row
    QHBoxLayout* suppressRow = new QHBoxLayout;
    suppressRow->addWidget(new QLabel("Remind me in:"));
    QComboBox* daysPicker = new QComboBox;
    for (int d : {1, 3, 7, 14, 30})
        daysPicker->addItem(QString("%1 day%2").arg(d).arg(d == 1 ? "" : "s"), d);
    daysPicker->setCurrentIndex(2); // default: 7 days
    suppressRow->addWidget(daysPicker);
    suppressRow->addStretch();
    vl->addLayout(suppressRow);
    vl->addSpacing(8);

    // Buttons
    QHBoxLayout* btnRow = new QHBoxLayout;
    QPushButton* updateBtn = new QPushButton("Update Now");
    QPushButton* remindBtn = new QPushButton("Remind Me");
    QPushButton* laterBtn  = new QPushButton("Later This Session");
    updateBtn->setDefault(true);
    btnRow->addWidget(updateBtn);
    btnRow->addWidget(remindBtn);
    btnRow->addStretch();
    btnRow->addWidget(laterBtn);
    vl->addLayout(btnRow);

    // Update Now: launch installer, quit app
    connect(updateBtn, &QPushButton::clicked, &dlg, [&]() {
        dlg.accept();
        if (!QProcess::startDetached(installerPath)) {
            QMessageBox::warning(nullptr, "Update Failed",
                "Could not launch the installer.\nPath: " + installerPath);
            return;
        }
        QApplication::quit();
    });

    // Remind Me: persist suppression date, dismiss
    connect(remindBtn, &QPushButton::clicked, &dlg, [&]() {
        const int days = daysPicker->currentData().toInt();
        QSettings().setValue(
            kSettingsKey,
            QDate::currentDate().addDays(days).toString(Qt::ISODate));
        dlg.reject();
    });

    // Later: suppress for this session only
    connect(laterBtn, &QPushButton::clicked, &dlg, [&]() {
        m_suppressedThisSession = true;
        dlg.reject();
    });

    dlg.exec();
}

} // namespace DVE
```

- [ ] **Step 2: Commit**

```bash
git add src/utils/UpdateChecker.cpp
git commit -m "feat: implement UpdateChecker — version scan, suppression, dialog"
```

---

## Task 4: Register in .pro and Wire into MainWindow

**Files:**
- Modify: `DataViewerEnterprise.pro`
- Modify: `src/MainWindow.h`
- Modify: `src/MainWindow.cpp`

- [ ] **Step 1: Add UpdateChecker to DataViewerEnterprise.pro**

In `DataViewerEnterprise.pro`, find the line:
```
    src/utils/SingleInstance.cpp
```
Add the new source after it:
```
    src/utils/SingleInstance.cpp \
    src/utils/UpdateChecker.cpp
```

Find the line:
```
    src/utils/SingleInstance.h
```
Add the new header after it:
```
    src/utils/SingleInstance.h \
    src/utils/UpdateChecker.h
```

- [ ] **Step 2: Add include and member to MainWindow.h**

In `src/MainWindow.h`, find the existing includes block near the top (around line 34):
```cpp
#include <QFileSystemWatcher>
```
Add after it:
```cpp
#include "utils/UpdateChecker.h"
```

In `src/MainWindow.h`, find the private members section. Locate the existing timer members (around line 180-190, near `m_excelWriteTimer` or `m_dbSaveTimer`). Add:
```cpp
    UpdateChecker*  m_updateChecker = nullptr;
```

- [ ] **Step 3: Construct and start UpdateChecker in MainWindow.cpp**

In `src/MainWindow.cpp`, find the end of the constructor (around line 76):
```cpp
    updateStatusBar("Ready");
}
```
Insert the UpdateChecker startup just before `updateStatusBar`:
```cpp
    m_updateChecker = new UpdateChecker(this);
    m_updateChecker->start();

    updateStatusBar("Ready");
}
```

- [ ] **Step 4: Re-run qmake (new files added to .pro)**

```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
"C:/Qt/6.10.1/mingw_64/bin/qmake.exe" DataViewerEnterprise.pro -spec win32-g++ 2>&1 | grep -v "^Project MESSAGE"
```
Expected: no errors.

- [ ] **Step 5: Build**

```bash
"C:/Qt/Tools/mingw1310_64/bin/windres.exe" "--preprocessor=C:/Qt/Tools/mingw1310_64/bin/gcc.exe" "--preprocessor-arg=-E" "--preprocessor-arg=-xc-header" "--preprocessor-arg=-DRC_INVOKED" -i DataViewer_resource.rc -o release/DataViewer_resource_res.o --include-dir=. -DUNICODE -D_UNICODE -DWIN32

"C:/Qt/Tools/mingw1310_64/bin/mingw32-make.exe" -f Makefile.Release CXX="C:/Qt/Tools/mingw1310_64/bin/g++.exe" LINKER="C:/Qt/Tools/mingw1310_64/bin/g++.exe" 2>&1 | grep -i error
```
Expected: no output (no errors).

- [ ] **Step 6: Verify the update check runs on startup**

To smoke-test without waiting for Synology:
1. Temporarily change `kCheckIntervalMs` to `5000` (5 seconds) in `UpdateChecker.cpp`
2. Temporarily create folder `C:\Users\S1134987\SynologyDrive\SDR\Device Group\Software Release\Current\99.0.0\` and place any file named `DataViewer-setup.exe` inside it (can be an empty file)
3. Run `release\DataViewer.exe`
4. Confirm the "Update Available" dialog appears within a few seconds showing v99.0.0
5. Test all three buttons: Update Now (verify it tries to launch the installer), Remind Me (verify QSettings has `update/suppressedUntil`), Later This Session (verify dialog doesn't reappear after 5 more seconds)
6. Revert `kCheckIntervalMs` back to `60 * 60 * 1000` and delete the test folder

- [ ] **Step 7: Commit**

```bash
git add DataViewerEnterprise.pro src/MainWindow.h src/MainWindow.cpp
git commit -m "feat: wire UpdateChecker into MainWindow"
```

---

## Task 5: Build and Ship Installer

**Files:**
- No code changes

- [ ] **Step 1: Run windeployqt**

```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
PATH="C:/Qt/Tools/mingw1310_64/bin:C:/Qt/6.10.1/mingw_64/bin:$PATH" \
  "C:/Qt/6.10.1/mingw_64/bin/windeployqt6.exe" --release --no-translations release/DataViewer.exe 2>&1 | tail -5
```
Expected: DLL lines ending with "is up to date."

- [ ] **Step 2: Build installer**

```bash
"C:/Program Files (x86)/Inno Setup 6/ISCC.exe" installer.iss 2>&1 | tail -4
```
Expected: `Successful compile` with output at `dist/DataViewer-setup.exe`.

- [ ] **Step 3: Push**

```bash
git push origin main
```

- [ ] **Step 4: Copy installer to Synology**

Manually copy `dist/DataViewer-setup.exe` to:
`C:\Users\S1134987\SynologyDrive\SDR\Device Group\Software Release\Current\0.8.0\DataViewer-setup.exe`

This establishes the baseline. Future releases go in a new numbered subfolder.

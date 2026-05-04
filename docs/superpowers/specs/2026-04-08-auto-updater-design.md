# Auto-Updater — Design Spec
**Date:** 2026-04-08  
**Status:** Approved

---

## Overview

DataViewer Enterprise checks a shared Synology Drive folder for newer installer versions on startup and once per hour. If a newer version is found and not suppressed, a dialog offers the user the choice to update now, be reminded in X days, or dismiss for the session. Fully modular — silent if Synology is not mounted or no update exists.

---

## Scope

- New `UpdateChecker` class (`src/utils/UpdateChecker.h/.cpp`)
- Update available dialog (inline in UpdateChecker)
- Hourly `QTimer` + startup check
- Per-session and date-based suppression via QSettings
- Minimal changes to MainWindow (construct + start)
- No changes to installer or build process

---

## Architecture

### Update Source

Shared folder: `QDir::homePath() + "/SynologyDrive/SDR/Device Group/Software Release/Current/"`

Subfolder naming convention: `X.Y.Z/` (e.g. `0.9.0/`). Each folder contains `DataViewer-setup.exe`.

Maintainer workflow: drop `0.9.0/DataViewer-setup.exe` into `Current/` → all users get the prompt on next check.

### UpdateChecker Class

**Header:** `src/utils/UpdateChecker.h`
**Source:** `src/utils/UpdateChecker.cpp`

```
class UpdateChecker : public QObject
{
    Q_OBJECT
public:
    explicit UpdateChecker(QObject* parent = nullptr);
    void start();   // called once from MainWindow after UI is ready

private slots:
    void check();

private:
    static QString updateRoot();
    static QVersionNumber latestAvailable();
    static QString installerPath(const QVersionNumber& v);
    bool isSuppressed() const;
    void showDialog(const QVersionNumber& latest, const QString& installerPath);

    QTimer*  m_timer;
    bool     m_suppressedThisSession = false;
};
```

### Check Logic (`check()`)

```
1. If m_suppressedThisSession → return
2. If updateRoot() path does not exist → return
3. latestAvailable() → scan subfolders for valid X.Y.Z names, return highest
4. If latest <= QVersionNumber::fromString(QApplication::applicationVersion()) → return
5. If isSuppressed() (QSettings date check) → return
6. showDialog(latest, installerPath(latest))
```

### Version Comparison

Parse folder names with `QVersionNumber::fromString()`. Compare with `QVersionNumber::compare()`. Correctly handles `0.10.0 > 0.9.0`.

### Update Dialog

`QMessageBox`-style custom dialog (or `QDialog`) with:
- Title: `"Update Available"`
- Message: `"DataViewer Enterprise vX.Y.Z is available.\nYou are running vA.B.C."`
- **[Update Now]** → `QProcess::startDetached(installerPath)` → `QApplication::quit()`
- **[Remind me in X days]** → `QComboBox` with options {1, 3, 7, 14, 30} → writes `update/suppressedUntil = today + X days` to QSettings → close dialog
- **[Later this session]** → `m_suppressedThisSession = true` → close dialog

### Suppression

- **Session suppression:** `m_suppressedThisSession = true` — timer keeps running but dialog won't show until next app launch.
- **Date suppression:** QSettings key `update/suppressedUntil` stored as `"yyyy-MM-dd"`. `isSuppressed()` returns true if `QDate::currentDate() <= suppressedUntil`. Stale entries (past date) are ignored and cleaned up.

### Timer

`QTimer` interval: 60 minutes (`60 * 60 * 1000` ms). Started in `start()` after an immediate `check()` call. Single-shot first call ensures the startup check fires before the first hour elapses.

### MainWindow Integration

`MainWindow.h`: add `UpdateChecker* m_updateChecker = nullptr;`

`MainWindow.cpp` constructor (after `setupUI()`):
```cpp
m_updateChecker = new UpdateChecker(this);
m_updateChecker->start();
```

`DataViewerEnterprise.pro`: add `UpdateChecker.h` to HEADERS, `UpdateChecker.cpp` to SOURCES. Set `VERSION = 0.8.0` and `QApplication::setApplicationVersion("0.8.0")` in `main.cpp`.

---

## Application Version

`QApplication::applicationVersion()` must return `"0.8.0"` for comparison to work. Set in `main.cpp`:
```cpp
app.setApplicationVersion("0.8.0");
```

This is the only place the version needs to be set in code — `installer.iss` already has `AppVersion=0.8.0`.

---

## Error Handling

| Condition | Behaviour |
|-----------|-----------|
| Synology path not mounted / doesn't exist | Silent return |
| No valid version subfolders found | Silent return |
| Installer exe missing from version folder | Silent return (skip that version) |
| `QProcess::startDetached` fails | `QMessageBox::warning` with installer path |
| QSettings write fails | Session suppression still works (in-memory) |

---

## Files Changed

| File | Change |
|------|--------|
| `src/utils/UpdateChecker.h` | New |
| `src/utils/UpdateChecker.cpp` | New |
| `src/MainWindow.h` | Add `UpdateChecker* m_updateChecker` |
| `src/MainWindow.cpp` | Construct and start UpdateChecker after setupUI |
| `src/main.cpp` | Add `app.setApplicationVersion("0.8.0")` |
| `DataViewerEnterprise.pro` | Add UpdateChecker to HEADERS + SOURCES |

---

## Out of Scope

- Automatic download (installer is already on Synology Drive)
- Delta/patch updates
- Moving old version to Archive (manual maintainer step)
- Network/internet-based update checking

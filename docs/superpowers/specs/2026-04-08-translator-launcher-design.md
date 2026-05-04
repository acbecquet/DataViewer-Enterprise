# Document Translator Launcher — Design Spec
**Date:** 2026-04-08  
**Status:** Approved

---

## Overview

Add a "Translator" button to the DataViewer Enterprise Tools ribbon that launches the bundled Document Translator application. If an Anthropic API key file is present alongside DataViewer.exe, it is silently forwarded to the translator's config before launch. Otherwise the user is prompted for the key.

---

## Scope

- New ribbon button in `MainWindow::buildToolsTab()`
- New slot `MainWindow::onLaunchTranslator()`
- Copy of `dev v2` translator directory bundled as `dataviewer_translator/` in the project
- Inno Setup optional component for the translator
- No changes to the translator Python source

---

## Architecture

### 1. Bundled Translator Directory

Copy `C:\Users\S1134987\Documents\Python\translator\dev v2\` into the DataViewer-Enterprise project root as `dataviewer_translator\`. This directory is the installer source and contains the compiled `dist\DocumentTranslator.exe`.

**Runtime install location:** `{app}\dataviewer_translator\dist\DocumentTranslator.exe`

### 2. Ribbon Button

In `MainWindow::buildToolsTab()`, add a new `RibbonGroup` called **"External Tools"** after the existing groups. Add one large button:

- **Label:** `"Translator"`
- **Tooltip:** `"Open Document Translator"`
- **Icon:** `resourcePath() + "/images/translator_icon.png"` — fall back to a Qt standard pixmap (`QStyle::SP_FileDialogDetailedView`) if icon not present
- **Signal:** `QToolButton::clicked` → `MainWindow::onLaunchTranslator()`

### 3. Slot: `onLaunchTranslator()`

```
1. Build exe path:
     exePath = applicationDirPath() + "/dataviewer_translator/dist/DocumentTranslator.exe"

2. If exe does not exist:
     QMessageBox::information — "Document Translator is not installed.
     Re-run the DataViewer installer and select the Document Translator component."
     return

3. Read API key:
     keyFile = applicationDirPath() + "/anthropic_api_key.txt"
     if keyFile exists and is non-empty:
         apiKey = trimmed contents of keyFile
     else:
         apiKey = QInputDialog::getText("Anthropic API Key",
                      "Enter your Anthropic API key to use the Document Translator:")
         if user cancelled or apiKey is empty: return

4. Write translator config:
     configDir  = QDir::homePath() + "/.document_translator"
     configPath = configDir + "/config.dat"
     QDir().mkpath(configDir)
     encodedKey = QByteArray(apiKey.toUtf8()).toBase64()
     write JSON: {"api_key": "<encodedKey>"}  to configPath (overwrite)

5. Launch:
     if !QProcess::startDetached(exePath):
         QMessageBox::warning — "Failed to launch Document Translator."
```

### 4. Installer Changes (`installer.iss`)

Add `[Types]` and `[Components]` sections. Tag all existing `[Files]` entries with `Components: main`. Add translator files as a separate optional component.

**New sections:**
```ini
[Types]
Name: "full";    Description: "Full installation"
Name: "compact"; Description: "Compact installation"
Name: "custom";  Description: "Custom installation"; Flags: iscustom

[Components]
Name: "main";       Description: "DataViewer Enterprise (required)"; Types: full compact custom; Flags: fixed
Name: "translator"; Description: "Document Translator";              Types: full
```

**New file entries (translator component):**
```ini
Source: "dataviewer_translator\dist\DocumentTranslator.exe"; DestDir: "{app}\dataviewer_translator\dist"; Components: translator; Flags: ignoreversion
Source: "dataviewer_translator\resources\*";                 DestDir: "{app}\dataviewer_translator\resources"; Components: translator; Flags: ignoreversion recursesubdirs
```

All existing `[Files]` entries get `Components: main` appended.

---

## Error Handling

| Condition | Behaviour |
|-----------|-----------|
| `DocumentTranslator.exe` not installed | `QMessageBox::information` with reinstall instructions |
| Key file missing + user cancels dialog | Silent return (no launch) |
| `QProcess::startDetached` returns false | `QMessageBox::warning` |
| Config dir cannot be created | Launch anyway — translator will prompt for key itself |

---

## Files Changed

| File | Change |
|------|--------|
| `src/MainWindow.h` | Declare `onLaunchTranslator()` slot |
| `src/MainWindow.cpp` | Add ribbon button in `buildToolsTab()`; implement `onLaunchTranslator()` |
| `installer.iss` | Add `[Types]`, `[Components]`; tag existing files; add translator entries |
| `dataviewer_translator/` | New directory — copy of translator `dev v2` |

---

## Out of Scope

- C++ port of the translator (future project)
- Modifying the translator Python source
- Storing the API key in DataViewer's own settings

# Sensory Mode Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix 6 usability and data-integrity issues in sensory mode: scroll-wheel guard on spinboxes, Synology database path, session rename, wider sample cards, sample properties persistence, and improved radar legend in report exports.

**Architecture:** All changes are contained within the existing C++/Qt6 codebase. Tasks are independent — they touch different files and can be executed in any order. No new classes are needed except a `NoWheelSpinBox` helper and a `setReportMode` flag on `RadarChartWidget`.

**Tech Stack:** C++17, Qt 6.10.1, QXlsx, SQLite (via Qt SQL), PptxWriter (internal)

---

## File Map

| Task | Files Modified |
|------|---------------|
| 1 – Scroll guard | `src/ui/SensoryPanel.cpp` |
| 2 – Synology DB path | `src/MainWindow.cpp` |
| 3 – Session rename | `src/MainWindow.cpp`, `src/ui/SensoryPanel.h`, `src/ui/SensoryPanel.cpp` |
| 4 – Wider cards + legend labels | `src/ui/SensoryPanel.cpp`, `src/ui/RadarChartWidget.cpp` |
| 5 – Sample properties persist | `src/pipeline/SensoryData.h`, `src/database/DatabaseManager.cpp`, `src/ui/SensoryPanel.cpp`, `src/MainWindow.cpp` |
| 6 – Report legend top-right | `src/ui/RadarChartWidget.h`, `src/ui/RadarChartWidget.cpp`, `src/ui/SensoryPanel.cpp` |

---

## Task 1: Scroll Wheel Guard on Spinboxes

**Problem:** Scrolling the scroll area accidentally changes spinbox values. A user should have to click (focus) a spinbox before the scroll wheel can change its value.

**Files:**
- Modify: `src/ui/SensoryPanel.cpp` — add `NoWheelSpinBox` class before `SampleCard` constructor; replace `new QSpinBox` with `new NoWheelSpinBox`

**Background:** Qt's `QSpinBox` responds to wheel events regardless of focus. The fix is a minimal subclass that checks `hasFocus()` before processing wheel events. When the scroll area is scrolled, the spinboxes are under the cursor but not focused — so they correctly ignore the wheel.

- [ ] **Step 1: Add `#include <QWheelEvent>` to `SensoryPanel.cpp`**

  At the top of `src/ui/SensoryPanel.cpp`, add:
  ```cpp
  #include <QWheelEvent>
  ```
  (Qt 6 does not always transitively expose `QWheelEvent` via `<QSpinBox>` in strict builds.)

- [ ] **Step 2: Add `NoWheelSpinBox` class**

  In `src/ui/SensoryPanel.cpp`, add this class definition **before** the `SampleCard` constructor (around line 165, after the includes block). The class must be defined in the `DVE` namespace:

  ```cpp
  // Ignores wheel events unless the spinbox has keyboard focus.
  // Prevents accidental value changes when scrolling past sample cards.
  class NoWheelSpinBox : public QSpinBox
  {
  public:
      explicit NoWheelSpinBox(QWidget* parent = nullptr) : QSpinBox(parent) {
          setFocusPolicy(Qt::StrongFocus);
      }
  protected:
      void wheelEvent(QWheelEvent* e) override {
          if (hasFocus()) QSpinBox::wheelEvent(e);
          else e->ignore();
      }
  };
  ```

- [ ] **Step 3: Replace `new QSpinBox` with `new NoWheelSpinBox`**

  In `src/ui/SensoryPanel.cpp`, in the `SampleCard` constructor (around line 189), change:
  ```cpp
  auto* spin = new QSpinBox;
  ```
  to:
  ```cpp
  auto* spin = new NoWheelSpinBox;
  ```
  No other changes needed — `NoWheelSpinBox` inherits all `QSpinBox` methods and signals.

- [ ] **Step 4: Build and verify**

  Run: `cmake --build build --target DataViewer -j8`
  Expected: Clean build, no errors.
  Manual test: Open sensory mode, scroll over a sample card without clicking → values must not change. Click a spinbox to focus it → scroll wheel should now change the value.

- [ ] **Step 5: Commit**
  ```bash
  git add src/ui/SensoryPanel.cpp
  git commit -m "fix: prevent scroll wheel changing spinbox values without focus"
  ```

---

## Task 2: Synology Drive Database Path

**Problem:** `defaultDbPath()` currently checks a UNC network path (`//SDRNASUSA/...`) then falls back to local AppData. It should also check the Synology Drive mount at `~/SynologyDrive/SDR/Device Group/DVE_Database` — which is a local path, so WAL journal mode is appropriate for it.

**Files:**
- Modify: `src/MainWindow.cpp` — update `defaultDbPath()` (around lines 2394–2406)

**Background:** The Synology Drive desktop app syncs the NAS to a local folder at `C:\Users\<username>\SynologyDrive\...`. Using `QDir::homePath()` makes this portable across machines. The local sync folder is the **primary** path; local AppData is the only fallback if Synology Drive is not installed or synced. The old UNC NAS path (`//SDRNASUSA/...`) is removed entirely.

- [ ] **Step 1: Update `defaultDbPath()`**

  Replace the existing function body (lines 2394–2406 in `src/MainWindow.cpp`):

  ```cpp
  QString MainWindow::defaultDbPath() const
  {
      // 1. Synology Drive local sync folder (primary — offline-capable, per-user)
      QString synoDir = QDir::homePath() + "/SynologyDrive/SDR/Device Group/DVE_Database";
      if (QDir(synoDir).exists())
          return synoDir + "/dataviewer.db";

      // 2. Local AppData fallback (if Synology Drive is not installed/synced)
      QString dir = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
      QDir().mkpath(dir);
      return dir + "/dataviewer.db";
  }
  ```

  Note: The old NAS path (`//SDRNASUSA/...`) is removed entirely. The `DatabaseManager::open()` uses WAL mode for non-network paths (checks for `//` or `\\` prefix). The Synology path starts with neither, so WAL will correctly be used for it.

- [ ] **Step 2: Build and verify**

  Run: `cmake --build build --target DataViewer -j8`
  Expected: Clean build.
  Manual test: With the Synology Drive folder present at `C:\Users\S1134987\SynologyDrive\SDR\Device Group\DVE_Database`, launch the app and confirm the status bar or about dialog shows the Synology path as the active database.

- [ ] **Step 3: Commit**
  ```bash
  git add src/MainWindow.cpp
  git commit -m "feat: add Synology Drive path as secondary database location"
  ```

---

## Task 3: Session Rename in Navigator (Click-Pause-Click)

**Problem:** There is no way to rename sessions in the navigator list widget. The user wants click-pause-click inline renaming (standard Windows Explorer behavior).

**Files:**
- Modify: `src/MainWindow.cpp` — enable editable flag on navigator items; connect `itemChanged` signal
- Modify: `src/ui/SensoryPanel.h` — add `renameSession(int, QString)` declaration
- Modify: `src/ui/SensoryPanel.cpp` — implement `renameSession`

**Background:** `QListWidget` supports inline editing by adding `Qt::ItemIsEditable` to item flags. When the user edits and commits, `itemChanged` fires. The rename maps back to `testTitle` (since `sessionLabel()` uses `testTitle - testerName` as the display). The split on `" - "` to isolate `testTitle` is intentionally simple; it will misbehave if a test title itself contains ` - ` (e.g., "Study A - Phase 2"). This is acceptable for the current use case but is a **known limitation** — do not add extra complexity to handle it.

- [ ] **Step 1: Add `renameSession` declaration to `SensoryPanel.h`**

  In `src/ui/SensoryPanel.h`, inside the `SensoryPanel` class's public section (after `closeSessions`), add:
  ```cpp
  void renameSession(int index, const QString& newLabel);
  ```

- [ ] **Step 2: Implement `renameSession` in `SensoryPanel.cpp`**

  Add the implementation after `closeSessions` (around line 660):
  ```cpp
  void SensoryPanel::renameSession(int index, const QString& newLabel)
  {
      if (index < 0 || index >= m_sessions.size()) return;
      // The navigator label is "testTitle - testerName". We update testTitle only.
      // If the label contains " - ", split at the FIRST " - " to preserve testerName.
      QString title = newLabel;
      QString tester = m_sessions[index].testerName;
      int sep = newLabel.indexOf(" - ");
      if (sep >= 0 && !tester.isEmpty()) {
          title = newLabel.left(sep).trimmed();
      }
      m_sessions[index].testTitle = title;

      // If renaming the current session, update the header field too
      if (index == m_currentTesterIdx)
          m_testTitleEdit->setText(title);

      emit sessionsChanged();
  }
  ```

- [ ] **Step 3: Enable editable flag in `refreshSensoryNavigator()`**

  In `src/MainWindow.cpp`, in `refreshSensoryNavigator()` (around line 1743), replace:
  ```cpp
  m_sensoryNav->addItem(m_sensoryPanel->sessionLabel(sessions[i]));
  ```
  with:
  ```cpp
  auto* navItem = new QListWidgetItem(m_sensoryPanel->sessionLabel(sessions[i]));
  navItem->setFlags(navItem->flags() | Qt::ItemIsEditable);
  m_sensoryNav->addItem(navItem);
  ```

- [ ] **Step 4: Connect `itemChanged` to `renameSession`**

  In `src/MainWindow.cpp`, in the `setupConnections()` method (or wherever the other `m_sensoryNav` connections are made, around line 574), add:
  ```cpp
  connect(m_sensoryNav, &QListWidget::itemChanged, this, [this](QListWidgetItem* item) {
      if (!m_sensoryPanel) return;
      int row = m_sensoryNav->row(item);
      if (row < 0) return;
      // Block signals to prevent re-entry when refreshSensoryNavigator repopulates
      m_sensoryNav->blockSignals(true);
      m_sensoryPanel->renameSession(row, item->text());
      m_sensoryNav->blockSignals(false);
  });
  ```

- [ ] **Step 5: Build and verify**

  Run: `cmake --build build --target DataViewer -j8`
  Expected: Clean build.
  Manual test: In sensory mode, click a session in the navigator, wait ~0.5s, click again → the item enters edit mode. Type a new name and press Enter → the label updates and the session's test title changes.

- [ ] **Step 6: Commit**
  ```bash
  git add src/MainWindow.cpp src/ui/SensoryPanel.h src/ui/SensoryPanel.cpp
  git commit -m "feat: enable click-pause-click rename for sessions in navigator"
  ```

---

## Task 4: Widen Sample Cards and Legend Labels

**Problem:** Sample cards are too narrow (175px) — names get cut off. Legend label areas are too narrow (100px wide) — sample names in the radar chart legend get truncated.

**Files:**
- Modify: `src/ui/SensoryPanel.cpp` — `SampleCard` constructor line 171
- Modify: `src/ui/RadarChartWidget.cpp` — `drawLegend()` line 151

**Background:**
- Card width: `setFixedSize(175, 340)` → increase width by 50% → 263px. Height unchanged.
- Legend `entryWidth`: `100` → double to `200`. The swatch (12px) + spacing (3px) stays the same; only the text label area grows.

- [ ] **Step 1: Widen SampleCard by 50%**

  In `src/ui/SensoryPanel.cpp`, in the `SampleCard` constructor (line 171):
  ```cpp
  // Before:
  setFixedSize(175, 340);
  // After:
  setFixedSize(263, 340);
  ```

- [ ] **Step 2: Double legend entry width**

  In `src/ui/RadarChartWidget.cpp`, in `drawLegend()` (line 151):
  ```cpp
  // Before:
  const int entryWidth = 100;
  // After:
  const int entryWidth = 200;
  ```

- [ ] **Step 3: Build and verify**

  Run: `cmake --build build --target DataViewer -j8`
  Expected: Clean build.
  Manual test: Open sensory mode with several samples with longer names → cards should be wider with names fully visible. Check the radar chart legend → labels should have more room and not be cut off.

- [ ] **Step 4: Commit**
  ```bash
  git add src/ui/SensoryPanel.cpp src/ui/RadarChartWidget.cpp
  git commit -m "fix: widen sample cards by 50% and double legend label width"
  ```

---

## Task 5: Sample Properties — Input, Persist to JSON/XLSX/DB

**Problem:** The "Device Properties" section in the properties panel (Burn, Clog, Leak, Puff Time) shows empty read-only fields. These should be editable, stored in the session data, and automatically saved to the database, JSON exports, and XLSX exports.

**Files:**
- Modify: `src/pipeline/SensoryData.h` — add `burnStatus`, `clogStatus`, `leakStatus` to `SensorySession`
- Modify: `src/database/DatabaseManager.cpp` — serialize/deserialize new fields in JSON blob
- Modify: `src/MainWindow.cpp` — make Device Properties rows editable; wire `onPropCellChanged` for sensory mode
- Modify: `src/ui/SensoryPanel.cpp` — include new fields in `saveToJson()` and `saveToExcel()`

**Background:** Session data is stored as a JSON blob in the `sensory_sessions.json_data` column. Adding new fields to the JSON and reading them back is all that's needed for DB persistence — no schema migration required. The `puffLength` field already exists on `SensorySession`; Burn, Clog, and Leak are new.

- [ ] **Step 1: Add new fields to `SensorySession`**

  In `src/pipeline/SensoryData.h`, extend `SensorySession` (after `puffLength`):
  ```cpp
  struct SensorySession {
      // ... existing fields ...
      QString  puffLength;
      QString  burnStatus;   // NEW
      QString  clogStatus;   // NEW
      QString  leakStatus;   // NEW
      // ... rest of struct ...
  };
  ```

- [ ] **Step 2: Serialize new fields in `DatabaseManager::saveSensorySession()`**

  In `src/database/DatabaseManager.cpp`, in `saveSensorySession()` (around line 774), add after `root["puff_length"]`:
  ```cpp
  root["burn_status"] = s.burnStatus;
  root["clog_status"] = s.clogStatus;
  root["leak_status"] = s.leakStatus;
  ```

- [ ] **Step 3: Deserialize new fields — TWO separate functions to update**

  There are **two independent deserialization paths** in `src/database/DatabaseManager.cpp`. Both must be updated:

  **Path A — `loadSensorySessions()` (around line 920, bulk load):**
  Find the block that reads `sess.puffLength = ...` and add after it:
  ```cpp
  sess.burnStatus = obj.value("burn_status").toString();
  sess.clogStatus = obj.value("clog_status").toString();
  sess.leakStatus = obj.value("leak_status").toString();
  ```

  **Path B — `loadSensorySession(int id)` (around line 964, single-session load by id):**
  Find the corresponding `sess.puffLength = ...` line in this separate function and add the same three lines after it:
  ```cpp
  sess.burnStatus = obj.value("burn_status").toString();
  sess.clogStatus = obj.value("clog_status").toString();
  sess.leakStatus = obj.value("leak_status").toString();
  ```

  These are NOT the same code block — they are two separate functions. Missing Path B causes silent data loss when loading a session by id.

- [ ] **Step 4: Add new fields to `saveToJson()` in `SensoryPanel.cpp`**

  In `src/ui/SensoryPanel.cpp`, in `saveToJson()` (around line 841), `puffLength` is currently **not serialized** in this function (it is only serialized in `DatabaseManager::saveSensorySession`). Add all four fields — `puffLength` plus the three new ones — to the JSON root object in `saveToJson()`:
  ```cpp
  root["puff_length"] = sess.puffLength;   // was missing — add this
  root["burn_status"] = sess.burnStatus;
  root["clog_status"] = sess.clogStatus;
  root["leak_status"] = sess.leakStatus;
  ```

- [ ] **Step 5: Add new fields to `saveToExcel()` in `SensoryPanel.cpp`**

  In `src/ui/SensoryPanel.cpp`, in `saveToExcel()` (around line 868), in the metadata section below the data table, add rows for the three new fields alongside existing metadata (Test Title, Tester, etc.):
  ```cpp
  // Add after the existing metadata rows:
  int metaRow = /* row after last existing metadata row */ + 1;
  xlsx.write(metaRow, 1, "Burn Status");
  xlsx.write(metaRow, 2, sess.burnStatus);
  ++metaRow;
  xlsx.write(metaRow, 1, "Clog Status");
  xlsx.write(metaRow, 2, sess.clogStatus);
  ++metaRow;
  xlsx.write(metaRow, 1, "Leak Status");
  xlsx.write(metaRow, 2, sess.leakStatus);
  ```

- [ ] **Step 6: Make Device Properties rows editable in `updateSensoryProperties()`**

  In `src/MainWindow.cpp`, in `updateSensoryProperties()` (around lines 1810–1814), change the four `makeReadOnly` calls for Device Properties to `makeEditable` (which already exists as a local lambda elsewhere in the file):

  First, copy the `makeEditable` lambda into `updateSensoryProperties()` (it is defined in `updateProperties()` but not available here):
  ```cpp
  auto makeEditable = [&](int row, const QString& label, const QString& value) {
      QTableWidgetItem* lbl = new QTableWidgetItem(label);
      lbl->setFlags(Qt::ItemIsEnabled);
      lbl->setForeground(QColor(0x55, 0x55, 0x55));
      QFont f = lbl->font(); f.setBold(true); f.setPointSize(8); lbl->setFont(f);
      m_propTable->setItem(row, 0, lbl);
      QTableWidgetItem* val = new QTableWidgetItem(value);
      val->setFlags(Qt::ItemIsEnabled | Qt::ItemIsEditable | Qt::ItemIsSelectable);
      m_propTable->setItem(row, 1, val);
      m_propTable->setRowHeight(row, 20);
  };
  ```

  Then change:
  ```cpp
  // Before:
  makeReadOnly(8,  "Burn",                 "");
  makeReadOnly(9,  "Clog",                 "");
  makeReadOnly(10, "Leak",                 "");
  makeReadOnly(11, "Puff Time (est., s)",  "");

  // After:
  makeEditable(8,  "Burn",                sess->burnStatus);
  makeEditable(9,  "Clog",               sess->clogStatus);
  makeEditable(10, "Leak",               sess->leakStatus);
  makeEditable(11, "Puff Time (est., s)", sess->puffLength);
  ```

- [ ] **Step 7: Wire `onPropCellChanged` to save sensory properties**

  In `src/MainWindow.cpp`, find `MainWindow::onPropCellChanged()`. It currently handles TPM properties. Add sensory mode handling at the top:

  The `blockSignals(true/false)` guard must be placed **inside** the sensory branch only — if placed at the function level, it will never be unblocked when `col != 1` returns early in the TPM path, permanently silencing the table's signals.

  ```cpp
  void MainWindow::onPropCellChanged(int row, int col)
  {
      if (col != 1) return;  // only value column

      if (m_sensoryMode && m_sensoryPanel) {
          // Guard signals to prevent re-entry from DB save triggering updateSensoryProperties
          m_propTable->blockSignals(true);
          SensorySession* sess = m_sensoryPanel->currentSession();
          if (sess) {
              QTableWidgetItem* valItem = m_propTable->item(row, 0);
              QTableWidgetItem* dataItem = m_propTable->item(row, 1);
              if (valItem && dataItem) {
                  QString label = valItem->text();
                  QString value = dataItem->text().trimmed();
                  if (label == "Burn")                     sess->burnStatus = value;
                  else if (label == "Clog")                sess->clogStatus = value;
                  else if (label == "Leak")                sess->leakStatus = value;
                  else if (label == "Puff Time (est., s)") sess->puffLength = value;
                  // Auto-save to database
                  if (m_db->isOpen()) m_db->saveSensorySession(*sess);
              }
          }
          m_propTable->blockSignals(false);
          return;
      }

      // ... existing TPM handling continues below (unchanged) ...
  ```

- [ ] **Step 8: Build and verify**

  Run: `cmake --build build --target DataViewer -j8`
  Expected: Clean build.
  Manual test:
  1. Open sensory mode, select a session
  2. Double-click Burn/Clog/Leak/Puff Time in the properties panel → should become editable
  3. Enter a value and press Enter → value persists after switching sessions and back
  4. Save as JSON → open the JSON file and confirm burn_status/clog_status/leak_status are present
  5. Save as XLSX → confirm the metadata section includes the new rows

- [ ] **Step 9: Commit**
  ```bash
  git add src/pipeline/SensoryData.h src/database/DatabaseManager.cpp \
          src/ui/SensoryPanel.cpp src/MainWindow.cpp
  git commit -m "feat: add editable burn/clog/leak/puff properties to sensory sessions, persisted to JSON/XLSX/DB"
  ```

---

## Task 6: Radar Legend in Report Export (Top-Right, Larger Font)

**Problem:** When generating a PPTX report, the radar chart is rendered from `RadarChartWidget` using its standard `paintEvent`. The legend appears in a 40px strip at the bottom and uses 8pt font — too small to read in a presentation slide. For the report image, the legend should move to the top-right corner with a larger font.

**Files:**
- Modify: `src/ui/RadarChartWidget.h` — add `m_reportMode` flag and `setReportMode()`
- Modify: `src/ui/RadarChartWidget.cpp` — modify `paintEvent` to respect report mode
- Modify: `src/ui/SensoryPanel.cpp` — call `setReportMode(true)` on `tempChart` before render

**Background:** The report renders by creating a `RadarChartWidget tempChart`, calling `render()` on it, and saving the result as PNG. A `bool m_reportMode` flag switches the legend from the bottom strip to a vertical list in the top-right corner with 12pt font, giving it far more space and readability in the exported image.

- [ ] **Step 1: Add `setReportMode` to `RadarChartWidget.h`**

  In `src/ui/RadarChartWidget.h`, add to the `public` section:
  ```cpp
  void setReportMode(bool reportMode) { m_reportMode = reportMode; }
  ```

  Add to the `private` section:
  ```cpp
  bool m_reportMode = false;
  ```

- [ ] **Step 2: Modify `paintEvent` to support report mode legend**

  In `src/ui/RadarChartWidget.cpp`, in `paintEvent` (lines 226–264), change the layout calculation to support two modes:

  Replace the current `paintEvent` with:
  ```cpp
  void RadarChartWidget::paintEvent(QPaintEvent* /*event*/)
  {
      QPainter p(this);
      p.setRenderHint(QPainter::Antialiasing);
      p.fillRect(rect(), QColor(248, 248, 248));

      if (m_reportMode) {
          // Report mode: legend in top-right corner, chart uses full height
          const int legendW = qMax(220, width() / 3);
          QRectF chartArea(0, 0, width() - legendW, height());
          QRectF legendArea(width() - legendW, 0, legendW, height());

          double margin = 48.0;
          double radius = (qMin(chartArea.width(), chartArea.height()) / 2.0) - margin;
          if (radius < 20) return;

          QPointF center(chartArea.left() + chartArea.width() / 2.0,
                         chartArea.top()  + chartArea.height() / 2.0);

          drawGrid(p, center, radius);
          drawAxes(p, center, radius);

          int colorIdx = 0;
          for (const SensorySession& sess : m_sessions) {
              for (const SensorySample& sample : sess.samples) {
                  if (!m_hiddenSamples.contains(colorIdx))
                      drawSample(p, sample, center, radius, kColors[colorIdx % kColors.size()]);
                  ++colorIdx;
              }
          }

          // Legend separator (vertical)
          p.setPen(QColor(200, 200, 200));
          p.drawLine(QPointF(legendArea.left(), 0), QPointF(legendArea.left(), height()));

          drawLegendReport(p, legendArea);
      } else {
          // Normal UI mode: legend at bottom
          const int legendH = 40;
          QRectF chartArea(0, 0, width(), height() - legendH);
          QRectF legendArea(0, height() - legendH, width(), legendH);

          double margin = 48.0;
          double radius = (qMin(chartArea.width(), chartArea.height()) / 2.0) - margin;
          if (radius < 20) return;

          QPointF center(chartArea.left() + chartArea.width() / 2.0,
                         chartArea.top()  + chartArea.height() / 2.0);

          drawGrid(p, center, radius);
          drawAxes(p, center, radius);

          int colorIdx = 0;
          for (const SensorySession& sess : m_sessions) {
              for (const SensorySample& sample : sess.samples) {
                  if (!m_hiddenSamples.contains(colorIdx))
                      drawSample(p, sample, center, radius, kColors[colorIdx % kColors.size()]);
                  ++colorIdx;
              }
          }

          p.setPen(QColor(200, 200, 200));
          p.drawLine(QPointF(0, legendArea.top()), QPointF(width(), legendArea.top()));

          drawLegend(p, legendArea);
      }
  }
  ```

- [ ] **Step 3: Add `drawLegendReport()` to `RadarChartWidget`**

  In `src/ui/RadarChartWidget.h`, add private declaration:
  ```cpp
  void drawLegendReport(QPainter& p, const QRectF& legendRect);
  ```

  In `src/ui/RadarChartWidget.cpp`, add the implementation (after `drawLegend`):
  ```cpp
  void RadarChartWidget::drawLegendReport(QPainter& p, const QRectF& legendRect)
  {
      struct Entry { QString name; QColor color; int globalIdx; };
      QList<Entry> entries;
      int colorIdx = 0;
      for (const SensorySession& sess : m_sessions) {
          for (const SensorySample& sample : sess.samples) {
              entries.append({sample.name.isEmpty()
                                  ? QString("Sample %1").arg(colorIdx + 1)
                                  : sample.name,
                              kColors[colorIdx % kColors.size()],
                              colorIdx});
              ++colorIdx;
          }
      }
      if (entries.isEmpty()) return;

      const int swatchSize = 16;
      const int rowH       = 24;
      const int margin     = 10;

      QFont f = p.font();
      f.setPointSize(12);
      f.setBold(false);
      p.setFont(f);

      // Title
      QFont titleFont = f;
      titleFont.setBold(true);
      titleFont.setPointSize(13);
      p.setFont(titleFont);
      p.setPen(QColor(40, 40, 40));
      p.drawText(QRectF(legendRect.left() + margin, legendRect.top() + margin,
                        legendRect.width() - 2 * margin, rowH),
                 Qt::AlignLeft | Qt::AlignVCenter, "Samples");

      p.setFont(f);
      int y = static_cast<int>(legendRect.top()) + margin + rowH + 4;

      for (const Entry& e : entries) {
          bool hidden = m_hiddenSamples.contains(e.globalIdx);
          QColor swatchColor = hidden ? QColor(200, 200, 200) : e.color;
          QColor textColor   = hidden ? QColor(170, 170, 170) : QColor(40, 40, 40);

          // Swatch
          p.fillRect(static_cast<int>(legendRect.left()) + margin, y,
                     swatchSize, swatchSize, swatchColor);
          p.setPen(swatchColor.darker(150));
          p.drawRect(static_cast<int>(legendRect.left()) + margin, y, swatchSize, swatchSize);

          // Label — use full available width
          p.setPen(textColor);
          p.drawText(static_cast<int>(legendRect.left()) + margin + swatchSize + 6,
                     y,
                     static_cast<int>(legendRect.width()) - margin - swatchSize - 10,
                     rowH,
                     Qt::AlignVCenter | Qt::AlignLeft,
                     e.name);

          y += rowH;
          if (y + rowH > legendRect.bottom()) break;  // clip if too many entries
      }
  }
  ```

- [ ] **Step 4: Enable report mode in `generateCombinedPptx()`**

  In `src/ui/SensoryPanel.cpp`, in `generateCombinedPptx()` (around lines 1740–1748), add `setReportMode(true)` after creating the temp chart.

  **IMPORTANT:** `generateCombinedPptx` is a `static` method (declared `static` in `SensoryPanel.h` line 113). Do **not** remove the `static` qualifier. The `tempChart` is a stack-local instance — calling a non-static method on it is perfectly valid C++.

  ```cpp
  RadarChartWidget tempChart;
  tempChart.setSessions({sess});
  tempChart.setReportMode(true);   // <-- ADD THIS LINE
  tempChart.resize(720, 600);
  ```

- [ ] **Step 5: Build and verify**

  Run: `cmake --build build --target DataViewer -j8`
  Expected: Clean build.
  Manual test:
  1. In sensory mode with multiple samples, generate a PPTX report
  2. Open the PPTX → the radar chart image should show the legend in the top-right corner with 11pt text and a "Samples" title
  3. The radar chart itself should use the full height of the image
  4. The UI radar chart (in the app) should still show the legend at the bottom (unchanged)

- [ ] **Step 6: Commit**
  ```bash
  git add src/ui/RadarChartWidget.h src/ui/RadarChartWidget.cpp src/ui/SensoryPanel.cpp
  git commit -m "feat: move radar legend to top-right with larger font in PPTX report export"
  ```

---

## Build System Note

Per existing project conventions (see `feedback_build_system.md`):
- Qt 6.10.1 `rcc` binary has a TSD workaround — do not regenerate resources unless needed
- Use `cmake --build build --target DataViewer -j8` (not `make`) for all builds
- If phantom make errors occur on first run, run the build command a second time

## Final Verification

After all 6 tasks:
- [ ] All 6 commits are clean (`git log --oneline -6`)
- [ ] App launches and reaches sensory mode without crash
- [ ] Each fix verified manually per the test steps above

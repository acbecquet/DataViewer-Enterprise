# Post-Wave Bug-fix Batch Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship six small post-v2.0 fixes (#1, #3, #4, #5, #6, #7) as one focused branch with reviewable per-task commits.

**Architecture:** Each task is self-contained — a presentation tweak, a config flip, a struct add, or a localized investigation+fix. Order is risk-ascending: lowest-risk presentation tweaks first, then investigations, then the feature add. No task depends on another; tasks could be reordered if needed.

**Tech Stack:** C++17, Qt 6.10, qmake + MinGW 13.1, PostgreSQL 16 (samples stored in JSONB so no schema migration needed for #7), QtTest framework.

**Spec:** `docs/superpowers/specs/2026-05-15-postwave-bugfix-batch-design.md`

---

## Pre-flight

Before starting any task, confirm working directory and branch:

```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
git status                              # should be on feature/postwave-bugfix-batch, clean
git log --oneline main..HEAD            # should show 0 commits (fresh from main)
```

Decrypt any MIP-encrypted files before any C++ build attempt:

```bash
python tools/decrypt_via_copy.py --apply
```

Build verification (used in multiple tasks below):

```bash
mkdir -p build && cd build
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
C:\Qt\6.10.2\mingw_64\bin\qmake.exe -spec win32-g++ ../DataViewerEnterprise.pro
mingw32-make -j8
```

Test suite invocation: `tests\run-tests.ps1` (PowerShell). Single-class run: `tests\run-tests.ps1 -Filter tst_<name>`.

---

## Task 1: #5 — PresenceDotsDelegate right-anchors dots when text elides

**Files:**
- Modify: `src/widgets/PresenceDotsDelegate.cpp` (rewrite paint, adjust sizeHint)
- Create: `tests/widgets/tst_presencedotsdelegate.cpp`
- Modify: `tests/widgets/widgets.pro` (or whichever .pro hosts widget tests — verify in step 0)

### Step 0: Locate the test .pro file for widget tests

- [ ] **Find existing widget test .pro location**

```bash
grep -rl "PresenceAvatarBar\|widgets" tests/*.pro tests/**/*.pro 2>/dev/null
ls tests/widgets/ 2>/dev/null || ls tests/ui/ 2>/dev/null
```

If no widgets test subdir exists, create `tests/widgets/widgets.pro` mirroring an existing simple test like `tests/database/database.pro`. Add the new subdir to the parent `tests/tests.pro` SUBDIRS line.

### Step 1: Write the failing test

- [ ] **Create `tests/widgets/tst_presencedotsdelegate.cpp`**

```cpp
#include <QtTest>
#include <QStandardItemModel>
#include <QStyleOptionViewItem>
#include <QPixmap>
#include <QPainter>

#include "widgets/PresenceDotsDelegate.h"

using namespace DVE;

class TstPresenceDotsDelegate : public QObject
{
    Q_OBJECT

private slots:
    void dotsStayVisibleWhenTextElides();
    void noDotsRendersAsPlainText();

private:
    QPixmap renderItem(const QString& text, const QStringList& colors,
                       int columnWidth);
};

QPixmap TstPresenceDotsDelegate::renderItem(const QString& text,
                                            const QStringList& colors,
                                            int columnWidth)
{
    QStandardItemModel model;
    auto* item = new QStandardItem(text);
    item->setData(colors,  PresenceDotsDelegate::kColorsRole);
    item->setData(QStringList{"viewing"}, PresenceDotsDelegate::kIntentsRole);
    model.appendRow(item);

    PresenceDotsDelegate delegate;
    QStyleOptionViewItem opt;
    opt.rect = QRect(0, 0, columnWidth, 20);
    opt.font = QFont();
    opt.state = QStyle::State_Enabled;

    QPixmap pm(columnWidth, 20);
    pm.fill(Qt::white);
    QPainter painter(&pm);
    delegate.paint(&painter, opt, model.index(0, 0));
    painter.end();
    return pm;
}

void TstPresenceDotsDelegate::dotsStayVisibleWhenTextElides()
{
    // Long text in a narrow column should elide, but the colored dot
    // must still be visible at the right edge.
    QStringList colors{"#FF0000"};
    QPixmap pm = renderItem(
        "this_is_a_very_long_filename_that_will_definitely_elide.xlsx",
        colors, /*columnWidth=*/120);

    // Scan a small column at the rightmost 12px for any non-white pixel.
    // If the dot is right-anchored, at least one pixel in that band is red.
    const QImage img = pm.toImage();
    bool foundColored = false;
    for (int x = img.width() - 14; x < img.width() - 2 && !foundColored; ++x) {
        for (int y = 0; y < img.height(); ++y) {
            const QRgb px = img.pixel(x, y);
            if (qRed(px) > 200 && qGreen(px) < 50 && qBlue(px) < 50) {
                foundColored = true;
                break;
            }
        }
    }
    QVERIFY2(foundColored,
             "expected a red dot in the rightmost 12px band when text elides");
}

void TstPresenceDotsDelegate::noDotsRendersAsPlainText()
{
    // Empty colors list should produce no dots and not crash.
    QPixmap pm = renderItem("short.xlsx", {}, 200);
    QVERIFY(!pm.isNull());
}

QTEST_MAIN(TstPresenceDotsDelegate)
#include "tst_presencedotsdelegate.moc"
```

### Step 2: Create the test .pro stub (if not created in step 0)

- [ ] **Write `tests/widgets/tst_presencedotsdelegate.pro`**

Mirror the pattern of an existing leaf .pro file. Example skeleton (compare with `tests/database/tst_presencemanager.pro` for exact include paths):

```pro
QT += testlib gui widgets
CONFIG += console testcase
TEMPLATE = app

TARGET = tst_presencedotsdelegate

INCLUDEPATH += ../../src

SOURCES += \
    tst_presencedotsdelegate.cpp \
    ../../src/widgets/PresenceDotsDelegate.cpp

HEADERS += \
    ../../src/widgets/PresenceDotsDelegate.h
```

Add this subdir to the parent `tests/widgets/widgets.pro` SUBDIRS list (create that file if missing), and add `widgets` to `tests/tests.pro` SUBDIRS.

### Step 3: Run the test — verify it fails

- [ ] **Build and run the new test, confirm `dotsStayVisibleWhenTextElides` FAILS**

```powershell
.\tests\run-tests.ps1 -Filter tst_presencedotsdelegate
```

Expected: `dotsStayVisibleWhenTextElides` FAILS with "expected a red dot in the rightmost 12px band when text elides" because current `paint()` places dots after the text and they get clipped or overdrawn by the elision.

### Step 4: Rewrite `PresenceDotsDelegate::paint` to right-anchor the dots

- [ ] **Replace the body of `paint()` and `sizeHint()` in `src/widgets/PresenceDotsDelegate.cpp`**

Replace the entire function bodies (keep the namespace, kDotDiameter/kDotSpacing/kTextPadding/kEditingRingPx constants, and the constructor as they are):

```cpp
void PresenceDotsDelegate::paint(QPainter* painter,
                                 const QStyleOptionViewItem& option,
                                 const QModelIndex& index) const
{
    const QStringList colors  = index.data(kColorsRole).toStringList();
    const QStringList intents = index.data(kIntentsRole).toStringList();

    if (colors.isEmpty()) {
        // No presence — defer entirely to the default renderer.
        QStyledItemDelegate::paint(painter, option, index);
        return;
    }

    // Reserve a right-edge region for dots. Text gets the remaining space
    // and elides into it if too long.
    const int dotsTotal = colors.size() * kDotDiameter
                        + (colors.size() - 1) * kDotSpacing;
    const int dotsRegionWidth = dotsTotal + kTextPadding;

    QStyleOptionViewItem opt = option;
    opt.rect = QRect(option.rect.left(), option.rect.top(),
                     std::max(0, option.rect.width() - dotsRegionWidth),
                     option.rect.height());

    // Render the item (text, selection background, icon) into the
    // shortened rect. Qt's default elision logic now uses the shorter
    // width.
    QStyledItemDelegate::paint(painter, opt, index);

    // Paint dots at the right edge of the ORIGINAL option.rect.
    int dotX = option.rect.right() - dotsTotal;
    const int dotY = option.rect.center().y() - kDotDiameter / 2;

    painter->save();
    painter->setRenderHint(QPainter::Antialiasing, true);

    for (int i = 0; i < colors.size(); ++i) {
        QColor fill(colors[i]);
        if (!fill.isValid()) fill = QColor("#888888");

        painter->setBrush(fill);
        painter->setPen(Qt::NoPen);
        painter->drawEllipse(dotX, dotY, kDotDiameter, kDotDiameter);

        const QString intent = i < intents.size() ? intents[i] : QString();
        if (intent.compare("editing", Qt::CaseInsensitive) == 0) {
            painter->setBrush(Qt::NoBrush);
            QPen ringPen(Qt::white);
            ringPen.setWidth(kEditingRingPx);
            painter->setPen(ringPen);
            const int inset = kEditingRingPx + 1;
            painter->drawEllipse(dotX + inset, dotY + inset,
                                 kDotDiameter - 2 * inset,
                                 kDotDiameter - 2 * inset);
        }

        dotX += kDotDiameter + kDotSpacing;
    }

    painter->restore();
}

QSize PresenceDotsDelegate::sizeHint(const QStyleOptionViewItem& option,
                                     const QModelIndex& index) const
{
    QSize base = QStyledItemDelegate::sizeHint(option, index);
    const QStringList colors = index.data(kColorsRole).toStringList();
    if (colors.isEmpty()) return base;

    const int dotsTotal = colors.size() * kDotDiameter
                        + (colors.size() - 1) * kDotSpacing;
    base.rwidth() += kTextPadding + dotsTotal;
    return base;
}
```

### Step 5: Run the test — verify it passes

- [ ] **Re-run `tst_presencedotsdelegate`**

```powershell
.\tests\run-tests.ps1 -Filter tst_presencedotsdelegate
```

Expected: PASS on both tests.

### Step 6: Commit

- [ ] **Commit with focused message**

```bash
git add src/widgets/PresenceDotsDelegate.cpp tests/widgets/
git commit -m "fix(ui): right-anchor presence dots so they survive text elision

Spec item #5. Reserve a right-edge region for dots and let the text
renderer elide into the remaining width, instead of painting dots after
the text where they get cropped on narrow columns.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: #6 — Draw pressure y-axis floor at 2, ceil up above

**Files:**
- Modify: `src/plotting/PlotWidget.cpp:556-602` (Draw Pressure branch)
- Modify: `src/reporting/ReportGenerator.cpp:152-193` (Draw Pressure plot block)
- Modify: `tests/plotting/tst_plotengine.cpp` (or equivalent)

### Step 1: Write the failing test for the scaling rule

- [ ] **Add a unit test for the helper function we'll extract**

Add a free helper function in a shared location both callers use. The cleanest spot is the bottom of `src/plotting/PlotEngine.h` since `PlotConfig` lives there.

Edit `src/plotting/PlotEngine.h` — add this declaration just before the closing `} // namespace DVE`:

```cpp
// Computes the y-axis upper bound for the draw-pressure chart.
// Returns max(2.0, ceil(seriesMax)) so the axis always shows at least
// 0–2 Pa range, expanding to the next integer when data exceeds 2 Pa.
double drawPressureYMax(double seriesMax);
```

In `src/plotting/PlotEngine.cpp`, append the definition at file end:

```cpp
double drawPressureYMax(double seriesMax)
{
    if (seriesMax <= 2.0) return 2.0;
    return std::ceil(seriesMax);
}
```

(`<cmath>` is already included via `PlotEngine.cpp`; verify with `grep -n "#include" src/plotting/PlotEngine.cpp` and add `#include <cmath>` if absent.)

Now add the test. Locate `tests/plotting/tst_plotengine.cpp` (verify path with `ls tests/plotting/`). Add inside the test class:

```cpp
private slots:
    void drawPressureYMax_appliesFloorAndCeil();
```

Implementation:

```cpp
void TstPlotEngine::drawPressureYMax_appliesFloorAndCeil()
{
    // Data well under floor → axis sits at floor.
    QCOMPARE(DVE::drawPressureYMax(0.5), 2.0);
    QCOMPARE(DVE::drawPressureYMax(1.8), 2.0);
    QCOMPARE(DVE::drawPressureYMax(2.0), 2.0);
    // Data above floor → ceil to next integer.
    QCOMPARE(DVE::drawPressureYMax(2.1), 3.0);
    QCOMPARE(DVE::drawPressureYMax(2.7), 3.0);
    QCOMPARE(DVE::drawPressureYMax(5.1), 6.0);
}
```

### Step 2: Run the test — verify it fails to compile

- [ ] **Build the test — expect link error**

```powershell
.\tests\run-tests.ps1 -Filter tst_plotengine
```

Expected: FAIL at link with `undefined reference to DVE::drawPressureYMax(double)`.

### Step 3: Implement `drawPressureYMax` per Step 1

The declaration + definition above already cover this. Build again:

- [ ] **Verify the helper compiles and the test passes**

```powershell
.\tests\run-tests.ps1 -Filter tst_plotengine
```

Expected: PASS.

### Step 4: Apply the helper to the UI plot

- [ ] **Replace `cfg.autoScale = true` with explicit yMin/yMax in `PlotWidget.cpp`**

In `src/plotting/PlotWidget.cpp` around line 597 (the Draw Pressure branch), replace:

```cpp
        cfg.autoScale  = true;
```

with:

```cpp
        // #6: floor draw pressure y-axis at 2, expand to ceil(max) above.
        double seriesMax = 0.0;
        for (const PlotSeries& ps : series)
            for (double v : ps.y) seriesMax = std::max(seriesMax, v);
        cfg.autoScale = false;
        cfg.yMin      = 0.0;
        cfg.yMax      = drawPressureYMax(seriesMax);
```

Confirm `#include <algorithm>` is in the .cpp (for `std::max`) — if not, add it near the top.

### Step 5: Apply the helper to the report plot

- [ ] **Same change in `ReportGenerator.cpp` around line 186**

Replace:

```cpp
            cfg.autoScale  = true;        // draw pressure stays auto-scaled
```

with:

```cpp
            // #6: floor draw pressure y-axis at 2, expand to ceil(max) above.
            double seriesMax = 0.0;
            for (const PlotSeries& ps : series)
                for (double v : ps.y) seriesMax = std::max(seriesMax, v);
            cfg.autoScale = false;
            cfg.yMin      = 0.0;
            cfg.yMax      = drawPressureYMax(seriesMax);
```

### Step 6: Build full app, verify it links

- [ ] **Full build**

```bash
cd build && mingw32-make -j8
```

Expected: clean build, no warnings (`-Werror` is on).

### Step 7: Manual visual check (optional but recommended)

- [ ] **Launch the app, open any file with TPM data, switch plot type to "Draw Pressure", verify y-axis floor is 2**

Recorded as a note in the commit — no automated UI verification on this machine.

### Step 8: Commit

```bash
git add src/plotting/PlotEngine.h src/plotting/PlotEngine.cpp \
        src/plotting/PlotWidget.cpp src/reporting/ReportGenerator.cpp \
        tests/plotting/tst_plotengine.cpp
git commit -m "fix(plot): draw pressure y-axis floors at 2, ceil(max) above

Spec item #6. Both the on-screen plot and the report image now use
drawPressureYMax() so the axis always shows at least 0-2 Pa and expands
to the next integer only when data exceeds 2 Pa. Eliminates the
'small variation looks dramatic' effect on low-pressure draws.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: #3 — Color picker shows already-taken colors

**Files:**
- Modify: `src/database/IdentityPromptDialog.h` (constructor signature + helper)
- Modify: `src/database/IdentityPromptDialog.cpp` (query taken colors, decorate swatches)
- Modify: `src/MainWindow.cpp` (pass PostgresConnection to dialog)

### Step 1: Extend the dialog to accept a connection for the lookup

- [ ] **Update `IdentityPromptDialog.h`**

Replace the existing constructor declaration:

```cpp
    explicit IdentityPromptDialog(IdentityManager* mgr, QWidget* parent = nullptr);
```

with:

```cpp
    // `conn` is optional; when provided the dialog queries the presence
    // table for currently-active user_color values and visually marks
    // those swatches as "taken" (the user can still pick them).
    explicit IdentityPromptDialog(IdentityManager* mgr,
                                  PostgresConnection* conn = nullptr,
                                  QWidget* parent = nullptr);
```

Add the forward declaration near the top:

```cpp
class PostgresConnection;
```

(Inside `namespace DVE { ... }`.) And add a member:

```cpp
private:
    PostgresConnection* m_conn = nullptr;
    QSet<QString> queryTakenColors() const;
```

Add `#include <QSet>` to the header.

### Step 2: Implement the query + visual decoration

- [ ] **Update `IdentityPromptDialog.cpp`**

Add includes near the top:

```cpp
#include "PostgresConnection.h"
#include <QSqlQuery>
#include <QSqlError>
#include <QVariant>
#include <QSet>
```

Replace the constructor signature:

```cpp
IdentityPromptDialog::IdentityPromptDialog(IdentityManager* mgr,
                                           PostgresConnection* conn,
                                           QWidget* parent)
    : QDialog(parent), m_mgr(mgr), m_conn(conn) {
```

Add this method:

```cpp
QSet<QString> IdentityPromptDialog::queryTakenColors() const
{
    QSet<QString> taken;
    if (!m_conn || !m_conn->isOpen()) return taken;

    QSqlQuery q(m_conn->database());
    if (!q.exec("SELECT DISTINCT user_color FROM presence "
                "WHERE last_heartbeat > now() - interval '30 seconds'")) {
        qWarning() << "IdentityPromptDialog: queryTakenColors failed:"
                   << q.lastError().text();
        return taken;
    }
    while (q.next()) {
        const QString c = q.value(0).toString().trimmed();
        if (!c.isEmpty()) taken.insert(c.toLower());
    }
    return taken;
}
```

In the constructor, just before the palette loop, add:

```cpp
    const QSet<QString> taken = queryTakenColors();
```

Change the swatch button creation block so the stylesheet branches on taken:

```cpp
    for (int i = 0; i < palette.size(); ++i) {
        auto* btn = new QToolButton;
        btn->setCheckable(true);
        btn->setMinimumSize(36, 36);
        const bool isTaken = taken.contains(palette[i].toLower());
        if (isTaken) {
            btn->setStyleSheet(QString(
                "QToolButton { background:%1; border:3px solid #888888; "
                "border-radius:18px; } "
                "QToolButton:checked { border:3px solid #000; }"
            ).arg(palette[i]));
            btn->setToolTip(tr("Taken — another active user has this color"));
        } else {
            btn->setStyleSheet(QString(
                "QToolButton { background:%1; border:2px solid #00000022; "
                "border-radius:18px; } "
                "QToolButton:checked { border:3px solid #000; }"
            ).arg(palette[i]));
        }
        btn->setProperty("dve_color_hex", palette[i]);
        m_colorGroup->addButton(btn, i);
        grid->addWidget(btn, i / 6, i % 6);
    }
```

### Step 3: Pass the connection from MainWindow

- [ ] **Update the dialog construction in `src/MainWindow.cpp`**

Around line 264 (the `IdentityPromptDialog dlg(m_identity, this);` call):

```cpp
            DVE::IdentityPromptDialog dlg(m_identity, m_pgConn, this);
```

Verify `m_pgConn` is accessible at that point. If `m_pgConn` is null on first launch (offline-boot), the dialog falls back to the no-decoration path — that's the expected behavior.

Search for any other instantiation of `IdentityPromptDialog`:

```bash
grep -rn "IdentityPromptDialog" src/
```

For every instantiation, supply the connection if available, else `nullptr`.

### Step 4: Build, run, and manually verify

- [ ] **Build full app**

```bash
cd build && mingw32-make -j8
```

Expected: clean build.

- [ ] **Run the existing identity dialog tests if any exist**

```bash
grep -rl "IdentityPromptDialog" tests/
```

If a test exists, run it; if it constructs the dialog with the old signature, update the call site to pass `nullptr` for the connection arg.

### Step 5: Commit

```bash
git add src/database/IdentityPromptDialog.h \
        src/database/IdentityPromptDialog.cpp \
        src/MainWindow.cpp
git commit -m "feat(identity): show already-taken colors in the picker dialog

Spec item #3. Query active presence rows for distinct user_color
values and decorate the matching palette swatches with a gray ring
and 'Taken' tooltip. User can still pick a taken color — this is
visual nudging only, not blocking.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: #1 — New-session collision investigation + fix

**Background:** Initial hypothesis (presence activation on unsaved sessions) was wrong — `MainWindow.cpp:1075` already gates `m_presence->activate(...)` on `sessId > 0`. Actual root cause is the auto-save timer (`MainWindow.cpp:251-254`) committing default-named "New Session" rows to the DB before the user has typed anything meaningful. Other clients then pick those up via NOTIFY.

**Files:**
- Modify: `src/MainWindow.cpp` (gate `onUpdateDatabase` for unnamed sensory sessions)
- Modify: `src/database/DatabaseManager.cpp` (or wherever the sensory upsert lives — verify before touching)
- Create or extend a test for the gate

### Step 0: Verify root cause by reading the auto-save path

- [ ] **Trace the auto-save flow**

```bash
grep -n "onUpdateDatabase\|m_dbSaveTimer" src/MainWindow.cpp
grep -n "saveSensorySession\|trySaveSensorySession" src/database/DatabaseManager.cpp
```

Read the body of `MainWindow::onUpdateDatabase()` and `DatabaseManager::trySaveSensorySession()` (or equivalent). Confirm the hypothesis: the auto-save writes new sessions with `sessionName = "New Session"` and no real content.

If the trace reveals a different root cause (e.g., the path is something else, like a direct INSERT happening in `newSession()` after all), STOP and ask the user before continuing. Document findings in a comment on this task.

### Step 1: Write the failing test

- [ ] **Add a test to the sensory save path test suite**

Locate the sensory save test (likely `tests/database/tst_sensorysave.cpp` or under `tst_databasemanager_*`). Add:

```cpp
void TstSensorySave::defaultNamedEmptySessionDoesNotPersist()
{
    // A session with no user-entered content (name = "New Session",
    // no samples, no tester/assessor) should NOT be persisted by an
    // auto-save trigger — it's a placeholder, not data.
    SensorySession s;
    s.sessionName = "New Session";
    s.date        = "2026-05-15";
    s.timestamp   = "2026-05-15T00:00:00Z";
    // intentionally empty otherwise

    SaveCoordinator::Outcome outcome =
        m_saveCoord->saveSensorySession(s, /*parent=*/nullptr);

    // We expect saves of pure placeholders to short-circuit cleanly
    // without hitting the DB.
    QCOMPARE(outcome, SaveCoordinator::Saved);  // or UserCancelled — pick one
    QCOMPARE(s.id, -1);  // never assigned an id

    // Double-check: count of sensory_sessions in DB should be 0.
    QSqlQuery q(m_conn->database());
    QVERIFY(q.exec("SELECT COUNT(*) FROM sensory_sessions"));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toInt(), 0);
}
```

Verify the exact class name / test fixture in the actual file — the snippet above assumes `TstSensorySave` and `m_saveCoord`/`m_conn` members. Mirror what's already there.

### Step 2: Run the test — verify it fails

- [ ] **Build the test**

```powershell
.\tests\run-tests.ps1 -Filter tst_sensorysave
```

Expected: FAIL — current code persists the empty session.

### Step 3: Add the placeholder predicate

- [ ] **In `src/pipeline/SensoryData.h`, add a free helper**

Just before the closing `} // namespace DVE`:

```cpp
inline bool isPlaceholderSession(const SensorySession& s)
{
    // A session is a "placeholder" — i.e., not yet meaningfully filled
    // in — when it still has the default new-session name AND no other
    // user content. Such sessions must not be persisted to the shared DB.
    const bool defaultName = (s.sessionName == QStringLiteral("New Session"));
    const bool noContent = s.samples.isEmpty()
                        && s.testerName.isEmpty()
                        && s.assessorName.isEmpty()
                        && s.media.isEmpty()
                        && s.testTitle.isEmpty();
    return defaultName && noContent;
}
```

### Step 4: Gate the save coordinator

- [ ] **In `src/database/SaveCoordinator.cpp`, short-circuit on placeholders**

At the top of `SaveCoordinator::saveSensorySession`, add:

```cpp
SaveCoordinator::Outcome
SaveCoordinator::saveSensorySession(SensorySession& s, QWidget* parent)
{
    if (isPlaceholderSession(s)) {
        // #1: do not push empty placeholder sessions to the shared DB —
        // they would otherwise broadcast via NOTIFY and pollute every
        // other client's navigator.
        return Saved;
    }
    // ...existing body...
```

Add `#include "../pipeline/SensoryData.h"` if not already present.

### Step 5: Also short-circuit auto-save batch in MainWindow

- [ ] **In `src/MainWindow.cpp::onUpdateDatabase()`, skip placeholder sensory sessions**

Find the loop that iterates sensory sessions and saves each. Wrap the save call:

```cpp
        for (SensorySession& s : sessions) {
            if (isPlaceholderSession(s)) continue;
            m_saveCoordinator->saveSensorySession(s);
        }
```

(Adapt to the actual loop variable / control flow you find.)

### Step 6: Run the test — verify it passes

- [ ] **Re-run**

```powershell
.\tests\run-tests.ps1 -Filter tst_sensorysave
```

Expected: PASS.

### Step 7: Run the full sensory test suite to catch regressions

- [ ] **Run all sensory tests**

```powershell
.\tests\run-tests.ps1 -Filter tst_sensory
```

Expected: PASS on all.

### Step 8: Commit

```bash
git add src/pipeline/SensoryData.h src/database/SaveCoordinator.cpp \
        src/MainWindow.cpp tests/
git commit -m "fix(db): don't persist empty 'New Session' placeholders

Spec item #1. Default-named, empty SensorySession objects are
placeholders — the result of clicking 'New Sensory Session' before
typing anything. Persisting them via the auto-save timer broadcasts
them through NOTIFY and they appear in every other user's navigator
as a shared 'New Session'. Short-circuit the save when
isPlaceholderSession() returns true.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: #4 — Smooth presence updates investigation + fix

**Background:** The NOTIFY → `presenceChanged` signal → `refreshPresenceFor()` path already exists (`MainWindow.cpp:208-215, 2139-2215`). The slot calls `setData()` on the matching nav item but does NOT call `viewport()->update()` or `dataChanged(...)` on the model — so Qt has no reason to repaint until the next layout invalidation. That's why dots only refresh on the next user action.

**Files:**
- Modify: `src/MainWindow.cpp::refreshPresenceFor` (force a viewport repaint after data role changes)

### Step 1: Write a smoke test for the connection

- [ ] **Verify the NOTIFY → refresh path fires within a tight time budget**

In `tests/database/tst_notificationlistener.cpp` (verify path), add:

```cpp
void TstNotificationListener::presenceNotifyArrivesQuickly()
{
    if (!m_listenerSubscribed) QSKIP("listener not subscribed");

    QSignalSpy spy(m_listener, &NotificationListener::presenceChanged);

    // Insert a presence row from a separate connection so the LISTEN
    // socket sees the trigger.
    insertForeignPresenceRow();

    QVERIFY(spy.wait(/*ms=*/1500));
    QVERIFY(spy.count() >= 1);
}
```

`insertForeignPresenceRow()` is a helper to be added — look at sibling tests like `tst_presencemanager` for how a second connection is set up.

This test verifies the signal arrives; the *UI repaint* part is harder to test in isolation and is verified manually.

### Step 2: Run the test — verify it passes (or skips cleanly when no DB)

- [ ] **Run**

```powershell
.\tests\run-tests.ps1 -Filter tst_notificationlistener
```

Expected: PASS (or SKIP if `DVE_TEST_PG_CONN` unset). If it FAILS, the signal pipeline itself is broken and the rest of this task changes — STOP and report.

### Step 3: Force a viewport repaint at the end of `refreshPresenceFor`

- [ ] **Edit `src/MainWindow.cpp::refreshPresenceFor`**

After each `setData()` block that matched a row (i.e., inside each `if (resourceType == ...)` branch, just after the `break;`), add a viewport update on the affected widget. Replace the three branches like so (showing the sensory branch as a template — apply the same pattern to file + detailed_sensory):

```cpp
    } else if (resourceType == QLatin1String("sensory_session") && m_sensoryNav) {
        QSignalBlocker blocker(m_sensoryNav);
        for (int i = 0; i < m_sensoryNav->count(); ++i) {
            QListWidgetItem* it = m_sensoryNav->item(i);
            if (!it) continue;
            if (it->data(Qt::UserRole).toLongLong() != resourceId) continue;
            it->setData(DVE::PresenceDotsDelegate::kColorsRole,  colors);
            it->setData(DVE::PresenceDotsDelegate::kIntentsRole, intents);
            it->setToolTip(tooltip);
            break;
        }
        m_sensoryNav->viewport()->update();   // force immediate repaint
    }
```

File tree branch: append `m_fileTree->viewport()->update();`.
Detailed sensory branch: append `m_detailedSensoryNav->viewport()->update();`.

### Step 4: Build full app

- [ ] **Verify clean build**

```bash
cd build && mingw32-make -j8
```

### Step 5: Manual two-client check (recorded in commit)

- [ ] **If running on the work machine, launch two installs, both connect to the shared DB, observe presence dots refreshing live without local-side clicks**

If on the home machine, skip this step and document in the commit that manual verification is deferred to the work machine.

### Step 6: Commit

```bash
git add src/MainWindow.cpp tests/database/tst_notificationlistener.cpp
git commit -m "fix(ui): force navigator repaint on remote presence updates

Spec item #4. refreshPresenceFor() was setting item data roles but
the view had no signal to repaint, so dots only updated on the next
user-driven layout pass. Call viewport()->update() on the affected
nav widget so NOTIFY-driven updates render immediately.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: #7 — Sensory power_type + puff_length

**Note on schema:** `sensory_sessions.json_data` is JSONB — sample-level fields serialize through whatever JSON converter the panel uses (`saveToJson` / `applySession`). No `ALTER TABLE` needed. Defaults are applied client-side via the struct initializer.

**Files:**
- Modify: `src/pipeline/SensoryData.h` (add 2 fields)
- Modify: `src/ui/SensoryPanel.cpp` (SampleCard widget — add combo + spin)
- Modify: `src/ui/SensoryPanel.h` (SampleCard member additions if any)
- Modify: `src/reporting/SensoryReportSource.cpp` (or wherever sensory report columns are built)
- Modify: `src/reporting/PptxWriter.cpp` / `ReportGenerator.cpp` (column layout)
- Modify: `tests/reporting/tst_sensoryreportsource.cpp`

### Step 1: Add the struct fields with defaults

- [ ] **Edit `src/pipeline/SensoryData.h`**

In `struct SensorySample`, add after `heatingTechnology`:

```cpp
    // #7: per-sample test conditions added 2026-05-15.
    QString powerType    = QStringLiteral("Constant Voltage");
    double  puffLengthSec = 3.0;
```

### Step 2: Write the failing test for round-trip

- [ ] **Extend `tests/reporting/tst_sensoryreportsource.cpp`**

Add a slot:

```cpp
void TstSensoryReportSource::powerTypeAndPuffLengthRoundTripIntoReport()
{
    SensorySession s;
    s.sessionName = "RoundTripTest";
    s.testerName  = "Alice";
    s.date        = "2026-05-15";

    SensorySample a;
    a.name          = "DeviceA";
    a.voltage       = 3.7;
    a.resistance    = 1.2;
    a.power         = 11.4;
    a.powerType     = "Variable Voltage";
    a.puffLengthSec = 2.5;
    s.samples.append(a);

    SensoryReportSource src;
    src.setSessions({s});
    // Walk the table data for slide 0 and locate the puff_length and the
    // voltage cell.
    const QVector<QStringList> table = src.tableRowsForSlide(0);
    bool foundPuff = false, foundPowerType = false;
    for (const QStringList& row : table) {
        for (const QString& cell : row) {
            if (cell.contains("2.5"))            foundPuff = true;
            if (cell.contains("Variable Voltage")) foundPowerType = true;
        }
    }
    QVERIFY(foundPuff);
    QVERIFY(foundPowerType);
}
```

Method names like `tableRowsForSlide` are illustrative — match the actual SensoryReportSource API. Use `Read tool` on `src/reporting/SensoryReportSource.h` to find the correct accessor.

### Step 3: Run the test — verify it fails

- [ ] **Run**

```powershell
.\tests\run-tests.ps1 -Filter tst_sensoryreportsource
```

Expected: FAIL (new fields aren't in the table output yet).

### Step 4: Add the report columns

- [ ] **Edit `src/reporting/SensoryReportSource.cpp` and the table builder**

Locate the function that builds table rows / columns for the sensory report. Add the new `Puff Length (s)` column immediately before the `Notes` column (search for the literal `"Notes"` in headers).

For the voltage/resistance/power cell formatter, append `\n(<powerType>)` when `powerType` is non-default OR always — confirm with sample inspection. Pattern:

```cpp
const QString vrp = QStringLiteral("%1V / %2Ω / %3W\n(%4)")
    .arg(sample.voltage, 0, 'f', 2)
    .arg(sample.resistance, 0, 'f', 2)
    .arg(sample.power, 0, 'f', 2)
    .arg(sample.powerType);
```

For the puff-length column:

```cpp
const QString puff = QString::number(sample.puffLengthSec, 'f', 1);
```

Apply identical changes in the legacy direct-to-PPTX builder (`PptxWriter.cpp` / `ReportGenerator.cpp`) by searching for the sensory column list and replicating the additions there.

### Step 5: Run the test — verify it passes

- [ ] **Re-run**

```powershell
.\tests\run-tests.ps1 -Filter tst_sensoryreportsource
```

Expected: PASS.

### Step 6: Add UI widgets in SampleCard

- [ ] **Edit `src/ui/SensoryPanel.h::SampleCard`**

Add private members after the existing widgets:

```cpp
    QComboBox*      m_powerTypeCombo;
    QDoubleSpinBox* m_puffLengthSpin;
```

- [ ] **Edit `src/ui/SensoryPanel.cpp::SampleCard` constructor**

After the voltage/resistance/power row, insert the power type combo:

```cpp
    auto* powerTypeRow = new QHBoxLayout;
    powerTypeRow->addWidget(new QLabel(tr("Power type:")));
    m_powerTypeCombo = new QComboBox;
    m_powerTypeCombo->addItems({
        tr("Constant Voltage"), tr("Constant Power"),
        tr("Variable Voltage"), tr("Variable Power")
    });
    m_powerTypeCombo->setMaximumWidth(220);
    powerTypeRow->addWidget(m_powerTypeCombo);
    powerTypeRow->addStretch();
    cardLayout->addLayout(powerTypeRow);
    connect(m_powerTypeCombo, &QComboBox::currentTextChanged,
            this, [this](const QString&){ emit changed(); });
```

(Adapt `cardLayout` to the actual layout variable name used in the constructor.)

Just before the comments edit row, insert the puff length spin:

```cpp
    auto* puffRow = new QHBoxLayout;
    puffRow->addWidget(new QLabel(tr("Puff length:")));
    m_puffLengthSpin = new QDoubleSpinBox;
    m_puffLengthSpin->setRange(0.1, 60.0);
    m_puffLengthSpin->setSingleStep(0.5);
    m_puffLengthSpin->setDecimals(1);
    m_puffLengthSpin->setSuffix(QStringLiteral(" s"));
    m_puffLengthSpin->setValue(3.0);
    m_puffLengthSpin->setMaximumWidth(120);
    puffRow->addWidget(m_puffLengthSpin);
    puffRow->addStretch();
    cardLayout->addLayout(puffRow);
    connect(m_puffLengthSpin, qOverload<double>(&QDoubleSpinBox::valueChanged),
            this, [this](double){ emit changed(); });
```

Update `toSample()`:

```cpp
    s.powerType     = m_powerTypeCombo->currentText();
    s.puffLengthSec = m_puffLengthSpin->value();
```

Update `fromSample()`:

```cpp
    int idx = m_powerTypeCombo->findText(s.powerType);
    if (idx >= 0) m_powerTypeCombo->setCurrentIndex(idx);
    else          m_powerTypeCombo->setCurrentIndex(0);
    m_puffLengthSpin->setValue(s.puffLengthSec > 0 ? s.puffLengthSec : 3.0);
```

### Step 7: Build full app

- [ ] **Clean rebuild — VERSION wasn't bumped but layout changes are wide**

```bash
cd build && mingw32-make -j8
```

### Step 8: Manual UI smoke test (recorded in commit)

- [ ] **Launch app, open Sensory mode, create a new sample, change power type and puff length, save, reopen, verify fields persist**

### Step 9: Commit

```bash
git add src/pipeline/SensoryData.h src/ui/SensoryPanel.h src/ui/SensoryPanel.cpp \
        src/reporting/SensoryReportSource.cpp \
        src/reporting/PptxWriter.cpp src/reporting/ReportGenerator.cpp \
        tests/reporting/tst_sensoryreportsource.cpp
git commit -m "feat(sensory): per-sample power type + puff length fields

Spec item #7. Adds powerType (4-way enum) and puffLengthSec to
SensorySample. Surfaces as a QComboBox below the V/R/P row and a
QDoubleSpinBox above comments in each sample card. Reports gain a
'Puff Length (s)' column before Notes; the V/R/P cell now appends
the power type on a second line. Schema is unchanged — samples
serialize through sensory_sessions.json_data.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 7: Final test sweep + merge to main

### Step 1: Full test suite

- [ ] **Run everything**

```powershell
.\tests\run-tests.ps1
```

Expected: all classes PASS or SKIP cleanly. Investigate any FAIL before proceeding.

### Step 2: Confirm clean tree

```bash
git status                                  # working tree clean
git log --oneline main..HEAD                # six fix commits, no garbage
```

### Step 3: Merge to main and push (per user's branch-to-main workflow)

```bash
git push origin HEAD:main
git checkout -b temp-staging
git fetch origin main
git branch -f main origin/main
git checkout main
git branch -D feature/postwave-bugfix-batch
git push origin --delete feature/postwave-bugfix-batch
git branch -D temp-staging
```

(Or, if `main` can be checked out directly without worktree conflicts: simpler `git checkout main && git pull && git branch -D feature/postwave-bugfix-batch && git push origin --delete feature/postwave-bugfix-batch`.)

### Step 4: Verify final state

```bash
git branch -vv
git log --oneline -10
```

Expected: on `main`, top six commits are the fix commits from this plan, `origin/main` matches.

---

## Self-review notes (already applied to this plan)

- **Spec coverage:** all six items in `2026-05-15-postwave-bugfix-batch-design.md` map to a task here (#5 → Task 1, #6 → Task 2, #3 → Task 3, #1 → Task 4, #4 → Task 5, #7 → Task 6).
- **Placeholders:** "verify before touching" calls in Task 4 Step 0 and Task 6 Step 4 are explicit investigation steps, not TBDs.
- **#1 root-cause correction:** the spec described "defer presence activation" but the code already does that. Task 4 redirects to the actual root cause (auto-save persisting placeholder sessions). If implementation reveals yet a different root cause, the plan tells the implementer to stop and report rather than proceeding with the wrong fix.
- **Type consistency:** `drawPressureYMax`, `isPlaceholderSession`, `kColorsRole`, `kIntentsRole`, `m_powerTypeCombo`, `m_puffLengthSpin` — all names used consistently across tasks.
- **#7 schema:** noted that no `ALTER TABLE` is required because `json_data` is JSONB. This changes the spec's intent without changing the user-visible outcome.

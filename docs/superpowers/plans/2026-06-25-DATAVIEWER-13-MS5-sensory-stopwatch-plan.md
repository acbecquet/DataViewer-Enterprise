# MS-5 — Per-Sample Sensory Stopwatch + Assignable Hotkey — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Start/Stop stopwatch to each Sensory `SampleCard` that, on Stop, fills the existing "Puff length" field with the elapsed seconds — plus a user-assignable hotkey (default **Space**) that toggles the **focused** card's stopwatch, suppressed while typing, with a focus outline on the active card.

**Architecture:** The puff-length field (`puff_length_sec`) already persists end-to-end (serializer + LiveSync + report). The stopwatch is a pure UI addition: a `QElapsedTimer` value member per `SampleCard`; on Stop we call `m_puffLengthSpin->setValue(elapsed)`, which fires the existing `valueChanged → cellCommitted("puff_length_sec") → LiveSync::commitCell` chain unchanged. The hotkey is a `qApp` event filter on `SensoryPanel` (gated on visibility + focus type); the focused card is tracked via `QApplication::focusChanged`.

**Tech Stack:** C++17, Qt 6.10 Widgets, qmake + MinGW. `-Werror -Wall -Wextra`. Ships as an internal v2.5.x patch inside the v2.6.0 train.

---

## Design decisions (locked — owner 2026-06-25)

- **Per-card Start/Stop button** AND **assignable hotkey** (both). Default key **Space**, rebindable in Settings, persisted in `QSettings`.
- Hotkey toggles the **focused** sample card's timer; falls back to the first card if none is focused.
- Hotkey is **suppressed when a text-entry widget has focus** (`QLineEdit`/`QPlainTextEdit`/`QTextEdit`/`QAbstractSpinBox`) so Space types normally (e.g. in the comments box). Owner's words: *"if a text entry box is selected then space should not activate it."*
- A **2px focus outline** is drawn around the card that owns keyboard focus. Owner: *"add a little outline around the focused sample border to make it clear."*
- Elapsed is clamped to **[0.1, 60.0] s** (the spin's range) and rounded to 1 decimal by the spin. >60 s is clamped (no cap-raise).
- **Sensory only.** Detailed Sensory has no `puff_length_sec` field — parity is a separate, larger item.

## Preconditions & grounding (verified 2026-06-25 via `git show d7ca362`)

- **`SampleCard`** is a `QGroupBox` subclass declared in `src/ui/SensoryPanel.h:32-58`, built at `src/ui/SensoryPanel.cpp:217-461`, fixed width **245px** (`:222`).
- **`puffRow`** is a `QHBoxLayout` (`SensoryPanel.cpp:377`). `m_puffLengthSpin` (a `NoWheelDoubleSpinBox`, range 0.1–60.0, step 0.5, suffix `" s"`, default 3.0, fixed 72×20, `NoButtons`) is added at **`:392`**, followed by `puffRow->addStretch()` at `:393`. **Insert the Start/Stop button between `:392` and `:393`.**
- `m_puffLengthSpin` is declared `QDoubleSpinBox*` at `SensoryPanel.h:54`; `NoWheelDoubleSpinBox` is defined at `SensoryPanel.cpp:184-194`.
- `m_puffLengthSpin`'s `valueChanged` already emits `cellCommitted("puff_length_sec", v)` (`SensoryPanel.cpp:397-400`); the per-card handler at `:837-857` routes `cellCommitted → dirtyCells + m_liveSync->commitCell("sensory_sessions", sessionId, "json_path:samples[i].puff_length_sec", value)`. **`setValue()` therefore reuses the entire persistence chain — no schema/serializer/LiveSync change.**
- Cards: `QVector<SampleCard*> m_cards` (`SensoryPanel.h:281`); appended in `addSampleCard()` (`:857`), cleared in `clearAllCards()` (`:851`). `applySession()` rebuilds all cards (so a per-card `QTimer` child dies with its card — no dangling timer). **No existing "selected/active card" concept.**
- `SampleCard` sets **no** `objectName`/stylesheet in its ctor (the inner comments frame uses `objectName "sensoryCommentsFrame"`).
- `QSettings("SDR", "DataViewerEnterprise")` (`MainWindow.cpp:6212`). `buildSettingsTab` is at `MainWindow.cpp:850-885` (group "Output Paths"; `grp->addLargeButton(label, icon, tooltip)` returns a `QToolButton*`).
- Window-key precedent: `QAction` + `setShortcut` + `addAction` (`MainWindow.cpp:1757-1776`). **No existing `eventFilter`** — we introduce one (a plain `QAction`/`QShortcut` with Space would steal Space from text boxes; an event filter with a focus-type guard is the correct idiom).
- **MIP:** `python tools/decrypt_via_copy.py --apply` before every build; create any new source file via the Python delete-and-rewrite convention. Build: `qmake CONFIG+=release -spec win32-g++ DataViewerEnterprise.pro && mingw32-make -j8`. Tests: windows-subsystem, run with `-o out.txt,txt`, Qt `bin` on PATH + `QT_QPA_PLATFORM=offscreen`.
- Guardrail: `/ponytail-review` the diff before each commit — no `Stopwatch` class, no new struct/JSON field, no migration, no sub-100ms precision, no pause/lap.

## File structure

- **Modify** `src/pipeline/SensoryData.h` / `.cpp` — add a tiny pure `clampPuffSeconds(double)` helper (the only unit-testable piece).
- **Modify** `src/ui/SensoryPanel.h` / `.cpp` — `SampleCard` stopwatch members + button + `toggleStopwatch`/`resetStopwatch`/`setStopwatchFocus`; `SensoryPanel` focus tracking + `qApp` event filter + hotkey load.
- **Modify** `src/MainWindow.cpp` — a "Stopwatch Hotkey" rebind control in `buildSettingsTab` + a key-capture dialog; notify the live panel on change.
- **Modify** `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp` — `clampPuffSeconds` bounds test.

---

### Task 1: SampleCard Start/Stop button + elapsed timer (reuses the puff persistence chain)

**Files:** `src/pipeline/SensoryData.h`/`.cpp` (clamp helper), `src/ui/SensoryPanel.h`/`.cpp`, `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp`.

- [ ] **Step 1: Write the failing test** for the clamp helper — in `tst_sensorydataplaceholder.cpp` add to the slots and body:

```cpp
void clampPuffSecondsBoundsToSpinRange();
// ...
void TstSensoryDataPlaceholder::clampPuffSecondsBoundsToSpinRange() {
    QCOMPARE(clampPuffSeconds(0.0),  0.1);   // below floor
    QCOMPARE(clampPuffSeconds(3.27), 3.27);  // in range (spin rounds to 1dp on setValue)
    QCOMPARE(clampPuffSeconds(75.0), 60.0);  // above cap → clamped, never out-of-range
}
```

Run: `cd tests/tst_sensorydataplaceholder && qmake && mingw32-make && QT_QPA_PLATFORM=offscreen ./release/tst_sensorydataplaceholder.exe -o out.txt,txt`. Expected: FAIL (`clampPuffSeconds` undefined).

- [ ] **Step 2: Add the helper** — in `src/pipeline/SensoryData.h` (namespace `DVE`, near the other free functions) declare `double clampPuffSeconds(double rawSeconds);` and in `SensoryData.cpp` define:

```cpp
double clampPuffSeconds(double rawSeconds) {
    // Bound to the puff-length spin's range; the spin itself rounds to 1 decimal.
    return qBound(0.1, rawSeconds, 60.0);
}
```

Add `#include <QtGlobal>` if not present. Run the test → PASS.

- [ ] **Step 3: Add stopwatch members to `SampleCard`** — in `src/ui/SensoryPanel.h` inside the `SampleCard` class add includes (`<QElapsedTimer>`, `<QTimer>`, `<QPushButton>` — or forward declarations + includes in the .cpp) and members/methods:

```cpp
public:
    void toggleStopwatch();          // idle→start (relabel "Stop", live readout); running→stop+commit elapsed
    void resetStopwatch();           // stop the tick + relabel "Start"; never commits (rebuild/reuse safety)
    void setStopwatchFocus(bool on); // toggles the focus-outline dynamic property + repolishes
private:
    QPushButton*  m_stopwatchBtn = nullptr;
    QElapsedTimer m_stopwatch;       // value member — Qt monotonic, never hand-rolled with QDateTime
    QTimer*       m_tickTimer = nullptr;
    bool          m_timing = false;
```

- [ ] **Step 4: Build the button in the ctor** — in `SensoryPanel.cpp`, immediately after `puffRow->addWidget(m_puffLengthSpin);` (`:392`) and before `puffRow->addStretch();`:

```cpp
m_stopwatchBtn = new QPushButton(tr("Start"));
m_stopwatchBtn->setFixedHeight(20);
m_stopwatchBtn->setFocusPolicy(Qt::NoFocus);   // never steals the card's focus outline / typing focus
m_stopwatchBtn->setStyleSheet("font-size: 7pt; padding: 0px 6px;");
m_stopwatchBtn->setToolTip(tr("Start/stop timing this puff (hotkey configurable in Settings)"));
connect(m_stopwatchBtn, &QPushButton::clicked, this, [this]() { toggleStopwatch(); });
puffRow->addWidget(m_stopwatchBtn);

m_tickTimer = new QTimer(this);                 // child of the card → destroyed with it
m_tickTimer->setInterval(200);
connect(m_tickTimer, &QTimer::timeout, this, [this]() {
    if (m_timing)
        m_stopwatchBtn->setText(tr("Stop %1s").arg(m_stopwatch.elapsed() / 1000.0, 0, 'f', 1));
});
```

- [ ] **Step 5: Implement the toggle + reset** — in `SensoryPanel.cpp` (free of the ctor):

```cpp
void SampleCard::toggleStopwatch() {
    if (!m_timing) {
        m_stopwatch.start();
        m_timing = true;
        m_stopwatchBtn->setText(tr("Stop"));
        m_tickTimer->start();
    } else {
        m_timing = false;
        m_tickTimer->stop();
        const double secs = clampPuffSeconds(m_stopwatch.elapsed() / 1000.0);
        m_stopwatchBtn->setText(tr("Start"));
        m_puffLengthSpin->setValue(secs);   // fires valueChanged → cellCommitted → LiveSync (one commit)
    }
}

void SampleCard::resetStopwatch() {
    m_timing = false;
    if (m_tickTimer) m_tickTimer->stop();
    if (m_stopwatchBtn) m_stopwatchBtn->setText(tr("Start"));
}
```

(`clampPuffSeconds` is in `SensoryData.h` — ensure it's included.)

- [ ] **Step 6: Build + smoke** — `qmake CONFIG+=release && mingw32-make -j8` clean. Run, enter Sensory mode, click Start then Stop on a card → "Puff length" updates to the elapsed seconds; the value survives Ctrl+U and appears in the sensory report (same path as a manual edit). **Commit:** `feat(sensory): per-sample stopwatch button fills puff length on stop (DATAVIEWER-13)` (+ Co-Authored-By trailer).

---

### Task 2: Focus outline on the active card

**Files:** `src/ui/SensoryPanel.h`/`.cpp`.

- [ ] **Step 1: Style the card + dynamic property** — in the `SampleCard` ctor (`SensoryPanel.cpp:~218`) add:

```cpp
setObjectName(QStringLiteral("sensoryCard"));
setProperty("stopwatchFocus", false);
setStyleSheet(
    "QGroupBox#sensoryCard { border: 2px solid transparent; border-radius: 4px; margin-top: 6px; }"
    "QGroupBox#sensoryCard::title { subcontrol-origin: margin; left: 8px; padding: 0 3px; }"
    "QGroupBox#sensoryCard[stopwatchFocus=\"true\"] { border: 2px solid #0066CC; }");
```

The transparent default border keeps the layout from jumping when the outline appears. (Visual sign-off: if the groupbox title overlaps the border, adjust `margin-top`/`::title left`.)

- [ ] **Step 2: Implement `setStopwatchFocus`**:

```cpp
void SampleCard::setStopwatchFocus(bool on) {
    if (property("stopwatchFocus").toBool() == on) return;
    setProperty("stopwatchFocus", on);
    style()->unpolish(this);
    style()->polish(this);
}
```

- [ ] **Step 3: Track focus in `SensoryPanel`** — add `SampleCard* m_focusedCard = nullptr;` to `SensoryPanel.h`, declare `void onAppFocusChanged(QWidget* old, QWidget* now);` and a helper `SampleCard* cardOwning(QWidget* w) const;`. In the `SensoryPanel` ctor: `connect(qApp, &QApplication::focusChanged, this, &SensoryPanel::onAppFocusChanged);`. Implement:

```cpp
SampleCard* SensoryPanel::cardOwning(QWidget* w) const {
    for (; w; w = w->parentWidget())
        for (SampleCard* c : m_cards) if (c == w) return c;
    return nullptr;
}

void SensoryPanel::onAppFocusChanged(QWidget* /*old*/, QWidget* now) {
    m_focusedCard = cardOwning(now);
    for (SampleCard* c : m_cards) c->setStopwatchFocus(c == m_focusedCard);
}
```

(Reset `m_focusedCard = nullptr;` in `clearAllCards()` so it never dangles after a rebuild.)

- [ ] **Step 4: Build + smoke + commit** — clicking into a card's field outlines that card; clicking another moves the outline. `feat(sensory): focus outline marks the active sample card (DATAVIEWER-13)`.

---

### Task 3: Persisted, rebindable stopwatch hotkey + Settings control

**Files:** `src/MainWindow.cpp` (Settings tab + capture dialog), `src/ui/SensoryPanel.h`/`.cpp` (load + reload).

- [ ] **Step 1: Load the key in `SensoryPanel`** — add `int m_stopwatchKey = Qt::Key_Space;` to `SensoryPanel.h` and `void reloadStopwatchHotkey();` (public). Implement:

```cpp
void SensoryPanel::reloadStopwatchHotkey() {
    QSettings s(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"));
    m_stopwatchKey = s.value(QStringLiteral("sensory/stopwatchHotkey"),
                             int(Qt::Key_Space)).toInt();
}
```

Call `reloadStopwatchHotkey();` once in the `SensoryPanel` ctor.

- [ ] **Step 2: Add the rebind control to `buildSettingsTab`** — in `MainWindow.cpp` (`:850-885`) add a group (e.g. "Sensory") with a large button whose caption shows the current key:

```cpp
auto* swGrp = tab->addGroup(tr("Sensory"));
auto* swBtn = swGrp->addLargeButton(stopwatchHotkeyLabel(), QIcon(), tr("Rebind the Sensory stopwatch start/stop key"));
connect(swBtn, &QToolButton::clicked, this, [this, swBtn]() {
    int key = captureStopwatchKey();          // modal capture dialog (Step 3); 0 = cancelled
    if (key == 0) return;
    QSettings s(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"));
    s.setValue(QStringLiteral("sensory/stopwatchHotkey"), key);
    swBtn->setText(stopwatchHotkeyLabel());
    if (m_sensoryPanel) m_sensoryPanel->reloadStopwatchHotkey();
});
```

Add a helper `QString MainWindow::stopwatchHotkeyLabel() const` returning `"Stopwatch Key: " + QKeySequence(QSettings(...).value("sensory/stopwatchHotkey", int(Qt::Key_Space)).toInt()).toString()`.

- [ ] **Step 3: Implement `captureStopwatchKey()`** — a minimal modal that records one key press:

```cpp
int MainWindow::captureStopwatchKey() {
    QDialog dlg(this);
    dlg.setWindowTitle(tr("Press a key for the stopwatch"));
    auto* lay = new QVBoxLayout(&dlg);
    lay->addWidget(new QLabel(tr("Press the key to use for the Sensory stopwatch.\n"
                                 "Esc cancels."), &dlg));
    int captured = 0;
    dlg.installEventFilter(new KeyGrab(&dlg, &captured));  // tiny QObject that sets *captured and accepts the dialog
    dlg.exec();
    return captured;   // Esc leaves it 0
}
```

Implement `KeyGrab` as a small local `QObject` whose `eventFilter` on `QEvent::KeyPress` ignores `Qt::Key_Escape` (sets 0, rejects) and otherwise stores `ke->key()` and accepts the dialog. (Keep it in an anonymous namespace in `MainWindow.cpp`.) Modifier-only keys (Shift/Ctrl/Alt/Meta) should be ignored so the user picks a real key.

- [ ] **Step 4: Build + smoke + commit** — Settings shows "Stopwatch Key: Space"; rebinding to e.g. F9 persists across restart. `feat(sensory): rebindable stopwatch hotkey setting (default Space) (DATAVIEWER-13)`.

---

### Task 4: Hotkey event filter — toggle the focused card, suppressed while typing

**Files:** `src/ui/SensoryPanel.h`/`.cpp`.

- [ ] **Step 1: Install the filter + override** — in the `SensoryPanel` ctor add `qApp->installEventFilter(this);`. Declare `bool eventFilter(QObject* obj, QEvent* ev) override;` and a static `isTextEntry`. Implement:

```cpp
static bool isTextEntry(QWidget* w) {
    return qobject_cast<QLineEdit*>(w)      || qobject_cast<QPlainTextEdit*>(w)
        || qobject_cast<QTextEdit*>(w)      || qobject_cast<QAbstractSpinBox*>(w) != nullptr;
}

bool SensoryPanel::eventFilter(QObject* obj, QEvent* ev) {
    if (ev->type() == QEvent::KeyPress && isVisible() && m_stopwatchKey != 0) {
        auto* ke = static_cast<QKeyEvent*>(ev);
        if (ke->key() == m_stopwatchKey && ke->modifiers() == Qt::NoModifier) {
            // Suppressed while typing — Space (or whatever key) reaches the editor unchanged.
            if (isTextEntry(QApplication::focusWidget()))
                return QWidget::eventFilter(obj, ev);
            SampleCard* target = m_focusedCard
                ? m_focusedCard
                : (m_cards.isEmpty() ? nullptr : m_cards.first());
            if (target) { target->toggleStopwatch(); return true; }  // consume
        }
    }
    return QWidget::eventFilter(obj, ev);
}
```

Add includes: `<QApplication>`, `<QKeyEvent>`, `<QLineEdit>`, `<QPlainTextEdit>`, `<QTextEdit>`, `<QAbstractSpinBox>`.

- [ ] **Step 2: Build + smoke** — in Sensory mode: focus a card (not a text box), press Space → that card's stopwatch toggles; focus the comments box, press Space → a space is typed (no toggle); switch to TPM mode, press Space → nothing (panel not visible). Rebind to F9 in Settings and re-verify. **Commit:** `feat(sensory): Space hotkey toggles the focused card's stopwatch, suppressed while typing (DATAVIEWER-13)`.

---

### Task 5: Edge cases, verification, lessons

- [ ] **Step 1:** Confirm timer hygiene — `applySession()`/`clearAllCards()` destroy and rebuild cards, so the per-card `QTimer` dies with its card (no dangling). If any reuse path keeps a card alive across `fromSample()`, call `resetStopwatch()` there. Verify removing a card mid-run does not crash and emits no spurious commit.
- [ ] **Step 2:** Confirm a programmatic `setValue` on a persisted (id>0) LiveSync session streams exactly **one** `json_path:samples[i].puff_length_sec` commit (watch the LiveSync log / a test DB), and reload shows a JSON **number** (not reverted to 3.0).
- [ ] **Step 3:** `/ponytail-review` the full diff — assert no `Stopwatch` class, no new struct/JSON field, no migration, no pause/lap/sub-100ms precision crept in.
- [ ] **Step 4:** `python tools/decrypt_via_copy.py --apply`; `qmake CONFIG+=release && mingw32-make clean && mingw32-make -j8` `-Werror` clean; `tests\run-tests.ps1` green (except known-flaky `tst_responsivelayout`); the new `clampPuffSeconds` test passes; the existing puff round-trip still passes.
- [ ] **Step 5:** Append `tasks/lessons.md`: "Programmatic `QDoubleSpinBox::setValue` re-fires `valueChanged`, so the stopwatch reuses the entire puff-length persistence/LiveSync/report chain with zero new plumbing; a context-sensitive global key needs an event filter with a focus-type guard, not a `QShortcut` (which would steal Space from text boxes)." **Commit** the lessons + any fixups.

---

## Self-review notes

- **Spec coverage:** master-spec MS-5 button → Task 1; today's hotkey/focus-outline/typing-guard decisions → Tasks 2–4; range clamp + single-commit + timer hygiene → Tasks 1/5; Sensory-only → whole plan (Detailed deferred).
- **Type consistency:** `toggleStopwatch()`/`resetStopwatch()`/`setStopwatchFocus(bool)` on `SampleCard`; `m_focusedCard`/`m_stopwatchKey`/`reloadStopwatchHotkey()`/`eventFilter` on `SensoryPanel`; `clampPuffSeconds(double)` in `SensoryData`.
- **Ponytail floor:** reuses `m_puffLengthSpin` + `cellCommitted` (no new field), `QElapsedTimer`/`QTimer`/`QSettings` (no hand-rolled time/persistence), one event filter (no new shortcut framework).

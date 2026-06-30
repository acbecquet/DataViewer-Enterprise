# Responsive UI Overhaul (v2.7.0) Implementation Plan
> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (- [ ]) syntax for tracking.

**Goal:** Make DataViewer Enterprise's UI fully responsive — the window shrinks to a ~480×360 floor at any aspect ratio with nothing inaccessible (horizontal *and* vertical scrollbars appear as-needed wherever content exceeds its viewport), split-screen/corner-snap work, OS text-scaling (125/150/200%) and varied DPI never clip text or overlap widgets, the ribbon "View Raw Data" label wraps to ≤2 lines and never spills into the Navigator, and there is **no visual regression** at today's standard sizes (≥1100px wide, 100% scale).

**Architecture:** View/layout-only change (the lowest-risk category in this app — no data-model, pipeline, DB, save-path, plotting-engine, or reporting changes). A new reusable `DVE::ScrollHost` (`QScrollArea` subclass) is the structural guarantee that nothing is ever clipped without a scrollbar; it wraps every major region (the three central-stack pages, both side docks, the ribbon group row, and dialog content). New pure `AppTheme` font-metric helpers (`lineUnit`/`controlHeight`/`em`) make clip-prone control dimensions track the active font. The existing `ResponsiveLayout` singleton is extended with one `VeryNarrow` breakpoint that auto-collapses both side docks. A systematic fixed→min sweep converts the remaining clip-prone `setFixed*`/alignment-only `setMaximum*` sites. Verification is layered: light, focused Qt Tests pin the unit-level invariants (`tst_scrollhost` scroll behaviour, `tst_ribbonlayout` ≤2-line wrap, `tst_responsivelayout` breakpoints, `tst_sizingsweep` dialog floor) — each linking only the few sources it needs, none instantiating the heavy `MainWindow` — while a new `--ui-stress` screenshot CLI flag proves the full-window "nothing clipped without a scrollbar" guarantee in-process inside the real `DataViewer.exe`, writing a per-case `pass`/`fail` verdict to a JSON index that the agent reads as a closed-loop gate.

**Tech Stack:** C++17, Qt 6.10 (Widgets), qmake+MinGW, namespace DVE

---

## Conventions for every task in this plan

- **MIP machine.** Run `python tools/decrypt_via_copy.py --apply` from the repo root before any C++ build. Create *all* new source files via Python's delete-and-rewrite pattern (per project CLAUDE.md), never via the Edit/Write create path — a freshly Edit/Write-created file can pick up a MIP label and become ciphertext at the next build.
- **Build flags** are `-Werror -Wall -Wextra`. Never downgrade the warning level to silence a warning; fix the underlying code.
- **VERSION bumps need a clean rebuild** (`mingw32-make clean && mingw32-make -j8`), but this plan does not bump VERSION until the final wrap — incremental builds are fine throughout.
- Match the surrounding style of the file you edit. New widgets follow `src/widgets/OfflineBanner.{h,cpp}` (`#pragma once`, `namespace DVE`, `QStringLiteral`, forward-declare in the header, full includes in the `.cpp`).

### Unified cross-section contracts (read before starting)

These were reconciled during synthesis so every task agrees on one API:

- **`DVE::ScrollHost` public API** (Task 1 is the single source of truth; later tasks consume it):
  - `explicit ScrollHost(QWidget* parent = nullptr)` — vertical+horizontal as-needed.
  - `static ScrollHost* wrap(QWidget* content, Qt::Orientations scroll = Qt::Horizontal | Qt::Vertical)` — the factory. The optional second arg restricts scrolling to one axis; passing `Qt::Horizontal` (as the ribbon does) forces the vertical policy to `Qt::ScrollBarAlwaysOff`.
  - `bool scrollbarActive(Qt::Orientation o) const` — true when that direction's scrollbar is currently shown (used by the verification sweep).
  - `bool contentOverflows() const` — true when either scrollbar's `maximum() > 0`.
- **`AppTheme` font-metric helpers** (Task 4 is the single source of truth):
  - `static int lineUnit(const QFont& f = fontDefault())`
  - `static int controlHeight(const QFont& f = fontDefault(), int vPad = 6)`
  - `static int em(qreal n, const QFont& f = fontDefault())`
- **Dependency order is enforced by the phase/task numbering below:** `ScrollHost` (Task 1) and the `AppTheme` helpers (Task 4) land before any component that uses them. Within Phase 1, Task 1 (ScrollHost) precedes Tasks 2–3 (MainWindow region wrapping + window floor) and Tasks 6–9 (ribbon, which also consumes the AppTheme helpers from Task 4 — so the ribbon tasks are sequenced *after* Task 4 even though they are conceptually "Phase 1 headline bug"; see the note on Task 6).
- **Window-floor ownership:** the `setMinimumSize(480, 360)` change lives in exactly one place — Task 3 Step 2. The HiDPI `PassThrough` line lives in exactly one place — Task 5 (`main.cpp`).
- **`m_centralScrollHost` / `m_navScrollHost` / `m_notesScrollHost` member names:** the test-support accessors in Task 16 expect MainWindow to expose the per-region ScrollHosts (consumed by the `--ui-stress` harness, Tasks 18–19). Tasks 2–3 wrap regions via `ScrollHost::wrap(...)` added directly into layouts/docks; **store each returned `ScrollHost*` in a member** (`m_centralScrollHost` is N/A because the central pages live in the stack — see Task 16's accessor, which resolves the central host via the stack's current page). Concretely: Task 3 stores `m_navScrollHost` (Navigator) and `m_notesScrollHost` (Notes) as new private members; the central host is reached via `m_centralStack->currentWidget()` cast to `DVE::ScrollHost*`. Task 16 wires the accessors to these.

---

## Phase 1 — Structural guarantee + the headline bug

Delivers "nothing inaccessible" (ScrollHost + region wrapping + window floor + HiDPI) and the View Raw Data fix (ribbon ≤2-line/grow). Per-region scroll wrapping is the correct Qt pattern for a `QMainWindow` with docks (a single whole-window scroll is not feasible because the dock manager owns the dock areas).

### Task 1: Create the `ScrollHost` widget (TDD)

**Files:**
- Create: `src/widgets/ScrollHost.h`, `src/widgets/ScrollHost.cpp`
- Create: `tests/tst_scrollhost/tst_scrollhost.pro`, `tests/tst_scrollhost/tst_scrollhost.cpp`
- Modify: `tests/tests.pro` (SUBDIRS list, ~line 51), `DataViewerEnterprise.pro` (SOURCES ~line 63, HEADERS ~line 143)

New public symbols (keep consistent across the whole plan): `DVE::ScrollHost` (class), `DVE::ScrollHost::ScrollHost(QWidget* parent = nullptr)`, `static DVE::ScrollHost* DVE::ScrollHost::wrap(QWidget* content, Qt::Orientations scroll = Qt::Horizontal | Qt::Vertical)`, `bool DVE::ScrollHost::scrollbarActive(Qt::Orientation) const`, `bool DVE::ScrollHost::contentOverflows() const`.

- [ ] **Step 1: Write the failing test.** Create `tests/tst_scrollhost/tst_scrollhost.cpp` via Python. It builds a `ScrollHost`, puts an oversized content widget inside, shows it small, and asserts the scrollbars activate in each direction:

```python
import os
path = r"tests/tst_scrollhost/tst_scrollhost.cpp"
content = r'''#include <QApplication>
#include <QScrollBar>
#include <QWidget>
#include <QtTest>

#include "ScrollHost.h"

using DVE::ScrollHost;

namespace {

// A widget that reports a fixed, oversized sizeHint in the requested
// direction(s) so the host's ScrollBarAsNeeded policy must engage when the
// viewport is smaller. minimumSizeHint mirrors sizeHint so widgetResizable
// can never shrink it below the overflow point.
class FixedHintWidget : public QWidget {
public:
    explicit FixedHintWidget(QSize hint, QWidget* parent = nullptr)
        : QWidget(parent), m_hint(hint) {}
    QSize sizeHint() const override { return m_hint; }
    QSize minimumSizeHint() const override { return m_hint; }
private:
    QSize m_hint;
};

// Lay the host out at a known small viewport and let the as-needed policy
// evaluate. A top-level show + event flush is required for the scrollbars to
// re-range against the viewport.
void settle(ScrollHost& host, QSize viewport) {
    host.resize(viewport);
    host.show();
    QApplication::processEvents();
    QTest::qWait(50);
    QApplication::processEvents();
}

} // namespace

class tst_ScrollHost : public QObject {
    Q_OBJECT
private slots:
    // The factory yields a configured, non-null host that adopts the content.
    void testWrapAdoptsContent() {
        auto* content = new QWidget;
        ScrollHost* host = ScrollHost::wrap(content);
        QVERIFY(host != nullptr);
        QCOMPARE(host->widget(), content);
        QVERIFY(host->widgetResizable());
        QCOMPARE(host->frameShape(), QFrame::NoFrame);
        QCOMPARE(host->horizontalScrollBarPolicy(), Qt::ScrollBarAsNeeded);
        QCOMPARE(host->verticalScrollBarPolicy(), Qt::ScrollBarAsNeeded);
        delete host;
    }

    // Content taller than the viewport -> vertical scrollbar engages.
    void testVerticalScrollWhenContentTaller() {
        ScrollHost host;
        host.setWidget(new FixedHintWidget(QSize(120, 4000)));
        settle(host, QSize(300, 200));
        QVERIFY2(host.verticalScrollBar()->maximum() > 0,
                 "vertical content overflow must produce a scrollable range");
        QVERIFY(host.scrollbarActive(Qt::Vertical));
        QVERIFY(host.contentOverflows());
    }

    // Content wider than the viewport -> horizontal scrollbar engages.
    void testHorizontalScrollWhenContentWider() {
        ScrollHost host;
        host.setWidget(new FixedHintWidget(QSize(4000, 120)));
        settle(host, QSize(300, 200));
        QVERIFY2(host.horizontalScrollBar()->maximum() > 0,
                 "horizontal content overflow must produce a scrollable range");
        QVERIFY(host.scrollbarActive(Qt::Horizontal));
    }

    // Content that fits -> no scroll range in either direction (invisible no-op).
    void testNoScrollWhenContentFits() {
        ScrollHost host;
        host.setWidget(new FixedHintWidget(QSize(100, 100)));
        settle(host, QSize(600, 600));
        QCOMPARE(host.verticalScrollBar()->maximum(), 0);
        QCOMPARE(host.horizontalScrollBar()->maximum(), 0);
        QVERIFY(!host.contentOverflows());
    }

    // Horizontal-only wrap forces the vertical axis off (ribbon group row).
    void testHorizontalOnlyWrapDisablesVerticalAxis() {
        ScrollHost* host = ScrollHost::wrap(new FixedHintWidget(QSize(4000, 4000)),
                                            Qt::Horizontal);
        QCOMPARE(host->verticalScrollBarPolicy(), Qt::ScrollBarAlwaysOff);
        QCOMPARE(host->horizontalScrollBarPolicy(), Qt::ScrollBarAsNeeded);
        delete host;
    }
};

QTEST_MAIN(tst_ScrollHost)
#include "tst_scrollhost.moc"
'''
if os.path.exists(path):
    os.remove(path)
os.makedirs(os.path.dirname(path), exist_ok=True)
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
```

- [ ] **Step 2: Create the test `.pro`** via Python (mirrors `tst_offlinebanner.pro`):

```python
import os
path = r"tests/tst_scrollhost/tst_scrollhost.pro"
content = '''QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_scrollhost

INCLUDEPATH += ../../src ../../src/widgets

HEADERS += ../../src/widgets/ScrollHost.h

SOURCES += tst_scrollhost.cpp            ../../src/widgets/ScrollHost.cpp
'''
if os.path.exists(path):
    os.remove(path)
os.makedirs(os.path.dirname(path), exist_ok=True)
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
```

- [ ] **Step 3: Run the test — expect BUILD FAILURE** (header/source don't exist yet). From the repo root in Git Bash:

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_scrollhost
PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" /c/Qt/6.10.1/mingw_64/bin/qmake.exe -spec win32-g++ tst_scrollhost.pro
PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: `fatal error: ScrollHost.h: No such file or directory` — the test is wired and red.

- [ ] **Step 4: Implement `ScrollHost.h`** via Python:

```python
import os
path = r"src/widgets/ScrollHost.h"
content = r'''#pragma once

#include <QScrollArea>

namespace DVE {

// Universal scroll wrapper used across MainWindow's central pages, side docks,
// the ribbon group row, and dialog content. It is the structural guarantee of
// the v2.7.0 responsive overhaul: any region wrapped in a ScrollHost can never
// be clipped without a scrollbar appearing, in either direction, the instant
// content exceeds the viewport.
//
// Behaviour (set once in the ctor):
//   - setWidgetResizable(true)  -> content expands to fill the viewport when
//     there is room; ScrollHost scrolls (not clips) when there is not.
//   - both scrollbar policies = Qt::ScrollBarAsNeeded -> invisible when content
//     fits, present the moment it overflows.
//   - QFrame::NoFrame + transparent background + zero viewport margins -> a
//     visual no-op until it actually scrolls (no border, no background fill,
//     no inset), so wrapping a region never changes its standard-size look.
//
// Use the static factory at call sites for a clean read:
//     m_centralStack->addWidget(ScrollHost::wrap(m_centralSplitter));
// The optional second argument restricts scrolling to one axis (the ribbon
// group row passes Qt::Horizontal so the single fixed-height row never grows a
// vertical scrollbar).
//
// Caveat -- QSplitter content: a QSplitter inside a widgetResizable scroll area
// reports its sizeHint as the sum of its children's hints, so the splitter is
// driven by the ScrollHost rather than the reverse. This is fine: at small
// sizes the ScrollHost scrolls the whole splitter; at normal sizes the splitter
// fills the viewport and its handles work as usual. Give splitter children a
// sensible minimumHeight/Width (already done for the TPM plot) so the overflow
// point is meaningful.
class ScrollHost : public QScrollArea {
    Q_OBJECT
public:
    explicit ScrollHost(QWidget* parent = nullptr);

    // Convenience factory: construct a ScrollHost, take ownership of `content`
    // via setWidget(), and return the host. `scroll` selects which axes scroll
    // as-needed; an axis not in `scroll` is set to Qt::ScrollBarAlwaysOff. The
    // host's parent is left null so the caller can re-parent it by adding it to
    // a layout / stacked widget / dock, exactly like a plain `new` widget.
    static ScrollHost* wrap(QWidget* content,
                            Qt::Orientations scroll = Qt::Horizontal | Qt::Vertical);

    // True when that direction's scrollbar is currently shown. Used by the
    // --ui-stress harness to assert the fits-or-scrolls guarantee per region.
    bool scrollbarActive(Qt::Orientation o) const;

    // True when either direction has a non-zero scrollable range.
    bool contentOverflows() const;
};

} // namespace DVE
'''
if os.path.exists(path):
    os.remove(path)
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
```

- [ ] **Step 5: Implement `ScrollHost.cpp`** via Python:

```python
import os
path = r"src/widgets/ScrollHost.cpp"
content = r'''#include "ScrollHost.h"

#include <QPalette>
#include <QScrollBar>

namespace DVE {

ScrollHost::ScrollHost(QWidget* parent)
    : QScrollArea(parent)
{
    setObjectName(QStringLiteral("scrollHost"));
    // Content fills the viewport when there is room; scrolls when there is not.
    setWidgetResizable(true);
    // Scrollbars appear only on overflow, independently per direction.
    setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    // Visual no-op until it scrolls: no frame, transparent viewport, no inset.
    setFrameShape(QFrame::NoFrame);
    setLineWidth(0);
    setViewportMargins(0, 0, 0, 0);
    setContentsMargins(0, 0, 0, 0);
    // Let the wrapped region's own background show through; do NOT paint our
    // own surface (the scroll area's viewport autofills by default).
    setBackgroundRole(QPalette::NoRole);
    viewport()->setAutoFillBackground(false);
    setAttribute(Qt::WA_TranslucentBackground, false);
}

ScrollHost* ScrollHost::wrap(QWidget* content, Qt::Orientations scroll)
{
    auto* host = new ScrollHost();
    if (!(scroll & Qt::Horizontal))
        host->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    if (!(scroll & Qt::Vertical))
        host->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    host->setWidget(content);   // ScrollHost takes ownership of content
    return host;
}

bool ScrollHost::scrollbarActive(Qt::Orientation o) const
{
    const QScrollBar* bar =
        (o == Qt::Horizontal) ? horizontalScrollBar() : verticalScrollBar();
    return bar && bar->isVisible() && bar->maximum() > 0;
}

bool ScrollHost::contentOverflows() const
{
    return (horizontalScrollBar() && horizontalScrollBar()->maximum() > 0) ||
           (verticalScrollBar()   && verticalScrollBar()->maximum()   > 0);
}

} // namespace DVE
'''
if os.path.exists(path):
    os.remove(path)
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
```

- [ ] **Step 6: Run the test — expect PASS.** From `tests/tst_scrollhost`:

```bash
PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" /c/Qt/6.10.1/mingw_64/bin/qmake.exe -spec win32-g++ tst_scrollhost.pro
PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
./release/tst_scrollhost.exe
```

Expected: `Totals: 5 passed, 0 failed, 0 skipped`. (The test deliberately uses `maximum()`/`scrollbarActive()`, not bare `isVisible()`, for headless robustness.)

- [ ] **Step 7: Wire the test into `tests/tests.pro`.** Change the final line (line 51) from `    tst_notesstory` to:

```
    tst_notesstory \
    tst_scrollhost
```

- [ ] **Step 8: Wire the widget into the app `.pro`.** In `DataViewerEnterprise.pro`, in SOURCES (after the `src/widgets/FlowLayout.cpp \` line, ~line 63) add `    src/widgets/ScrollHost.cpp \`, and in HEADERS (after `src/widgets/FlowLayout.h \`, ~line 143) add `    src/widgets/ScrollHost.h \`.

- [ ] **Step 9: Commit.**

```bash
git add src/widgets/ScrollHost.h src/widgets/ScrollHost.cpp tests/tst_scrollhost/tst_scrollhost.pro tests/tst_scrollhost/tst_scrollhost.cpp tests/tests.pro DataViewerEnterprise.pro
git commit -m "feat(ui): add ScrollHost universal scroll wrapper (v2.7.0 Phase 1)

QScrollArea subclass with widgetResizable + ScrollBarAsNeeded (both
directions) + NoFrame + zero margins, plus a ScrollHost::wrap(content,
orientations) factory and scrollbarActive()/contentOverflows() queries.
This is the structural guarantee that no region is ever clipped without
a scrollbar. Includes tst_scrollhost and .pro wiring.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 2: Wrap the three `m_centralStack` pages in MainWindow

**Files:**
- Modify: `src/MainWindow.cpp` (add `#include "widgets/ScrollHost.h"`; `setupCentralWidget()` ~line 1046-1047; `initSensoryPanel()` ~line 3959; `initDetailedSensoryPanel()` ~line 3999; the four `setCurrentWidget(...)` sites ~3886/3895/3930/3938)
- Modify: `tests/tst_scrollhost/tst_scrollhost.cpp` (add the parent-chain contract guard)

`QStackedWidget::setCurrentWidget()` requires its argument to be a direct child *page*. After wrapping, the page becomes the `ScrollHost`, not the inner widget. The member pointers (`m_centralSplitter`, `m_sensoryPanel`, `m_detailedSensoryPanel`) stay pointing at the inner widgets; a file-local `stackPageFor()` helper hops from inner widget to its wrapping page. No new members, no header change.

- [ ] **Step 1: Add the parent-chain contract guard to `tst_scrollhost.cpp`.** Append this slot to `tst_ScrollHost` (before the closing `};`):

```cpp
    // Locks the MainWindow wrapping contract: after wrap(), the page that a
    // QStackedWidget must switch to is the ScrollHost, reachable from the inner
    // widget via parentWidget()->parentWidget() (viewport is the intermediate).
    void testInnerWidgetReachesHostViaParent() {
        auto* inner = new QWidget;
        ScrollHost* host = ScrollHost::wrap(inner);
        QCOMPARE(inner->parentWidget(), host->viewport());
        QCOMPARE(inner->parentWidget()->parentWidget(),
                 static_cast<QWidget*>(host));
        delete host;
    }
```

Run it (expect PASS — documents the Qt contract the code relies on; a guard, not red-first):

```bash
cd tests/tst_scrollhost && ./release/tst_scrollhost.exe
```

Expected: `6 passed`. (Qt 6.10's `QScrollArea::setWidget` reparents into `viewport()`, so grandparent is the host.)

- [ ] **Step 2: Add the include.** In `src/MainWindow.cpp`, add directly after `#include "MainWindow.h"`:

```cpp
#include "widgets/ScrollHost.h"
```

- [ ] **Step 3: Wrap the TPM page in `setupCentralWidget()`** (~line 1046-1047). Replace:

```cpp
    // Wrap in a stacked widget (index 0 = TPM, index 1 = sensory, added lazily)
    m_centralStack = new QStackedWidget(this);
    m_centralStack->addWidget(m_centralSplitter);   // index 0
```

with:

```cpp
    // Wrap in a stacked widget (index 0 = TPM, index 1 = sensory, added lazily).
    // Each page is wrapped in a ScrollHost so it scrolls instead of clipping at
    // small window sizes (v2.7.0). m_centralSplitter stays pointing at the inner
    // splitter; the page added to the stack is its ScrollHost. setCurrentWidget
    // call sites use stackPageFor() to resolve the inner widget to its page.
    m_centralStack = new QStackedWidget(this);
    m_centralStack->addWidget(ScrollHost::wrap(m_centralSplitter));   // index 0
```

- [ ] **Step 4: Add a file-local resolver helper** near the top of `MainWindow.cpp` (in the existing anonymous namespace if one exists, else open one after the includes):

```cpp
namespace {
// v2.7.0: each m_centralStack page is a ScrollHost wrapping an inner widget
// (the TPM splitter / SensoryPanel / DetailedSensoryPanel). setCurrentWidget()
// needs the *page* (the ScrollHost), so resolve the inner widget up to whatever
// widget is the stack's direct child. Returns `inner` unchanged if it is not
// wrapped (defensive -- keeps old behaviour if a page is ever added unwrapped).
QWidget* stackPageFor(QStackedWidget* stack, QWidget* inner) {
    QWidget* w = inner;
    while (w && w->parentWidget() && stack->indexOf(w) < 0) {
        w = w->parentWidget();
    }
    return (w && stack->indexOf(w) >= 0) ? w : inner;
}
} // namespace
```

- [ ] **Step 5: Update the two TPM `setCurrentWidget(m_centralSplitter)` sites** (in `toggleSensoryMode` ~line 3895 and `toggleDetailedSensoryMode` ~line 3938). Replace each `m_centralStack->setCurrentWidget(m_centralSplitter);` with:

```cpp
        m_centralStack->setCurrentWidget(stackPageFor(m_centralStack, m_centralSplitter));
```

- [ ] **Step 6: Wrap + switch the sensory page in `initSensoryPanel()`** (~line 3957-3959). Replace:

```cpp
    m_sensoryPanel = new SensoryPanel(m_db, this);
    m_sensoryPanel->setLiveSync(m_liveSync);                // nullptr-safe
    m_centralStack->addWidget(m_sensoryPanel);   // index 1
```

with:

```cpp
    m_sensoryPanel = new SensoryPanel(m_db, this);
    m_sensoryPanel->setLiveSync(m_liveSync);                // nullptr-safe
    // Wrap in a ScrollHost so the horizontal cards/chart split scrolls instead
    // of clipping at small sizes (v2.7.0). m_sensoryPanel stays the inner ptr.
    m_centralStack->addWidget(ScrollHost::wrap(m_sensoryPanel));   // index 1
```

Then update the activation site in `toggleSensoryMode` (~line 3886). Replace `m_centralStack->setCurrentWidget(m_sensoryPanel);` with:

```cpp
        m_centralStack->setCurrentWidget(stackPageFor(m_centralStack, m_sensoryPanel));
```

- [ ] **Step 7: Wrap + switch the detailed-sensory page in `initDetailedSensoryPanel()`** (~line 3997-3999). Replace:

```cpp
    m_detailedSensoryPanel = new DetailedSensoryPanel(m_db, this);
    m_detailedSensoryPanel->setLiveSync(m_liveSync);                // nullptr-safe
    m_centralStack->addWidget(m_detailedSensoryPanel);
```

with:

```cpp
    m_detailedSensoryPanel = new DetailedSensoryPanel(m_db, this);
    m_detailedSensoryPanel->setLiveSync(m_liveSync);                // nullptr-safe
    // Wrap in a ScrollHost (v2.7.0): the 2-column grid + 4-quadrant chart can
    // overflow at small sizes; scroll instead of clip. inner ptr unchanged.
    m_centralStack->addWidget(ScrollHost::wrap(m_detailedSensoryPanel));
```

Then update the activation site in `toggleDetailedSensoryMode` (~line 3930). Replace `m_centralStack->setCurrentWidget(m_detailedSensoryPanel);` with:

```cpp
        m_centralStack->setCurrentWidget(stackPageFor(m_centralStack, m_detailedSensoryPanel));
```

- [ ] **Step 8: Build the app `-Werror` clean.** From the repo root:

```bash
python tools/decrypt_via_copy.py --apply
mkdir -p build && cd build
PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" /c/Qt/6.10.1/mingw_64/bin/qmake.exe -spec win32-g++ ../DataViewerEnterprise.pro
PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: clean build, no warnings. If `stackPageFor` is flagged unused, a `setCurrentWidget` site was missed — find it via `grep -n "setCurrentWidget(m_centralSplitter\|setCurrentWidget(m_sensoryPanel\|setCurrentWidget(m_detailedSensoryPanel" src/MainWindow.cpp` and convert it.

- [ ] **Step 9: Manual smoke (mode switching still works).** Launch `build/DataViewer.exe`, toggle into Sensory mode, into Detailed Sensory mode, and back to TPM via the ribbon. Verify each page shows and that shrinking the window makes scrollbars appear on the central area rather than clipping. Automated coverage is the Task 19 `--ui-stress` no-clip check.

- [ ] **Step 10: Commit.**

```bash
git add src/MainWindow.cpp tests/tst_scrollhost/tst_scrollhost.cpp
git commit -m "feat(ui): wrap the three central-stack pages in ScrollHost (v2.7.0)

TPM splitter, SensoryPanel, and DetailedSensoryPanel pages are now each
wrapped in a ScrollHost so they scroll instead of clip at small window
sizes. setCurrentWidget call sites resolve the inner widget to its
wrapping page via stackPageFor(). Adds a parent-chain contract guard to
tst_scrollhost.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 3: Wrap the Navigator + Notes docks; lower the window floor

**Files:**
- Modify: `src/MainWindow.h` (add `m_navScrollHost` / `m_notesScrollHost` private members near the dock members)
- Modify: `src/MainWindow.cpp` (ctor `setMinimumSize` ~line 125; Notes dock `notesPane` ~line 1031; Navigator dock `m_sidebarStack` ~line 1444)

Wrap the **whole `m_sidebarStack`** (not just the full panel) so both the full splitter and the 32px icon strip ride inside one ScrollHost — the 32px strip never overflows, so compact mode is unaffected. The Navigator full panel's Properties + Test Averages + image-button stack below the file tree can overflow vertically — that is exactly what the ScrollHost catches.

- [ ] **Step 1: Add the ScrollHost members to `MainWindow.h`.** Near the dock members (after `QDockWidget* m_fileDock;`), add:

```cpp
    // v2.7.0: per-region scroll wrappers exposed to the verification sweep via
    // scrollHostFor() (Task 16). Store the docks' content hosts so the
    // --ui-stress harness can assert the fits-or-scrolls guarantee per region.
    DVE::ScrollHost* m_navScrollHost = nullptr;     // wraps m_sidebarStack
    DVE::ScrollHost* m_notesScrollHost = nullptr;   // wraps the Notes pane
```

Add `namespace DVE { class ScrollHost; }` to the forward-declare block near the top of `MainWindow.h` if not already present.

- [ ] **Step 2: Lower the window floor** in the ctor at ~line 125. Replace:

```cpp
    setMinimumSize(1280, 800);
```

with:

```cpp
    // v2.7.0: a 480x360 floor lets the window corner-snap (quarter screen) and
    // half-split on common monitors. The per-region ScrollHosts guarantee that
    // anything that no longer fits scrolls into reach rather than clipping.
    setMinimumSize(480, 360);
```

(`resize(1600, 900)` on the next line is unchanged — the default opening size stays large. This is the *only* place the floor is set; do not also edit it elsewhere.)

- [ ] **Step 3: Wrap the Notes dock content** in `setupCentralWidget()` at ~line 1031. Replace `m_notesDock->setWidget(notesPane);` with:

```cpp
    // v2.7.0: wrap the Notes content so the sample-nav bar + story panel scroll
    // (both directions) instead of clipping when the dock is narrow/short.
    m_notesScrollHost = ScrollHost::wrap(notesPane);
    m_notesDock->setWidget(m_notesScrollHost);
```

The dock's `setMinimumWidth(220)` stays (it constrains the dock, not the inner content).

- [ ] **Step 4: Wrap the Navigator dock content** in `setupDockPanels()` at ~line 1444. Replace `m_fileDock->setWidget(m_sidebarStack);` with:

```cpp
    // v2.7.0: wrap the whole sidebar stack so the Navigator + Test Averages +
    // Properties + image-button column scroll instead of clipping at small dock
    // heights. The 32px icon-strip page never overflows, so compact mode is
    // unaffected.
    m_navScrollHost = ScrollHost::wrap(m_sidebarStack);
    m_fileDock->setWidget(m_navScrollHost);
```

The dock's `setMinimumWidth(32)` / `setMaximumWidth(QWIDGETSIZE_MAX)` and the compact-mode 32px cap (the `breakpointChanged` lambda, Task 11) act on `m_fileDock`, not the inner stack, so they continue to work. `m_sidebarStack`, `m_sidebarFullPanel`, `m_sidebarIconStrip` are unchanged and still valid.

- [ ] **Step 5: Build `-Werror` clean.** From the repo root:

```bash
python tools/decrypt_via_copy.py --apply
cd build
PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" /c/Qt/6.10.1/mingw_64/bin/qmake.exe -spec win32-g++ ../DataViewerEnterprise.pro
PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: clean build, no warnings.

- [ ] **Step 6: Manual smoke (docks + small window).** Launch `build/DataViewer.exe`. Shrink toward 480×360. Verify: (a) the Navigator dock shows a vertical scrollbar when the Properties/Test-Averages stack no longer fits; (b) the Notes dock scrolls instead of clipping; (c) the window resizes down to ~480×360; (d) compact icon-strip mode still shows the 32px strip with no spurious scrollbar.

- [ ] **Step 7: Commit.**

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(ui): lower window floor to 480x360 and wrap Navigator/Notes docks in ScrollHost (v2.7.0)

setMinimumSize 1280x800 -> 480x360 so the window corner-snaps and
half-splits; the Notes dock content and the full Navigator sidebar stack
are each wrapped in a ScrollHost (stored as members for the verification
sweep) so they scroll instead of clipping at small dock sizes. Icon-strip
compact mode and dock min/max widths are unaffected.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 4: AppTheme font-metric helpers (`lineUnit` / `controlHeight` / `em`)

**Files:**
- Modify: `src/utils/AppTheme.h` (declarations after the font declarations at lines 88–94), `src/utils/AppTheme.cpp` (definitions after the font definitions at lines 56–74), `tests/tst_apptheme/tst_apptheme.cpp` (add slots; flip `QTEST_APPLESS_MAIN`→`QTEST_MAIN`)

Three **pure** functions of the active font (no caching, no static state) so they re-evaluate correctly on OS-scale change. `lineUnit` = one text line's pixel height; `controlHeight` = a single-line control's height (line + vertical padding), replacing hard-coded 20/22/24px heights; `em` = `n` average-char-widths. These are the dependency consumed by the ribbon tasks (6–9) and the fixed→min sweep (Tasks 15–17), so they land first (Phase 1, before the ribbon).

**Critical test-infra note:** `tst_apptheme.cpp` ends in `QTEST_APPLESS_MAIN`, which constructs no application object; `QFontMetrics` needs a `QGuiApplication`/`QApplication`. Switch to `QTEST_MAIN` (the `.pro` already links `QApplication`). The CI runner already sets `QT_QPA_PLATFORM=offscreen` for GUI suites. The pre-existing color tests are unaffected.

- [ ] **Step 1: Write the failing tests.** In `tests/tst_apptheme/tst_apptheme.cpp`, add three slot declarations to `private slots:` (after `testGroupedComparisonUnique();`):

```cpp
    void testLineUnitPositive();
    void testControlHeightGrowsWithFont();
    void testEmScalesWithCount();
```

Add includes after the `#include "AppTheme.h"` line:

```cpp
#include <QFont>
#include <QFontMetrics>
```

Add the slot bodies just above the final `QTEST_*_MAIN` line:

```cpp
void TestAppTheme::testLineUnitPositive()
{
    const QFont small("Segoe UI", 8);
    const QFont big("Segoe UI", 16);
    QVERIFY(AppTheme::lineUnit(small) > 0);
    QVERIFY2(AppTheme::lineUnit(big) > AppTheme::lineUnit(small),
             "lineUnit must increase as the font point size increases");
    QCOMPARE(AppTheme::lineUnit(), AppTheme::lineUnit(AppTheme::fontDefault()));
}

void TestAppTheme::testControlHeightGrowsWithFont()
{
    const QFont small("Segoe UI", 9);
    const QFont big("Segoe UI", 18);
    const int hSmall = AppTheme::controlHeight(small);
    const int hBig   = AppTheme::controlHeight(big);
    QVERIFY2(hBig > hSmall,
             "controlHeight must grow when the font point size grows "
             "(this is the text-scaling clip guard)");
    QCOMPARE(AppTheme::controlHeight(small, 6),
             AppTheme::lineUnit(small) + 6);
    QCOMPARE(AppTheme::controlHeight(small, 12),
             AppTheme::lineUnit(small) + 12);
    QVERIFY(AppTheme::controlHeight(small, 12) > AppTheme::controlHeight(small, 6));
}

void TestAppTheme::testEmScalesWithCount()
{
    const QFont f("Segoe UI", 9);
    const int oneEm  = AppTheme::em(1.0, f);
    const int fiveEm = AppTheme::em(5.0, f);
    QVERIFY(oneEm > 0);
    QVERIFY2(fiveEm > oneEm, "em must scale up with n");
    QCOMPARE(AppTheme::em(3.0, f),
             int(3.0 * QFontMetrics(f).averageCharWidth()));
}
```

Flip the entry point from `QTEST_APPLESS_MAIN(TestAppTheme)` to `QTEST_MAIN(TestAppTheme)`.

- [ ] **Step 2: Run the test — expect a COMPILE failure** (helpers don't exist yet):

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_apptheme && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ tst_apptheme.pro && mingw32-make -j8' < /dev/null
```

Expected: `'lineUnit' is not a member of 'AppTheme'` etc. This is the red state.

- [ ] **Step 3: Declare the helpers in `AppTheme.h`** (after `static QFont fontMono();`, before the icon-helper comment):

```cpp

    // -- Font-metric sizing helpers (v2.7.0 responsive) -----------------------
    // Pure functions of the active font -- no caching, no state -- so they re-
    // evaluate correctly after a DPI/scale change (Qt re-polishes the app font).
    // Use these instead of hard-coded pixel heights so text-bearing controls
    // grow with the font under OS text-scaling rather than clipping.
    static int lineUnit(const QFont& f = fontDefault());                 // one text line's height, px
    static int controlHeight(const QFont& f = fontDefault(), int vPad = 6); // single-line control height
    static int em(qreal n, const QFont& f = fontDefault());              // n average-char-widths, px
```

(`<QFont>` is already included at `AppTheme.h:4`.)

- [ ] **Step 4: Define the helpers in `AppTheme.cpp`.** Add `#include <QFontMetrics>` to the include block (after `#include <QFont>`). After `QFont AppTheme::fontMono() { ... }`, before `QIcon AppTheme::icon(...)`, add:

```cpp

int AppTheme::lineUnit(const QFont& f)
{
    return QFontMetrics(f).height();
}

int AppTheme::controlHeight(const QFont& f, int vPad)
{
    return lineUnit(f) + vPad;
}

int AppTheme::em(qreal n, const QFont& f)
{
    return int(n * QFontMetrics(f).averageCharWidth());
}
```

- [ ] **Step 5: Rebuild and run the test — expect PASS:**

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_apptheme && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ tst_apptheme.pro && mingw32-make -j8 && set QT_QPA_PLATFORM=offscreen && release\tst_apptheme.exe' < /dev/null
```

Expected: `Totals: 10 passed, 0 failed, 0 skipped` (7 pre-existing color + 3 new), including `PASS : TestAppTheme::testControlHeightGrowsWithFont()`.

- [ ] **Step 6: Run the suite to confirm the `QTEST_MAIN` flip didn't regress the SUBDIRS runner:**

```bash
MSYS_NO_PATHCONV=1 cmd /c 'powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1 -Only tst_apptheme' < /dev/null
```

Expected: `tst_apptheme` all-pass. (If `run-tests.ps1` has no `-Only`, run the whole suite and confirm `tst_apptheme` is green.)

- [ ] **Step 7: Commit.**

```bash
git add src/utils/AppTheme.h src/utils/AppTheme.cpp tests/tst_apptheme/tst_apptheme.cpp
git commit -m "feat(theme): add font-metric helpers lineUnit/controlHeight/em (v2.7.0)

Pure functions of the active font (no caching) so clip-prone control heights
track the font under OS text-scaling. These replace hard-coded 20/22/24px
heights used by the ribbon and the fixed->min sweep. tst_apptheme now boots
via QTEST_MAIN (QFontMetrics needs a QApplication) and asserts controlHeight
grows with point size.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 5: HiDPI fractional rounding → PassThrough (before QApplication)

**Files:**
- Modify: `src/main.cpp` (insert immediately before the `QApplication app(argc, argv);` line, ~line 88)

`Qt::HighDpiScaleFactorRoundingPolicy` is latched when the `QApplication` is constructed; setting it afterward is a no-op. `PassThrough` keeps the fractional OS scale (1.25, 1.5) instead of Qt's default `Round` (which snaps 1.25→1.0 / 1.5→2.0 and causes abrupt text-size jumps). High-DPI scaling itself is already on by default in Qt 6 — only the rounding changes.

- [ ] **Step 1: Confirm the QApplication construction site.**

```bash
grep -n "QApplication app(argc, argv);" src/main.cpp
```

Expected: `88:    QApplication app(argc, argv);`

- [ ] **Step 2: Insert the rounding-policy call before construction.** Change the function opening:

```cpp
int main(int argc, char* argv[])
{
    QApplication app(argc, argv);
```

to:

```cpp
int main(int argc, char* argv[])
{
    // v2.7.0: keep the OS's fractional scale factor (1.25 / 1.5) instead of Qt 6's
    // default Round policy, which snaps to integer multiples and makes text jump
    // abruptly between scale steps. MUST be set before the QApplication is built --
    // the policy is latched at construction; setting it later is a no-op.
    QApplication::setHighDpiScaleFactorRoundingPolicy(
        Qt::HighDpiScaleFactorRoundingPolicy::PassThrough);

    QApplication app(argc, argv);
```

`Qt::HighDpiScaleFactorRoundingPolicy` and the static setter live in `<QApplication>` (already included at `src/main.cpp:1`), so no new include is needed.

- [ ] **Step 3: Build to prove it compiles `-Werror` clean.**

```bash
python tools/decrypt_via_copy.py --apply
cd build && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && mingw32-make -j8' < /dev/null
```

Expected: `main.o` recompiles, links clean, no warnings. (Verified visually via the Task 19 `--ui-stress` text-scale matrix; no headless unit test for a process-global DPI policy.)

- [ ] **Step 4: Commit.**

```bash
git add src/main.cpp
git commit -m "feat(ui): set HighDpiScaleFactorRoundingPolicy PassThrough before QApplication

Smooth fractional text scaling at 125/150% OS scale; Qt 6's default Round
policy snapped to integer multiples and caused abrupt jumps. Must precede
QApplication construction (policy is latched there).

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 6: Add a failing ribbon test that pins the ≤2-line wrap + button-fit invariant

> **Sequencing note:** Tasks 6–9 are the "headline View Raw Data bug" (spec Phase 1 §4) but they consume the `AppTheme` helpers from Task 4, so they run after it. They are grouped here at the end of Phase 1.

**Files:**
- Create: `tests/tst_ribbonlayout/tst_ribbonlayout.pro`, `tests/tst_ribbonlayout/tst_ribbonlayout.cpp`
- Modify: `tests/tests.pro` (add `tst_ribbonlayout` to SUBDIRS after `tst_scrollhost`)

This test must be written and failing before any RibbonWidget edit. It exercises the new public static helpers introduced in Task 7, so it will not even compile until Task 7 lands — the intended red state.

- [ ] **Step 1: Create the test `.pro`** via Python (mirrors `tst_apptheme.pro`; adds RibbonWidget + ScrollHost + AppTheme TUs):

```bash
python - <<'PY'
import os
path = "tests/tst_ribbonlayout/tst_ribbonlayout.pro"
content = """QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle

TEMPLATE = app
TARGET = tst_ribbonlayout

INCLUDEPATH += ../../src ../../src/utils ../../src/widgets ../common

SOURCES += tst_ribbonlayout.cpp \\
           ../../src/widgets/RibbonWidget.cpp \\
           ../../src/widgets/ScrollHost.cpp \\
           ../../src/utils/AppTheme.cpp

HEADERS += ../../src/widgets/RibbonWidget.h \\
           ../../src/widgets/ScrollHost.h \\
           ../../src/utils/AppTheme.h
"""
os.makedirs(os.path.dirname(path), exist_ok=True)
if os.path.exists(path):
    os.remove(path)
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
print("wrote", path)
PY
```

- [ ] **Step 2: Create the test source** via Python. It asserts (a) `wrapLabelText` splits "View Raw Data" into exactly 2 fitting lines at the 8pt label font, (b) the rendered button produces ≤2 layout lines at standard scale, (c) the same at simulated 150% (12pt), and (d) the live button fits its bounds:

```bash
python - <<'PY'
import os
path = "tests/tst_ribbonlayout/tst_ribbonlayout.cpp"
content = r'''#include <QtTest/QtTest>
#include <QApplication>
#include <QToolButton>
#include <QFont>
#include <QFontMetrics>
#include <QStringList>
#include <QIcon>
#include "RibbonWidget.h"
#include "ScrollHost.h"

// Count how many visual lines QToolButton's word-wrap would produce for a
// given (already newline-split) label inside a text rectangle of `textW`
// pixels wide, at font `f`. Mirrors QToolButton's Qt::TextWordWrap behavior
// closely enough to detect a clipped 3rd row.
static int wrappedLineCount(const QString& label, int textW, const QFont& f)
{
    const QFontMetrics fm(f);
    int lines = 0;
    for (const QString& hardLine : label.split('\n')) {
        const QRect br = fm.boundingRect(QRect(0, 0, textW, 100000),
                                         Qt::TextWordWrap, hardLine);
        const int h = qMax(1, fm.lineSpacing());
        lines += qMax(1, (br.height() + h - 1) / h);
    }
    return lines;
}

class TstRibbonLayout : public QObject
{
    Q_OBJECT
private slots:
    void wrapSplitsViewRawDataIntoTwoFittingLines() {
        const QFont f = RibbonGroup::largeButtonFont();
        const QString wrapped = RibbonGroup::wrapLabelText("View Raw Data", f);
        const QStringList parts = wrapped.split('\n');
        QCOMPARE(parts.size(), 2);
        const int textW = RibbonGroup::largeButtonTextWidth();
        const QFontMetrics fm(f);
        QVERIFY2(fm.horizontalAdvance(parts[0]) <= textW,
                 qPrintable(QString("line1 '%1' overflows %2px").arg(parts[0]).arg(textW)));
        QVERIFY2(fm.horizontalAdvance(parts[1]) <= textW,
                 qPrintable(QString("line2 '%1' overflows %2px").arg(parts[1]).arg(textW)));
    }

    void renderedButtonNeverExceedsTwoLines_standardScale() {
        const QFont f = RibbonGroup::largeButtonFont();
        const QString wrapped = RibbonGroup::wrapLabelText("View Raw Data", f);
        QVERIFY(wrappedLineCount(wrapped, RibbonGroup::largeButtonTextWidth(), f) <= 2);
    }

    void renderedButtonNeverExceedsTwoLines_at150Percent() {
        QFont f = RibbonGroup::largeButtonFont();
        f.setPointSizeF(f.pointSizeF() * 1.5);
        const QString wrapped = RibbonGroup::wrapLabelText("View Raw Data", f);
        const int textW = RibbonGroup::largeButtonTextWidth(f);
        QVERIFY(wrappedLineCount(wrapped, textW, f) <= 2);
    }

    void liveButtonFitsWithinItsBounds() {
        RibbonGroup grp("Data");
        QToolButton* b = grp.addLargeButton("View Raw Data", QIcon());
        QVERIFY(b != nullptr);
        b->resize(b->minimumSize());
        const QFont f = b->font();
        const int twoLineH = RibbonGroup::largeButtonHeight(f);
        QVERIFY2(b->minimumHeight() >= twoLineH,
                 qPrintable(QString("button min height %1 < required %2")
                                .arg(b->minimumHeight()).arg(twoLineH)));
        QVERIFY(b->text().contains('\n'));
    }
};

QTEST_MAIN(TstRibbonLayout)
#include "tst_ribbonlayout.moc"
'''
if os.path.exists(path):
    os.remove(path)
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
print("wrote", path)
PY
```

- [ ] **Step 3: Register the suite in `tests/tests.pro`.** Change:

```
    tst_notesstory \
    tst_scrollhost
```

to:

```
    tst_notesstory \
    tst_scrollhost \
    tst_ribbonlayout
```

- [ ] **Step 4: Run the test — expect a COMPILE failure (red).** The helpers `wrapLabelText`/`largeButtonFont`/`largeButtonTextWidth`/`largeButtonHeight` don't exist yet:

```bash
python tools/decrypt_via_copy.py --apply
MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && cd tests\tst_ribbonlayout && C:\Qt\6.10.1\mingw_64\bin\qmake.exe tst_ribbonlayout.pro && mingw32-make' < /dev/null
```

Expected: `error: 'wrapLabelText' is not a member of 'RibbonGroup'` (and the other 3).

- [ ] **Step 5: Commit the red test.**

```bash
git add tests/tst_ribbonlayout/tst_ribbonlayout.pro tests/tst_ribbonlayout/tst_ribbonlayout.cpp tests/tests.pro
git commit -m "test(ribbon): pin <=2-line wrap + button-fit invariant for View Raw Data (red)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 7: Balanced font-derived ≤2-line wrap and grow-not-clip large-button sizing

**Files:**
- Modify: `src/widgets/RibbonWidget.h` (RibbonGroup public section after `addLargeButton`; private section)
- Modify: `src/widgets/RibbonWidget.cpp` (rewrite `addLargeButton` lines 103–167; update `setCompactMode` lines 169–186; the ctor `setFixedHeight(98)` is updated in Task 8)

**Root cause (confirmed in source):** `addLargeButton` measures wrap width with `QFontMetrics(btn->font())` — the inherited 9pt app font — but the QSS forces `font-size: 8pt`, and the magic width `74` is wrong for both; it inserts one `\n` at the first-fitting space, then `setFixedSize(80, 76)`. Because height is fixed and word-wrap is on, a too-wide "line 2" re-wraps into a clipped 3rd row. The fix computes the real available text width, picks the single split that makes both lines fit, derives height from `AppTheme` font metrics, and uses `setMinimumSize` + size policy.

New helpers are `static` so the test (and compact-mode path) can call them without a constructed button.

- [ ] **Step 1: Add the header includes to `RibbonWidget.h`.** The new static helpers use `QFont` default arguments in the header (and Task 8 adds defaults that call `AppTheme::fontDefault()` / `fontSmall()`), so the header must include both rather than rely on transitive includes under `-Werror`. After `#include <QFrame>` (line 8) add:

```cpp
#include <QFont>
#include "../utils/AppTheme.h"   // AppTheme::fontDefault()/fontSmall() default args (Task 8)
```

(`AppTheme.h` already includes `<QFont>`, but include it here explicitly anyway — `QFont` is used directly in this header's signatures.)

- [ ] **Step 2: Declare the new helpers + constant in `RibbonWidget.h`** (in `RibbonGroup` public, right after the `addLargeButton` declaration):

```cpp

    // -- Large-button sizing/wrap helpers (font-derived; spec v2.7.0 §4) ------
    // Standard large-button width. Height/wrap derive from the active font so
    // the button grows under OS text-scaling instead of clipping to 3 lines.
    static constexpr int kLargeButtonWidth = 80;

    // The font the large-button LABEL is actually rendered with (8pt Segoe UI,
    // matching the QSS), used for all wrap measurement.
    static QFont largeButtonFont();

    // Real available text width inside the button: width - frame - padding.
    static int largeButtonTextWidth(const QFont& f = largeButtonFont());

    // Minimum button height that holds the 32px icon band + up to two lines of
    // label at font `f` (replaces the hard-coded 76px).
    static int largeButtonHeight(const QFont& f = largeButtonFont());

    // Split `text` into at most two lines that each fit largeButtonTextWidth(f).
    // Picks the split point closest to a balanced halfway split among the splits
    // that fit; if no two-line split fits (very long single word at large
    // scale), returns the text unchanged (the button grows in width instead).
    static QString wrapLabelText(const QString& text, const QFont& f = largeButtonFont());
```

- [ ] **Step 3: Implement the helpers + rewrite `addLargeButton` in `RibbonWidget.cpp`.** Add to the include block (after `#include <QFontMetrics>`): `#include "ScrollHost.h"`, `#include "utils/AppTheme.h"`, `#include <QStyle>`, `#include <QStyleOptionToolButton>`. Replace the entire `addLargeButton` function (lines 103–167) with:

```cpp
QFont RibbonGroup::largeButtonFont()
{
    // Matches the QSS `font-size: 8pt; font-family: 'Segoe UI'` below.
    QFont f("Segoe UI", 8);
    return f;
}

int RibbonGroup::largeButtonTextWidth(const QFont& f)
{
    // Button width minus: 1px QSS border each side, 2px QSS padding each side,
    // plus a 2px QToolButton internal text inset each side. Floor at a few px
    // so a tiny font never yields a non-positive width.
    Q_UNUSED(f);
    const int frameAndPad = 2 * (1 /*border*/ + 2 /*padding*/ + 2 /*tool inset*/);
    return qMax(8, kLargeButtonWidth - frameAndPad);   // 80 - 10 = 70px standard
}

int RibbonGroup::largeButtonHeight(const QFont& f)
{
    // Two label lines + the 32px icon band + the QSS vertical padding/border.
    // AppTheme::lineUnit(f) is QFontMetrics(f).height(); two of them is the
    // 2-line label block. 32 = icon, +10 = 2px padding + 1px border (x2) + 4
    // spacing. At 8pt Segoe UI this evaluates to ~76, preserving today's look;
    // it grows with the font under scaling.
    const int iconBand = 32;
    const int chrome = 10;
    return iconBand + 2 * AppTheme::lineUnit(f) + chrome;
}

QString RibbonGroup::wrapLabelText(const QString& text, const QFont& f)
{
    const QFontMetrics fm(f);
    const int maxW = largeButtonTextWidth(f);

    if (fm.horizontalAdvance(text) <= maxW)
        return text;   // already fits on one line

    // Candidate split points are the space positions. Choose the split where
    // BOTH halves fit AND the split is closest to the visual midpoint, so the
    // two lines are balanced (not ragged splits that leave line 2 wide enough
    // to re-wrap).
    const int mid = text.length() / 2;
    int bestSplit = -1;
    int bestDist  = text.length() + 1;
    for (int i = 1; i < text.length(); ++i) {
        if (text[i] != ' ')
            continue;
        const QString l1 = text.left(i);
        const QString l2 = text.mid(i + 1);
        if (fm.horizontalAdvance(l1) <= maxW && fm.horizontalAdvance(l2) <= maxW) {
            const int dist = qAbs(i - mid);
            if (dist < bestDist) {
                bestDist  = dist;
                bestSplit = i;
            }
        }
    }

    if (bestSplit < 0)
        return text;   // no fitting 2-line split: let the button grow in width

    return text.left(bestSplit) + "\n" + text.mid(bestSplit + 1);
}

QToolButton* RibbonGroup::addLargeButton(const QString& text,
                                         const QIcon&   icon,
                                         const QString& tooltip)
{
    QToolButton* btn = new QToolButton(this);
    btn->setText(text);
    btn->setIcon(icon);
    btn->setIconSize(QSize(32, 32));
    btn->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
    btn->setToolTip(tooltip.isEmpty() ? text : tooltip);
    btn->setAutoRaise(true);
    btn->setFocusPolicy(Qt::NoFocus);

    // Balanced <=2-line wrap measured against the real available text width and
    // the actual label font (8pt), so QToolButton never spawns a clipped 3rd
    // row (the "View Raw Data" overflow bug).
    const QFont labelFont = largeButtonFont();
    btn->setText(wrapLabelText(text, labelFont));

    // Min-size, not fixed: standard scale renders 80x76 (no visual regression);
    // under text-scaling the button grows instead of clipping. Vertical policy
    // Fixed keeps the ribbon row tidy; horizontal Minimum lets a too-long
    // single word widen the button rather than re-wrap to a 3rd line.
    btn->setMinimumSize(kLargeButtonWidth, largeButtonHeight(labelFont));
    btn->setSizePolicy(QSizePolicy::Minimum, QSizePolicy::Fixed);

    btn->setStyleSheet(R"(
        QToolButton {
            background-color: transparent;
            border: 1px solid transparent;
            border-radius: 3px;
            padding: 2px;
            font-family: 'Segoe UI';
            font-size: 8pt;
            color: #1A1A1A;
            text-align: center;
        }
        QToolButton:hover {
            background-color: #CCE4FF;
            border-color: #99CAFF;
        }
        QToolButton:pressed {
            background-color: #99CAFF;
            border-color: #0066CC;
        }
        QToolButton:checked {
            background-color: #CCE4FF;
            border-color: #0066CC;
        }
        QToolButton:disabled {
            color: #AAAAAA;
        }
    )");

    m_largeButtonLayout->addWidget(btn);
    m_largeButtons.append(btn);
    return btn;
}
```

`largeButtonHeight` deliberately does **not** call `AppTheme::controlHeight` (that helper targets single-line input controls); the 2-line button geometry is its own formula.

- [ ] **Step 3 (cont.): Update `setCompactMode`** non-compact branch (lines 176–182). Change:

```cpp
        if (compact) {
            b->setFixedSize(40, 40);
            b->setIconSize(QSize(20, 20));
        } else {
            b->setFixedSize(80, 76);
            b->setIconSize(QSize(32, 32));
        }
```

to:

```cpp
        if (compact) {
            b->setMinimumSize(40, 40);
            b->setMaximumSize(40, 40);
            b->setIconSize(QSize(20, 20));
        } else {
            b->setMaximumSize(QWIDGETSIZE_MAX, QWIDGETSIZE_MAX);
            b->setMinimumSize(kLargeButtonWidth, largeButtonHeight(largeButtonFont()));
            b->setIconSize(QSize(32, 32));
        }
```

(Compact stays genuinely fixed at 40×40 — icon-only buttons have no label to clip; `QWIDGETSIZE_MAX` clears the compact cap before restoring.)

- [ ] **Step 4: Run the test — expect PASS:**

```bash
python tools/decrypt_via_copy.py --apply
MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && cd tests\tst_ribbonlayout && C:\Qt\6.10.1\mingw_64\bin\qmake.exe tst_ribbonlayout.pro && mingw32-make && release\tst_ribbonlayout.exe' < /dev/null
```

Expected: `Totals: 4 passed, 0 failed`. If `liveButtonFitsWithinItsBounds` is short, adjust the `chrome` constant in `largeButtonHeight`, not the test.

- [ ] **Step 5: Commit.**

```bash
git add src/widgets/RibbonWidget.h src/widgets/RibbonWidget.cpp
git commit -m "fix(ribbon): balanced <=2-line wrap + grow-not-clip large buttons (v2.7.0 §4)

Measure wrap against the real text width and the actual 8pt label font;
pick the balanced fitting split so QToolButton never spawns a clipped 3rd
row (the View Raw Data overflow). Large buttons are now setMinimumSize +
size policy (80x76 at standard scale, grow under text-scaling) instead of
setFixedSize.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 8: Font-derive the RibbonGroup / RibbonWidget heights so the group-title row never clips

**Files:**
- Modify: `src/widgets/RibbonWidget.h` (RibbonGroup public — `groupMinimumHeight`; RibbonWidget public — `ribbonMinimumHeight`)
- Modify: `src/widgets/RibbonWidget.cpp` (ctor `setFixedHeight(98)` line 29; title label `setFixedHeight(16)` line 85; RibbonWidget ctor `setFixedHeight(108)` line 310; `setCompactMode` `setFixedHeight(compact ? 64 : 108)` line 409)
- Modify: `tests/tst_ribbonlayout/tst_ribbonlayout.cpp`

The fixed 98/108/16 heights clip the "Data" group title under scaling. Replace with font-derived minimums equal to today's values at 8/9pt that grow with the font. Keep `QSizePolicy::Fixed` vertical so the ribbon doesn't expand to fill.

- [ ] **Step 1: Declare the two height helpers in `RibbonWidget.h`.** In `RibbonGroup` public (after `wrapLabelText`):

```cpp

    // Font-derived minimum group height: large-button band + 1px separator +
    // title line. Equals ~98 at standard scale; grows under text-scaling.
    static int groupMinimumHeight(const QFont& f = largeButtonFont());
```

In `RibbonWidget` public (after `bool isCompactMode() const { ... }`):

```cpp

    // Font-derived minimum ribbon height: tab bar + group band + bottom rule.
    // Equals ~108 at standard scale; grows under text-scaling. Defaults reuse
    // AppTheme's fonts (single source of truth) -- fontDefault() is the 9pt tab
    // font, fontSmall() the 8pt button font -- not freshly-built QFont literals.
    static int ribbonMinimumHeight(const QFont& tabFont = AppTheme::fontDefault(),
                                   const QFont& btnFont = AppTheme::fontSmall());
```

- [ ] **Step 2: Implement `groupMinimumHeight` + use it in the RibbonGroup ctor.** Add above `RibbonGroup::RibbonGroup`:

```cpp
int RibbonGroup::groupMinimumHeight(const QFont& f)
{
    // 4px top margin + button band + 1px separator + title line + 2px slack.
    const int titleH = AppTheme::lineUnit(f);
    return 4 + largeButtonHeight(f) + 1 + titleH + 2;
}
```

In the ctor, replace `setFixedHeight(98);` (line 29) with:

```cpp
    // Font-derived minimum so the group-title row never clips under scaling;
    // equals ~98 at standard scale. Vertical policy stays Fixed.
    setMinimumHeight(groupMinimumHeight());
```

Replace `m_titleLabel->setFixedHeight(16);` (line 85) with:

```cpp
    m_titleLabel->setMinimumHeight(AppTheme::lineUnit(largeButtonFont()) + 2);
```

- [ ] **Step 3: Implement `ribbonMinimumHeight` + use it.** Add above `RibbonWidget::RibbonWidget`:

```cpp
int RibbonWidget::ribbonMinimumHeight(const QFont& tabFont, const QFont& btnFont)
{
    const int tabBarH = AppTheme::lineUnit(tabFont) + 14;
    return tabBarH + RibbonGroup::groupMinimumHeight(btnFont) + 1;
}
```

Replace `setFixedHeight(108);` (line 310) with `setMinimumHeight(ribbonMinimumHeight());`. The `setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);` on line 311 stays. Replace `setFixedHeight(compact ? 64 : 108);` (line 409) in `setCompactMode` with:

```cpp
    if (compact) {
        // Reuse AppTheme's 9pt tab font (single source of truth) rather than a
        // fresh QFont("Segoe UI", 9) literal -- matches ribbonMinimumHeight().
        const int tabBarH = AppTheme::lineUnit(AppTheme::fontDefault()) + 14;
        setMinimumHeight(tabBarH + 40 + 1);
    } else {
        setMinimumHeight(ribbonMinimumHeight());
    }
```

- [ ] **Step 4: Extend the test** with a slot asserting heights grow at 150%. Add the declaration to `private slots:` and the body before the closing `};`:

```cpp
    void groupAndRibbonHeightsGrowWithFont() {
        const QFont base = RibbonGroup::largeButtonFont();
        QFont scaled = base;
        scaled.setPointSizeF(base.pointSizeF() * 1.5);
        QVERIFY(RibbonGroup::groupMinimumHeight(base) >= 90);
        QVERIFY(RibbonGroup::groupMinimumHeight(base) <= 104);
        QVERIFY(RibbonGroup::groupMinimumHeight(scaled) >
                RibbonGroup::groupMinimumHeight(base));
        QVERIFY(RibbonWidget::ribbonMinimumHeight() >=
                RibbonGroup::groupMinimumHeight(base));
    }
```

- [ ] **Step 5: Run the test — expect PASS:**

```bash
python tools/decrypt_via_copy.py --apply
MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && cd tests\tst_ribbonlayout && C:\Qt\6.10.1\mingw_64\bin\qmake.exe tst_ribbonlayout.pro && mingw32-make && release\tst_ribbonlayout.exe' < /dev/null
```

Expected: `Totals: 5 passed, 0 failed`. If `groupMinimumHeight(base)` falls outside 90–104, nudge the chrome constants in `largeButtonHeight`/`groupMinimumHeight` (target ~98 at 8pt), not the test.

- [ ] **Step 6: Commit.**

```bash
git add src/widgets/RibbonWidget.h src/widgets/RibbonWidget.cpp tests/tst_ribbonlayout/tst_ribbonlayout.cpp
git commit -m "fix(ribbon): font-derive group/ribbon heights so the title row never clips

setFixedHeight(98/108/16) -> setMinimumHeight from AppTheme::lineUnit so the
group title (e.g. 'Data') grows with OS text-scaling instead of clipping.
Standard scale stays ~98/108.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 9: Wrap the RibbonTab group row in a horizontal ScrollHost

**Files:**
- Modify: `src/widgets/RibbonWidget.cpp` (`RibbonWidget::addTab` lines 383–387; `setCompactMode` loop lines 400–410; `ScrollHost.h` include already added in Task 7)
- Modify: `tests/tst_ribbonlayout/tst_ribbonlayout.cpp`

Compact-mode (icons-only) is the first-line narrow response; the horizontal `ScrollHost` is the *fallback* so groups never clip off the right edge when even compact mode overflows. Wrap the `RibbonTab` content widget as it is added as a tab page.

- [ ] **Step 1: Wrap the tab in `addTab`.** Replace:

```cpp
RibbonTab* RibbonWidget::addTab(const QString& label)
{
    RibbonTab* tab = new RibbonTab(m_tabs);
    m_tabs->addTab(tab, label);
    return tab;
}
```

with:

```cpp
RibbonTab* RibbonWidget::addTab(const QString& label)
{
    RibbonTab* tab = new RibbonTab(m_tabs);
    // Horizontal scroll fallback: when groups overflow even in compact mode,
    // the row scrolls sideways rather than clipping off-screen. Vertical scroll
    // is disabled (the ribbon is a single fixed-height row). ScrollHost::wrap
    // re-parents `tab` into the returned scroll area.
    ScrollHost* host = ScrollHost::wrap(tab, Qt::Horizontal);
    m_tabs->addTab(host, label);
    return tab;
}
```

Then fix the `setCompactMode` loop (lines 400–410): after wrapping, `qobject_cast<RibbonTab*>(m_tabs->widget(i))` returns `nullptr` (the page is now a `ScrollHost`). Replace:

```cpp
    for (int i = 0; i < m_tabs->count(); ++i) {
        if (auto* tab = qobject_cast<RibbonTab*>(m_tabs->widget(i))) {
            tab->setCompactMode(compact);
        }
    }
```

with:

```cpp
    for (int i = 0; i < m_tabs->count(); ++i) {
        QWidget* page = m_tabs->widget(i);
        RibbonTab* tab = qobject_cast<RibbonTab*>(page);
        if (!tab) {
            if (auto* host = qobject_cast<ScrollHost*>(page))
                tab = qobject_cast<RibbonTab*>(host->widget());
        }
        if (tab)
            tab->setCompactMode(compact);
    }
```

**Cross-file consequence:** grep `qobject_cast<RibbonTab*>` across `src/` (notably `src/MainWindow.cpp`). Any site that fetches a tab page via `m_tabs->widget(i)`/`tabWidget()->widget(i)` and expects a `RibbonTab*` must apply the same reach-through (cast to `ScrollHost*` then `->widget()`). Fix each occurrence in this step.

- [ ] **Step 2: Add a scroll-fallback assertion** to `tests/tst_ribbonlayout.cpp` (declaration in `private slots:` + body). `ScrollHost.h` is already included (Task 6 Step 2):

```cpp
    void groupRowScrollsHorizontallyWhenNarrow() {
        RibbonWidget ribbon;
        RibbonTab* tab = ribbon.addTab("Home");
        RibbonGroup* g = tab->addGroup("Data");
        for (int i = 0; i < 8; ++i)
            g->addLargeButton(QString("Button %1").arg(i), QIcon());
        ribbon.resize(200, 120);
        ribbon.show();
        QVERIFY(QTest::qWaitForWindowExposed(&ribbon));
        QWidget* page = ribbon.tabWidget()->widget(0);
        ScrollHost* host = qobject_cast<ScrollHost*>(page);
        QVERIFY2(host != nullptr, "ribbon tab page is not wrapped in a ScrollHost");
        QCOMPARE(host->verticalScrollBarPolicy(), Qt::ScrollBarAlwaysOff);
    }
```

(`RibbonWidget::tabWidget()` ALREADY EXISTS — `RibbonWidget.h:108` declares `QTabWidget* tabWidget() { return m_tabs; }` (non-const, returns `m_tabs`). Use it as-is; do NOT add another accessor. The Task 16 `scrollHostFor("ribbonGroups")` accessor also calls it.)

- [ ] **Step 3: Run the test — expect PASS:**

```bash
python tools/decrypt_via_copy.py --apply
MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && cd tests\tst_ribbonlayout && C:\Qt\6.10.1\mingw_64\bin\qmake.exe tst_ribbonlayout.pro && mingw32-make && release\tst_ribbonlayout.exe' < /dev/null
```

Expected: `Totals: 6 passed, 0 failed`.

- [ ] **Step 4: Build the whole app `-Werror` clean** (the reach-through cast in MainWindow must compile warning-free):

```bash
python tools/decrypt_via_copy.py --apply
cd build && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8' < /dev/null
```

Expected: clean build, `DataViewer.exe` produced.

- [ ] **Step 5: Standard-scale visual-parity spot check.** Launch at 100% scale on a ≥1100px window; confirm the ribbon is identical to v2.6.0 — large buttons 80×76, "View Raw Data" on 2 lines inside its button, group titles uncut, no horizontal scrollbar on the group row.

- [ ] **Step 6: Commit.**

```bash
git add src/widgets/RibbonWidget.cpp src/widgets/RibbonWidget.h src/MainWindow.cpp tests/tst_ribbonlayout/tst_ribbonlayout.cpp
git commit -m "feat(ribbon): horizontal ScrollHost fallback on the group row (v2.7.0 §4)

Wrap each RibbonTab page in a horizontal-only ScrollHost so groups scroll
sideways instead of clipping off-screen when even compact mode overflows.
setCompactMode and any MainWindow tab-page cast now reach through the
ScrollHost to the RibbonTab.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Phase 2 — Text-scale safety

The `AppTheme` helpers landed in Task 4 (moved into Phase 1 because the ribbon depends on them). Phase 2's remaining work is the `ResponsiveLayout` extension that auto-collapses docks at very-narrow widths — the layout-level companion to the per-region scroll guarantee.

### Task 10: Add the `VeryNarrow` breakpoint to `ResponsiveLayout` (TDD)

**Files:**
- Modify: `src/utils/ResponsiveLayout.h` (enum lines 25–29; thresholds 31–35; accessors 50–52)
- Modify: `src/utils/ResponsiveLayout.cpp` (`recompute` lines 51–61)
- Modify: `tests/tst_responsivelayout/tst_responsivelayout.cpp`

One new breakpoint below `Compact`. Threshold `kVeryNarrowThreshold = 760`: below `kCompactThreshold` (1100) so VeryNarrow is strictly stronger than Compact; above the corner-snap target (480) and the ~800 half-screen region so a half-width split on a 1366-wide laptop (~683px) lands in VeryNarrow; below the sensory-narrow (700) and detailed-narrow (800) reflow points by design (those panel-internal reflows are orthogonal to dock collapse). The aspect-ratio splitter-orientation hook from the spec is **explicitly deferred** — `ResponsiveLayout` continues to track width only; no stub/enum/signal is added for it.

- [ ] **Step 1: Write the failing tests.** Add two slots to `tests/tst_responsivelayout/tst_responsivelayout.cpp` after `thresholdBoundary()`:

```cpp
    void veryNarrowBoundary() {
        static constexpr int kWaitMs = ResponsiveLayout::kDebounceIntervalMs * 5;

        ResponsiveLayout& rl = ResponsiveLayout::instance();
        QWidget w;
        w.resize(900, 800);                 // Compact band (760..1100)
        w.show();
        rl.beginTracking(&w);
        QTRY_COMPARE_WITH_TIMEOUT(rl.currentBreakpoint(), ResponsiveLayout::Compact, kWaitMs);
        QVERIFY(rl.isCompact());
        QVERIFY(!rl.isVeryNarrow());

        QSignalSpy bpSpy(&rl, &ResponsiveLayout::breakpointChanged);

        w.resize(700, 800);                 // Compact -> VeryNarrow
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 1, kWaitMs);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::VeryNarrow);
        QVERIFY(rl.isVeryNarrow());
        QVERIFY(!rl.isCompact());

        w.resize(900, 800);                 // VeryNarrow -> Compact
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 2, kWaitMs);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Compact);
        QVERIFY(rl.isCompact());
        QVERIFY(!rl.isVeryNarrow());
    }

    void veryNarrowExactThreshold() {
        static constexpr int kWaitMs = ResponsiveLayout::kDebounceIntervalMs * 5;

        ResponsiveLayout& rl = ResponsiveLayout::instance();
        QWidget w;
        w.resize(760, 800);                 // exactly at the threshold
        w.show();
        rl.beginTracking(&w);
        QTRY_COMPARE_WITH_TIMEOUT(rl.currentBreakpoint(), ResponsiveLayout::Compact, kWaitMs);

        QSignalSpy bpSpy(&rl, &ResponsiveLayout::breakpointChanged);
        w.resize(759, 800);
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 1, kWaitMs);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::VeryNarrow);
    }
```

- [ ] **Step 2: Run the test, expect a COMPILE failure:**

```bash
MSYS_NO_PATHCONV=1 cmd /c 'powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1 -Only tst_responsivelayout' < /dev/null
```

Expected: `'VeryNarrow' is not a member of 'DVE::ResponsiveLayout'`. (If `run-tests.ps1` lacks `-Only`, build the single suite directly.)

- [ ] **Step 3: Add the enum value, threshold, and accessor to `ResponsiveLayout.h`.** Extend the enum:

```cpp
    enum Breakpoint {
        Standard,    // >= kCompactThreshold (default 1100 px)
        Compact,     // [kVeryNarrowThreshold, kCompactThreshold) -> icons-only ribbon + 32px sidebar strip
        VeryNarrow   // < kVeryNarrowThreshold (default 760 px) -> also auto-collapse both side docks
    };
    Q_ENUM(Breakpoint)
```

Add the threshold after `kCompactThreshold`:

```cpp
    static constexpr int kCompactThreshold = 1100;
    static constexpr int kVeryNarrowThreshold = 760;          // < 760 -> auto-collapse both side docks
    static constexpr int kSensoryNarrowThreshold = 700;        // < 700 -> 1-up cards
```

Add the accessor next to `isCompact()`:

```cpp
    bool isVeryNarrow() const { return m_breakpoint == VeryNarrow; }
```

- [ ] **Step 4: Update `recompute()` to classify three bands** (keep the rest of `recompute` byte-for-byte):

```cpp
void ResponsiveLayout::recompute(int width) {
    Breakpoint newBp;
    if (width < kVeryNarrowThreshold)      newBp = VeryNarrow;
    else if (width < kCompactThreshold)    newBp = Compact;
    else                                   newBp = Standard;

    const bool widthChangedFlag = (width != m_lastWidth);
    const bool bpChanged = (newBp != m_breakpoint);

    m_lastWidth = width;
    m_breakpoint = newBp;

    if (widthChangedFlag) emit widthChanged(width);
    if (bpChanged)        emit breakpointChanged(newBp, width);
}
```

- [ ] **Step 5: Run the test, expect PASS:**

```bash
MSYS_NO_PATHCONV=1 cmd /c 'powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1 -Only tst_responsivelayout' < /dev/null
```

Expected: all `TstResponsiveLayout` slots pass including the two new ones; pre-existing `thresholdBoundary`/`debounceCoalescesResizes` (1000–1500 band) unaffected.

- [ ] **Step 6: Commit.**

```bash
git add src/utils/ResponsiveLayout.h src/utils/ResponsiveLayout.cpp tests/tst_responsivelayout/tst_responsivelayout.cpp
git commit -m "feat(responsive): add VeryNarrow breakpoint (<760px) below Compact

ResponsiveLayout now classifies three width bands: Standard (>=1100),
Compact ([760,1100)), VeryNarrow (<760). Adds kVeryNarrowThreshold and
isVeryNarrow(); recompute() classifies all three while keeping the
existing widthChanged/breakpointChanged emission contract unchanged.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 11: Auto-collapse both side docks on the `VeryNarrow` transition (MainWindow)

**Files:**
- Modify: `src/MainWindow.h` (add `bool m_docksAutoCollapsed = false;` near the dock members; declare `void applyVeryNarrowDockState(bool veryNarrow);`)
- Modify: `src/MainWindow.cpp` (the `breakpointChanged` lambda lines 400–414; add the method definition)

The Notes dock is TPM-only (hidden in sensory/detailed modes). The VeryNarrow handler must not blindly `show()` docks on exit (that would resurrect Notes in sensory mode). It records whether VeryNarrow performed the collapse and restores each dock only to its mode-appropriate visibility.

- [ ] **Step 1: Add the state member + method declaration to `MainWindow.h`.** After `QDockWidget* m_fileDock;`:

```cpp
    // True while the VeryNarrow breakpoint has force-hidden the side docks, so
    // the exit transition only restores docks that VeryNarrow itself collapsed
    // (and never resurrects the TPM-only Notes dock while in a sensory mode).
    bool          m_docksAutoCollapsed = false;
```

Group the method declaration with the other UI helpers (near `void updateRibbonForMode(...)`):

```cpp
    // VeryNarrow responsive rule: collapse both side docks when veryNarrow is
    // true (reclaiming width for the central ScrollHost), and restore them to
    // their current mode-appropriate visibility when it is false. Idempotent.
    void applyVeryNarrowDockState(bool veryNarrow);
```

- [ ] **Step 2: Implement `applyVeryNarrowDockState()` in `MainWindow.cpp`** (next to the other responsive helpers, e.g. immediately before `MainWindow::restoreSettings()`):

```cpp
void MainWindow::applyVeryNarrowDockState(bool veryNarrow)
{
    if (veryNarrow) {
        if (m_docksAutoCollapsed) return;   // already collapsed; idempotent
        m_docksAutoCollapsed = true;
        // Reclaim the side-dock width for the central ScrollHost. hide() keeps
        // the docks out of the layout entirely so the central page gets the
        // full client width before it must scroll.
        if (m_fileDock)  m_fileDock->hide();
        if (m_notesDock) m_notesDock->hide();
        return;
    }

    // Leaving VeryNarrow: only restore docks that WE collapsed, and only to
    // their current mode-appropriate visibility. The Navigator is always
    // reachable; the Notes dock is TPM-only, so it stays hidden in a sensory
    // mode even after the window widens.
    if (!m_docksAutoCollapsed) return;
    m_docksAutoCollapsed = false;
    if (m_fileDock)  m_fileDock->show();
    if (m_notesDock) m_notesDock->setVisible(!m_sensoryMode && !m_detailedSensoryMode);
}
```

- [ ] **Step 3: Wire it into the existing `breakpointChanged` lambda** (lines 402–414). The Compact logic must trigger for both Compact and VeryNarrow; the dock collapse layers on top. Replace the lambda body with:

```cpp
    connect(&DVE::ResponsiveLayout::instance(),
            &DVE::ResponsiveLayout::breakpointChanged,
            this, [this](DVE::ResponsiveLayout::Breakpoint bp, int) {
        const bool compactOrNarrower =
            (bp == DVE::ResponsiveLayout::Compact ||
             bp == DVE::ResponsiveLayout::VeryNarrow);
        if (m_ribbon) m_ribbon->setCompactMode(compactOrNarrower);
        setStatusBreadcrumb(m_lastBreadcrumbSegments);
        if (m_sidebarStack) {
            m_sidebarStack->setCurrentIndex(compactOrNarrower ? 1 : 0);
            m_fileDock->setMinimumWidth(compactOrNarrower ? 32  : 220);
            m_fileDock->setMaximumWidth(compactOrNarrower ? 32  : QWIDGETSIZE_MAX);
        }
        // VeryNarrow (<760): also fully collapse both side docks so the central
        // ScrollHost gets maximum room before it has to scroll.
        applyVeryNarrowDockState(bp == DVE::ResponsiveLayout::VeryNarrow);
    });
```

(Folding Compact + VeryNarrow into `compactOrNarrower` prevents a re-expand regression: the original lambda only matched `Compact`, so moving to `VeryNarrow` would have re-expanded the ribbon and sidebar strip.)

- [ ] **Step 4: Build clean and run the full suite.**

```bash
python tools/decrypt_via_copy.py --apply
cd build && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && mingw32-make -j8' < /dev/null
cd .. && MSYS_NO_PATHCONV=1 cmd /c 'powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1' < /dev/null
```

Expected: `-Werror` clean; full suite green except the documented pre-existing flaky `tst_excelreader`/`tst_dataprocessor`.

- [ ] **Step 5: Verify in the running app.** Open a TPM file, drag narrower past ~760px: both docks disappear and the central area takes the full width (scrolling as-needed). Widen past 760: Navigator returns; Notes returns only in TPM mode. Switch to a sensory mode while narrow, then widen: Notes stays hidden.

- [ ] **Step 6: Commit.**

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(responsive): auto-collapse both side docks at VeryNarrow (<760px)

When the window enters the VeryNarrow band the Navigator and Notes docks
are hidden so the central ScrollHost reclaims their width before it has
to scroll. Exit restores each dock to its mode-appropriate visibility
(Navigator always; Notes only in TPM mode) via m_docksAutoCollapsed. The
existing Compact ribbon/sidebar-strip behavior now spans Compact AND
VeryNarrow (compactOrNarrower), preventing a re-expand at 760.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Phase 3 — Completeness + verification

The per-panel fixed→min sweep, dialog floors, the `MainWindow` test-support accessors, the `tst_responsivelayout` flakiness fix, and the closed-loop `--ui-stress` screenshot/no-clip harness. At extreme aspect ratios (600×1200, 1600×500) the guarantee is "scrolls, not reflows" — content stays reachable via the per-region ScrollHosts; the spec's optional splitter-orientation reflow is explicitly deferred (out of scope for v2.7.0; `ResponsiveLayout` tracks width only).

### Task 12: Failing closed-loop guard test for the sizing sweep

**Files:**
- Create: `tests/tst_sizingsweep/tst_sizingsweep.pro`, `tests/tst_sizingsweep/tst_sizingsweep.cpp`
- Modify: `tests/tests.pro` (add `tst_sizingsweep` to SUBDIRS)

Asserts the headline dialog shrinks below the old fixed floor and that clip-prone text widgets no longer carry a `Fixed` width band. Written first (red), then the sweep makes it green.

- [ ] **Step 1: Write the test `.pro`** via Python. This is a **light** test in the established `tst_responsivelayout.pro` style — it links only the dialog under test and its direct deps, NOT the whole app. `DataCleanupDialog.cpp` includes only `AppTheme.h` (+ `ReportData.h`, which is header-only), and Task 14 wraps it with `ScrollHost`, so the exact minimal SOURCES list is `DataCleanupDialog.cpp` + `AppTheme.cpp` + `ScrollHost.cpp`:

```bash
python - <<'PY'
import os
path = r"tests/tst_sizingsweep/tst_sizingsweep.pro"
content = """QT += core gui widgets testlib
CONFIG += c++17 console
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_sizingsweep

INCLUDEPATH += ../../src ../../src/ui ../../src/utils ../../src/widgets

SOURCES += tst_sizingsweep.cpp \\
           ../../src/ui/DataCleanupDialog.cpp \\
           ../../src/utils/AppTheme.cpp \\
           ../../src/widgets/ScrollHost.cpp

HEADERS += ../../src/ui/DataCleanupDialog.h \\
           ../../src/utils/AppTheme.h \\
           ../../src/widgets/ScrollHost.h \\
           ../../src/pipeline/ReportData.h
"""
os.makedirs(os.path.dirname(path), exist_ok=True)
if os.path.exists(path):
    os.remove(path)
with open(path, "w", encoding="utf-8", newline="\\n") as f:
    f.write(content)
print("wrote", path)
PY
```

(No `app_sources.pri` / `common-sources.pri` — this suite does NOT instantiate `MainWindow`; it constructs `DataCleanupDialog` directly, exactly the way `tst_responsivelayout` links only the few sources it needs. If a future build error names an unresolved symbol from `DataCleanupDialog.cpp`, add that one `.cpp` to the SOURCES list — do not pull in a whole-app include.)

- [ ] **Step 2: Write the test body** via Python:

```bash
python - <<'PY'
import os
path = r"tests/tst_sizingsweep/tst_sizingsweep.cpp"
content = r'''#include <QtTest/QtTest>
#include <QApplication>
#include <QDoubleSpinBox>
#include "ui/DataCleanupDialog.h"
#include "pipeline/ReportData.h"

using namespace DVE;

class TstSizingSweep : public QObject
{
    Q_OBJECT
private slots:
    void cleanupDialogFloorLowered();
    void cleanupDialogControlsNotVerticallyFixed();
};

void TstSizingSweep::cleanupDialogFloorLowered()
{
    SheetResult sheet;
    sheet.sheetName = QStringLiteral("S1");
    SampleResult sample;
    sample.sampleName = QStringLiteral("A");
    DataRow row; row.puffs = 1; row.beforeWeight = 1.0; row.afterWeight = 0.9;
    sample.rows.append(row);
    sheet.samples.append(sample);

    QMap<int, QSet<int>> none;
    DataCleanupDialog dlg(sheet, none);
    QVERIFY2(dlg.minimumHeight() <= 400,
             qPrintable(QStringLiteral("minimumHeight=%1").arg(dlg.minimumHeight())));
    QVERIFY2(dlg.minimumWidth() <= 560,
             qPrintable(QStringLiteral("minimumWidth=%1").arg(dlg.minimumWidth())));
}

void TstSizingSweep::cleanupDialogControlsNotVerticallyFixed()
{
    SheetResult sheet;
    sheet.sheetName = QStringLiteral("S1");
    SampleResult sample;
    sample.sampleName = QStringLiteral("A");
    sheet.samples.append(sample);

    QMap<int, QSet<int>> none;
    DataCleanupDialog dlg(sheet, none);
    const auto spins = dlg.findChildren<QDoubleSpinBox*>();
    QVERIFY(!spins.isEmpty());
    for (auto* sp : spins) {
        const bool fixedBand =
            sp->sizePolicy().horizontalPolicy() == QSizePolicy::Fixed &&
            sp->minimumWidth() == sp->maximumWidth();
        QVERIFY2(!fixedBand, "spin box must be growable, not a fixed width band");
    }
}

int main(int argc, char** argv)
{
    QApplication app(argc, argv);
    TstSizingSweep tc;
    return QTest::qExec(&tc, argc, argv);
}

#include "tst_sizingsweep.moc"
'''
os.makedirs(os.path.dirname(path), exist_ok=True)
if os.path.exists(path):
    os.remove(path)
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
print("wrote", path)
PY
```

- [ ] **Step 3: Register the suite in `tests/tests.pro`** next to the other UI suites (after `tst_ribbonlayout`), matching the one-per-line style.

- [ ] **Step 4: Run it — expect red.** Before the sweep, the floor is 520 and the spins are fixed-width:

```bash
MSYS_NO_PATHCONV=1 cmd /c '.\tests\run-tests.ps1 -Only tst_sizingsweep' < /dev/null
```

Expected: `FAIL! : TstSizingSweep::cleanupDialogFloorLowered() minimumHeight=520`.

- [ ] **Step 5: Commit the red test.**

```bash
git add tests/tst_sizingsweep/ tests/tests.pro
git commit -m "test(responsive): failing sizing-sweep guard (dialog floor + growable controls)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 13: Apply the fixed→min sweep across the catalog

**Files (verify line numbers before editing; they are from the pre-sweep tree):** `src/MainWindow.cpp`, `src/ui/SensoryPanel.cpp`, `src/ui/DetailedSensoryPanel.cpp`, `src/widgets/NotesStoryPanel.cpp`, `src/plotting/PlotWidget.cpp`, `src/ui/PropertiesPanel.cpp`, `src/ui/SamplesCheckboxPanel.cpp`, `src/ui/ReportPreviewDialog.cpp`, `src/ui/DataCleanupDialog.cpp`.

This task is a **transform RULE + a fully-enumerated grouped catalog + 3 worked examples** — deliberately not 50 near-identical blocks (DRY). Every catalog row follows one of the 3 worked patterns verbatim.

**Out of scope (owned elsewhere — do NOT touch here):** `src/widgets/RibbonWidget.cpp` (Tasks 6–9); `src/MainWindow.cpp:125-126` window floor (Task 3) + `src/main.cpp` HiDPI (Task 5); central-page/dock/Navigator wrapping (Tasks 2–3); `*.resize(W,H)` off-screen QChart/QByteArray render buffers in `SensoryReportSource.cpp`, `DetailedSensoryPanel.cpp:1984/1992`, `ZipWriter.cpp` (report PNG geometry is load-bearing — must stay fixed).

#### The Transform Rule

| Pattern | Transform |
|---|---|
| **R1.** `setFixedWidth/Height(px)` on a **text-bearing** widget (QLineEdit, QComboBox, QSpinBox, QLabel-with-text, QPlainTextEdit, QPushButton/QToolButton-with-label) | Replace with the matching `setMinimumWidth/Height`. For single-line heights, prefer `setMinimumHeight(AppTheme::controlHeight())` so it tracks the font; keep the literal only when it is larger than `controlHeight()` at standard scale (multi-line/padded). Default policies are already `Preferred`/`Expanding`; add `setSizePolicy` only where the code relied on `Fixed`. |
| **R2.** `setFixedSize(W, H)` on a text-bearing widget | Replace with `setMinimumSize(W, H)`. No maximum. |
| **R3.** `setFixedSize(N, N)`/`setFixedWidth(N)` on a **genuinely-fixed** element (square icon-strip buttons, fixed thumbnail tiles, icon previews, 1px separators, save-chart icon button) | **LEAVE UNCHANGED** — tagged `KEEP`. |
| **R4.** `setMaximumWidth(px)` existing **only for cross-card/row alignment** | Delete the max; replace with `setMinimumWidth(px)` (grow-allowed). The ScrollHost catches residual overflow. |
| **R5.** A `setMaximumWidth` pairing with a `setMinimumWidth` to form a **genuine layout band** | Keep both unless the catalog row says otherwise. |
| **R6.** Dialog `setMinimumSize(W, H)` floors | Lower per Task 14. |
| **R7.** Dialog/widget `resize(W, H)` / runtime viewport resizes | **LEAVE UNCHANGED** — initial/runtime size, not a floor. |

**Genuinely-fixed allowlist (R3 — never converted):** `MainWindow.cpp:1416` (icon strip 32px), `MainWindow.cpp:1429` (32×32 strip buttons), `PlotWidget.cpp:54` (32×26 icon toggle), `PlotWidget.cpp:228` (1px separator), `ImageInboxDialog.cpp:179` (160×160 preview), `ImageInboxDialog.cpp:765` (96×96 thumbnail), `SensoryPanel.cpp:888` (24×24 save-chart icon), `RibbonWidget.cpp:77/253/378` (1px separators — owned by Tasks 6–9 anyway), `IdentityPromptDialog.cpp:57` (36×36 avatar — already a `setMinimumSize`).

#### Catalog

**`src/MainWindow.cpp`:** 972 `m_prevBtn` `setFixedSize(28,24)`→R2 `setMinimumSize(28, AppTheme::controlHeight())`; 973 `m_nextBtn` same; 1296 `avgHeader` `setFixedHeight(22)`→R1 `setMinimumHeight(AppTheme::controlHeight())`; 1355 `propHeader` `setFixedHeight(22)`→R1 same; 1393 `imgBar` `setFixedHeight(40)`→R1 `setMinimumHeight(40)`; 1416/1429 KEEP; 1526/1527 `m_progressBar` `setMaximumWidth(200)`/`setMaximumHeight(16)`→R5 keep; 3735/3736 `picker.setMinimumSize(500,400)`/`resize(550,450)`→R6 `setMinimumSize(360, 280)` + keep resize.

**`src/ui/SensoryPanel.cpp`:** 230 card base `setFixedWidth(245)`→R1 `setMinimumWidth(245)`; 272/273 `m_voltageEdit` `setFixedWidth(52)`/`setFixedHeight(20)`→R1 `setMinimumWidth(52)`/`setMinimumHeight(AppTheme::controlHeight())`; 280/281 `m_resistanceEdit` same; 288/289 `m_heatingTechCombo` `setFixedWidth(72)`/`setFixedHeight(20)`→R1 `setMinimumWidth(72)`/`setMinimumHeight(controlHeight())`; 303 `m_powerTypeCombo` `setFixedHeight(20)`→R1 `setMinimumHeight(controlHeight())`; 371/372 metric `spin` `setFixedWidth(65)`/`setFixedHeight(22)`→R1; 383/384 metric `lbl` `setMinimumWidth(72)`/`setMaximumWidth(72)`→R4 drop max; 411/412 `m_puffLengthSpin` `setFixedWidth(72)`/`setFixedHeight(20)`→R1; 421 `m_stopwatchBtn` `setFixedHeight(20)`→R1 `setMinimumHeight(controlHeight())`; 448/449 `m_commentsEdit` `setMinimumHeight(36)`/`setMaximumHeight(50)`→R5 keep min, raise max to `AppTheme::controlHeight()*3`; 492/493 `removeBtn` `setFixedWidth(60)`/`setFixedHeight(20)`→R1; 763 each `card` `setFixedWidth(targetCardWidth)`→R1 `setMinimumWidth(targetCardWidth)`; 778 header `edit` `setFixedWidth(150)`→R1 `setMinimumWidth(150)`; 794 `m_roundCombo` `setFixedWidth(60)`→R1 `setMinimumWidth(60)`; 888 KEEP; 153 `popup->setFixedSize(popupW,popupH)`→R2 `setMinimumSize(popupW, popupH)`; 2414/2415, 2593/2594 pickers→R6 `setMinimumSize(420,300)` and `(360,280)` + keep resizes.

**`src/ui/DetailedSensoryPanel.cpp`:** 266 `edit` `setMaximumWidth(maxW)`→R4 drop max, `setMinimumWidth(maxW)`; 355/356 `m_prevBtn`/`m_nextBtn` `setFixedSize(28,24)`→R2 `setMinimumSize(28, controlHeight())`; 492 `spin` `setFixedWidth(70)`→R1; 517 `combo` `setMaximumWidth(kComboMaxW)`→R4; 544 `m_oilSmellSpin` `setFixedWidth(70)`→R1; 591 `m_mouthpieceCombo` `setMaximumWidth(kComboMaxW)`→R4; 609 `m_clogCombo` same→R4; 133 `popup->setFixedSize`→R2; 1984/1992 EXCLUDED (report buffers).

**`src/widgets/NotesStoryPanel.cpp`:** 137 `note` (QPlainTextEdit) `setFixedHeight(48)`→R1 `setMinimumHeight(48)` + `setSizePolicy(QSizePolicy::Preferred, QSizePolicy::Minimum)`; 185 `clog` `setFixedWidth(34)`→R1 `setMinimumWidth(34)`; 198 `smell` same; 351 collapsible `table` `setFixedHeight(...)`→**SPECIAL** (see worked Example 3 — keep the fixed-to-exact-rows policy, only swap a hard-coded `rowH` literal to `AppTheme::controlHeight()`).

**`src/plotting/PlotWidget.cpp`:** 48 `m_plotTypeCombo` `setMaximumWidth(200)`→R5 keep; 54 KEEP; 82 `m_regimeCombo` `setMaximumWidth(200)`→R5 keep; 100 `topBar` `setFixedHeight(40)`→R1 `setMinimumHeight(40)`; 115 `m_checkboxScrollArea` `setFixedHeight(32)`→R1 `setMinimumHeight(AppTheme::controlHeight() + 6)`; 228 KEEP; 413/422/428 runtime `m_plotLabel->resize(...)`→R7 keep.

> **Plot top-bar honest note (spec §4 lists it as a no-scroll region):** this sweep only converts the PlotWidget top-bar's `setFixedHeight(40)` to `setMinimumHeight(40)` — it does NOT give the top-bar its own horizontal scroll. The top-bar's narrow-width guarantee is "fixed→min growth + the central-page ScrollHost" (Task 2 wraps the whole TPM splitter, which includes the plot, in a ScrollHost): if the combos+buttons row outgrows the window, the central ScrollHost scrolls the plot region into reach. It does not reflow. If a future `--ui-stress` case shows the top-bar row itself clipping inside the plot area (combos cut off with no scrollbar), the follow-up is to wrap just that one top-bar row in a horizontal `ScrollHost::wrap(topBar, Qt::Horizontal)` — that is a one-line addition, deliberately not done now to avoid an unneeded nested scroll at standard sizes.

**`src/ui/PropertiesPanel.cpp`:** 21 `s` font-name combo `setFixedWidth(75)`→R1 `setMinimumWidth(75)`; 74 `m_fontSize` `setFixedWidth(75)`→R1 same.

**`src/ui/SamplesCheckboxPanel.cpp`:** 14 panel `setMaximumWidth(200)`→R4 drop max, `setMinimumWidth(200)`.

**`src/ui/ReportPreviewDialog.cpp`:** 122 `m_thumbList` `setFixedWidth(180)`→R1 `setMinimumWidth(180)`; 175 `m_propsPanel` `setMaximumWidth(180)`→R5 keep; 93 `resize(1600,720)`→R7 keep.

#### Worked Example 1 — R1 single-line control height (covers ~20 height sites)

`SensoryPanel.cpp` lines 272-273, the voltage edit. **Before:**

```cpp
    m_voltageEdit = new QLineEdit;
    m_voltageEdit->setFixedWidth(52);
    m_voltageEdit->setFixedHeight(20);
    m_voltageEdit->setPlaceholderText("0.00");
    m_voltageEdit->setStyleSheet("font-size: 7pt;");
```

**After:**

```cpp
    m_voltageEdit = new QLineEdit;
    m_voltageEdit->setMinimumWidth(52);
    m_voltageEdit->setMinimumHeight(AppTheme::controlHeight());
    m_voltageEdit->setPlaceholderText("0.00");
    m_voltageEdit->setStyleSheet("font-size: 7pt;");
```

Every R1 height site takes exactly this shape.

#### Worked Example 2 — R4 alignment-only width cap

`SensoryPanel.cpp` lines 383-384. **Before:**

```cpp
            lbl->setMinimumWidth(72);
            lbl->setMaximumWidth(72);
```

**After:**

```cpp
            lbl->setMinimumWidth(72);
```

The grid/box layout keeps the column aligned to the widest natural label. Apply identically at `DetailedSensoryPanel.cpp:266/517/591/609` and `SamplesCheckboxPanel.cpp:14`.

#### Worked Example 3 — NotesStoryPanel collapsible table (the one SPECIAL case)

`NotesStoryPanel.cpp:137` (the note editor — straightforward R1). **After:**

```cpp
    note->setMinimumHeight(48);
    note->setSizePolicy(QSizePolicy::Preferred, QSizePolicy::Minimum);
```

`NotesStoryPanel.cpp:351` (collapsible table). This is NOT a clip — the height is computed to show all `nRows` with no inner vertical scrollbar, by design. Keep that intent; only make the per-row term scale-aware **if** `rowH` is a hard-coded literal. Inspect `rowH` above line 351:
- If `rowH` is already `table->verticalHeader()->defaultSectionSize()` or font-derived → leave line 351 unchanged.
- If `rowH` is a literal (e.g. `const int rowH = 22;`) → change its definition to `const int rowH = AppTheme::controlHeight();` and leave line 351's expression as-is.

#### Sweep steps

- [ ] **Step 1: Decrypt.** `python tools/decrypt_via_copy.py --apply` from the repo root.
- [ ] **Step 2: Apply the catalog.** Edit each file per its rows, using the 3 worked patterns. For each file add `#include "../utils/AppTheme.h"` only if absent (check the include block). Do NOT touch any `KEEP`/`EXCLUDED` site.
- [ ] **Step 3: Build `-Werror` clean.**

```bash
MSYS_NO_PATHCONV=1 cmd /c 'cd build && set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8' < /dev/null
```

Expected: clean compile (no unused-include or narrowing warnings).

- [ ] **Step 4: Run the full suite + the sizing guard.**

```bash
MSYS_NO_PATHCONV=1 cmd /c '.\tests\run-tests.ps1' < /dev/null
```

Expected: `tst_sizingsweep` PASSES (floor 520→≤400, spins growable); previously-green suites stay green (modulo known-flaky `tst_excelreader`/`tst_dataprocessor`).

- [ ] **Step 5: Commit.**

```bash
git add src/
git commit -m "refactor(responsive): fixed->min sweep across panels (Phase 3 catalog)

Convert clip-prone setFixed*/alignment-only setMaximum* sites to
min-size + growable policy per the v2.7.0 responsive spec. Single-line
control heights now track AppTheme::controlHeight() so they grow under OS
text-scaling instead of clipping. Genuinely-fixed icon strips, separators,
thumbnails, and off-screen report render buffers left fixed. No visual
change at standard scale (old value kept as min).

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 14: Dialog floors + ScrollHost wrap

**Files:** `src/ui/DataCleanupDialog.cpp` (38), `src/ui/HeaderEditDialog.cpp` (20, 31-34, 77-78), `src/ui/ImageViewDialog.cpp` (309, 339), `src/ui/ImageInboxDialog.cpp` (99, 116, 122, 125, 272, 274, 276), `src/ui/DatabaseBrowserDialog.cpp` (158).

Per R6: lower the `setMinimumSize` floor; wrap fixed-layout content in `ScrollHost` so the lowered floor never hides controls. Dialogs that already have a scrolling top-level (`DatabaseBrowserDialog` QTabWidget, `HeaderEditDialog` raw `QScrollArea`) only need the floor lowered (and HeaderEditDialog's raw `QScrollArea` swapped to `ScrollHost::wrap` for both-direction as-needed).

| File | Line | Current floor | New floor | Wrap |
|---|---|---|---|---|
| `DataCleanupDialog.cpp` | 38 | `setMinimumSize(700, 520)` | `setMinimumSize(520, 360)` | splitter+table scroll internally; no extra wrap |
| `HeaderEditDialog.cpp` | 20 | `setMinimumSize(420, 420)` | `setMinimumSize(360, 300)` | swap raw `QScrollArea` (31-34) → `ScrollHost::wrap(content)` |
| `ImageViewDialog.cpp` | 309 | `setMinimumSize(860, 580)` | `setMinimumSize(560, 400)` | `m_view` (QGraphicsView) scrolls |
| `ImageViewDialog.cpp` | 339 | `m_view->setMinimumSize(820, 465)` | `m_view->setMinimumSize(480, 320)` | — |
| `ImageInboxDialog.cpp` | 99 | `setMinimumSize(720, 620)` | `setMinimumSize(560, 420)` | tree + thumb grid scroll internally |
| `DatabaseBrowserDialog.cpp` | 158 | `setMinimumSize(900, 500)` | `setMinimumSize(560, 380)` | top-level QTabWidget scrolls each table |

`ImageInboxDialog.cpp` width-cap sites (116/122/125/272/274/276) are R1 button-width caps — convert `setFixedWidth(px)`→`setMinimumWidth(px)` in the same pass.

#### Worked Example — HeaderEditDialog (raw QScrollArea → ScrollHost::wrap)

**Before** (lines 20, 31-34, 77-78):

```cpp
    setMinimumSize(420, 420);
    ...
    QScrollArea* scroll = new QScrollArea(this);
    scroll->setWidgetResizable(true);
    QWidget* content = new QWidget;
    ...
    scroll->setWidget(content);
    mainLayout->addWidget(scroll);
```

**After:**

```cpp
    setMinimumSize(360, 300);
    ...
    // Scrollable content (both-direction as-needed via the shared ScrollHost)
    QWidget* content = new QWidget;
    ...
    ScrollHost* scroll = ScrollHost::wrap(content);
    mainLayout->addWidget(scroll);
```

Remove the now-dead `scroll->setWidgetResizable(true)` and `scroll->setWidget(content)` lines (`ScrollHost::wrap` does both). Add `#include "../widgets/ScrollHost.h"`. `QScrollArea` is no longer referenced after this change, so **remove** `#include <QScrollArea>` to stay `-Werror` clean.

#### Steps

- [ ] **Step 1: Decrypt.** `python tools/decrypt_via_copy.py --apply`.
- [ ] **Step 2: Edit each dialog** per the table: lower floors, swap HeaderEditDialog's `QScrollArea`→`ScrollHost::wrap`, convert ImageInboxDialog button width caps. Add `#include "../widgets/ScrollHost.h"` only to files that now call `ScrollHost::wrap` (HeaderEditDialog).
- [ ] **Step 3: Build `-Werror` clean** (same command as Task 13 Step 3). Watch for an unused-include warning on the removed `<QScrollArea>`.
- [ ] **Step 4: Run the guard.**

```bash
MSYS_NO_PATHCONV=1 cmd /c '.\tests\run-tests.ps1 -Only tst_sizingsweep' < /dev/null
```

Expected: `cleanupDialogFloorLowered()` PASSES (`minimumHeight=360 ≤ 400`, `minimumWidth=520 ≤ 560`).

- [ ] **Step 5: Commit.**

```bash
git add src/ui/
git commit -m "refactor(responsive): lower dialog floors + ScrollHost wrap (Phase 3)

DataCleanup/HeaderEdit/ImageView/ImageInbox/DatabaseBrowser dialogs now
floor at quarter-screen-friendly minimums; HeaderEditDialog uses the
shared ScrollHost for both-direction as-needed scrolling. No standard-size
visual change (resize() initial sizes unchanged).

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 15: Verify no clip-prone fixed height remains on text widgets

**Files:** none (verification only).

- [ ] **Step 1: Grep that no on-screen text widget keeps a `setFixedHeight`/`setFixedSize` literal** beyond the KEEP allowlist:

```bash
MSYS_NO_PATHCONV=1 cmd /c 'rg -n "setFixedHeight|setFixedSize" src/MainWindow.cpp src/ui/SensoryPanel.cpp src/ui/DetailedSensoryPanel.cpp src/widgets/NotesStoryPanel.cpp src/plotting/PlotWidget.cpp src/ui/DataCleanupDialog.cpp src/ui/HeaderEditDialog.cpp src/ui/ImageViewDialog.cpp src/ui/ImageInboxDialog.cpp src/ui/DatabaseBrowserDialog.cpp src/ui/PropertiesPanel.cpp src/ui/ReportPreviewDialog.cpp src/ui/SamplesCheckboxPanel.cpp' < /dev/null
```

Expected ONLY the KEEP allowlist:
- `src/MainWindow.cpp:1429: btn->setFixedSize(32, 32);`
- `src/plotting/PlotWidget.cpp:54: btn->setFixedSize(32, 26);`
- `src/ui/SensoryPanel.cpp:888: saveChartBtn->setFixedSize(24, 24);`
- `src/ui/ImageInboxDialog.cpp:179: m_previewLabel->setFixedSize(160, 160);`
- `src/ui/ImageInboxDialog.cpp:765: thumb->setFixedSize(96, 96);`

Any `setFixedHeight(<literal>)` on a QLineEdit/QComboBox/QSpinBox/QLabel/QPushButton/QToolButton/QPlainTextEdit that appears was missed — convert it (R1) and re-run.

- [ ] **Step 2: Confirm the NotesStoryPanel table case is intentional.** `rg -n "setFixedHeight" src/widgets/NotesStoryPanel.cpp` should show ONLY the collapsible table's exact-fit height (~line 351). Line 137 (`note`) must now be `setMinimumHeight`.

- [ ] **Step 3: Final full-suite green + `-Werror`.** Re-run `MSYS_NO_PATHCONV=1 cmd /c '.\tests\run-tests.ps1' < /dev/null`; confirm `tst_sizingsweep` green and no new failures. No commit (verification only).

---

### Task 16: Add `MainWindow` test-support accessors for the wrapped regions

The `--ui-stress` harness (Tasks 18–19) inspects, per region, the wrapping `ScrollHost` and the content widget to compute its closed-loop no-clip pass/fail. Add two small public const accessors keyed by a stable string — return existing pointers, no new state. (These are consumed by `--ui-stress` inside the real `DataViewer.exe`, not by any headless Qt Test — the Qt Tests stay light.)

**Files:** `src/MainWindow.h` (public section), `src/MainWindow.cpp` (definitions).

- [ ] **Step 1: Declare the accessors in `MainWindow.h`** (public section):

```cpp
    // --- Test-support accessors (v2.7.0 responsive-UI verification) ---
    // Return the per-region ScrollHost / content widget that the --ui-stress
    // harness inspects to compute its closed-loop no-clip pass/fail. regionKey
    // is one of: "central", "navigator", "notes", "ribbonGroups".
    // Returns nullptr for an unknown key or a not-yet-constructed lazy panel.
    // These expose existing pointers only; they create nothing and have no
    // side effects.
    DVE::ScrollHost* scrollHostFor(const QString& regionKey) const;
    QWidget*         regionWidget(const QString& regionKey) const;
```

(The `DVE::ScrollHost` forward declaration was added in Task 3.)

- [ ] **Step 2: Define the accessors in `MainWindow.cpp`** (near the other trivial getters). The central host is resolved from the stack's current page (Task 2 wrapped each page in a ScrollHost); the navigator/notes hosts are the members from Task 3; the ribbon group ScrollHost is reached through the ribbon's active tab page:

```cpp
DVE::ScrollHost* MainWindow::scrollHostFor(const QString& regionKey) const
{
    if (regionKey == QLatin1String("central")) {
        return m_centralStack
            ? qobject_cast<DVE::ScrollHost*>(m_centralStack->currentWidget())
            : nullptr;
    }
    if (regionKey == QLatin1String("navigator")) return m_navScrollHost;
    if (regionKey == QLatin1String("notes"))     return m_notesScrollHost;
    if (regionKey == QLatin1String("ribbonGroups")) {
        if (!m_ribbon || !m_ribbon->tabWidget()) return nullptr;
        return qobject_cast<DVE::ScrollHost*>(
            m_ribbon->tabWidget()->currentWidget());
    }
    return nullptr;
}

QWidget* MainWindow::regionWidget(const QString& regionKey) const
{
    if (regionKey == QLatin1String("central"))   return m_centralStack;
    if (regionKey == QLatin1String("navigator")) return m_sidebarFullPanel;
    if (regionKey == QLatin1String("notes"))     return m_notesDock ? m_notesDock->widget() : nullptr;
    if (regionKey == QLatin1String("ribbonGroups")) return m_ribbon;
    return nullptr;
}
```

(`RibbonWidget::tabWidget()` already exists — `RibbonWidget.h:108`, non-const, returns `m_tabs`.)

- [ ] **Step 3: Verify the app builds `-Werror` clean.**

```bash
python tools/decrypt_via_copy.py --apply
cd build && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && mingw32-make -j8' < /dev/null
```

Expected: clean build (trivial const getters).

- [ ] **Step 4: Commit.**

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(mainwindow): test-support scrollHostFor/regionWidget accessors for --ui-stress

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 17: Stabilize `tst_responsivelayout` (singleton-reset flakiness fix)

**Diagnosis of the existing flakiness:** `ResponsiveLayout` is a heap singleton; `recompute()` emits only on a delta vs retained `m_lastWidth`/`m_breakpoint`, and `stopTracking()` does not reset that state, nor does it drain an already-queued debounce `timeout()`. Slot-run order then changes residual state → intermittent `QSignalSpy` counts. **Fix:** in `cleanup()`, after `stopTracking()`, drain the event loop and re-seed the singleton to a known baseline.

This suite stays **light** — it links only `ResponsiveLayout.cpp` (its existing `.pro` is unchanged) and drives the singleton with plain `QWidget`s. It does **not** instantiate `MainWindow`: `MainWindow`'s constructor opens a Postgres/offline-snapshot DB connection and spawns a Python subprocess, which is heavy and flaky inside a headless Qt Test. The full-window "no region clipped without a scrollbar" sweep is verified by the `--ui-stress` harness instead (Tasks 19–20), which runs inside the real `DataViewer.exe` where a fully-constructed `MainWindow` exists with live DB/python context.

**Files:** `tests/tst_responsivelayout/tst_responsivelayout.cpp` only (the `.pro` is left as-is — it already links `../../src/utils/ResponsiveLayout.cpp` + its `.h` and nothing else).

- [ ] **Step 1: Implement the singleton-reset in `cleanup()`** (replace the existing `cleanup()`):

```cpp
void cleanup() {
    ResponsiveLayout& rl = ResponsiveLayout::instance();
    rl.stopTracking();
    // Drain any debounce timeout queued before stopTracking().
    QCoreApplication::processEvents();
    QCoreApplication::sendPostedEvents(nullptr, QEvent::Timer);
    // Re-seed the singleton to a known Standard baseline so the next test does
    // not inherit this test's m_lastWidth / m_breakpoint.
    QWidget seed;
    seed.resize(1280, 800);
    rl.beginTracking(&seed);              // recompute(1280) -> Standard
    rl.stopTracking();
    QVERIFY(!rl.isCompact());
}
```

(`#include <QEvent>` / `#include <QCoreApplication>` are needed for `sendPostedEvents(..., QEvent::Timer)`; add them to the top of the file if absent.)

- [ ] **Step 2: Add `singletonStateResetsBetweenWindows()`** — the flakiness-fix regression guard. It uses only plain `QWidget`s and the singleton (no `MainWindow`). Add the declaration to `private slots:` and the body:

```cpp
void singletonStateResetsBetweenWindows() {
    ResponsiveLayout& rl = ResponsiveLayout::instance();

    QWidget narrow;
    narrow.resize(900, 700);
    narrow.show();
    rl.beginTracking(&narrow);
    QTRY_COMPARE_WITH_TIMEOUT(rl.currentBreakpoint(),
        ResponsiveLayout::Compact,
        ResponsiveLayout::kDebounceIntervalMs * 5);

    rl.stopTracking();
    QCoreApplication::processEvents();
    QWidget seed; seed.resize(1280, 800);
    rl.beginTracking(&seed);
    rl.stopTracking();
    QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Standard);

    QWidget again; again.resize(900, 700); again.show();
    QSignalSpy bpSpy(&rl, &ResponsiveLayout::breakpointChanged);
    rl.beginTracking(&again);
    QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 1,
        ResponsiveLayout::kDebounceIntervalMs * 5);
    QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Compact);
}
```

(The `VeryNarrow` breakpoint itself is covered by Task 10's `veryNarrowBoundary()` / `veryNarrowExactThreshold()` slots, also singleton-only. The dock auto-collapse at very-narrow widths is verified in-process by `--ui-stress` via the per-case `nav_visible`/`notes_visible` JSON fields — not here, because asserting dock visibility needs a real `MainWindow`.)

- [ ] **Step 3: Build + run the suite, expect green.**

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_responsivelayout && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ tst_responsivelayout.pro && mingw32-make -j8 && release\tst_responsivelayout.exe' < /dev/null
```

Expected: `Totals: 6 passed, 0 failed` (the 3 pre-existing + 2 from Task 10's `VeryNarrow` + this one).

- [ ] **Step 4: Run the suite 10× to confirm the flakiness is gone.**

```bash
cd tests/tst_responsivelayout && for i in $(seq 1 10); do ./release/tst_responsivelayout.exe || echo "FAIL on run $i"; done
```

Expected: 10 clean passes.

- [ ] **Step 5: Commit.**

```bash
git add tests/tst_responsivelayout/tst_responsivelayout.cpp
git commit -m "test(responsive): singleton-reset in cleanup() fixes tst_responsivelayout flakiness

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 18: `--ui-stress` screenshot harness skeleton + flag wiring

A new CLI mode mirroring `--self-test`: cycle a real `MainWindow` through a (window-size × text-scale) matrix, compute a per-case no-clip pass/fail, `grab()` each case to a PNG, write a JSON index, exit. Logic lives in a new `DVE::runUiStress` so `main.cpp` stays a thin dispatcher. This is the closed-loop verification of the full-window "nothing clipped without a scrollbar" guarantee — it runs inside the real `DataViewer.exe`, where `MainWindow` exists with live DB/python context, which is why that guarantee is **not** asserted in a headless Qt Test.

**Files:** Create `src/utils/UiStress.h`, `src/utils/UiStress.cpp`; modify `src/main.cpp`, `DataViewerEnterprise.pro` (add `UiStress.cpp` to SOURCES, `UiStress.h` to HEADERS).

- [ ] **Step 1: Create the header `src/utils/UiStress.h`** (Python rewrite):

```cpp
#pragma once

#include <QString>

namespace DVE {

/// Run the responsive-UI stress screenshot harness and return a process
/// exit code.
///
/// Cycles a real MainWindow through a matrix of (window size x text-scale
/// factor) presets. For each case it resizes the window, applies the scale
/// by multiplying the base application font point size, lets the layout
/// settle, computes a closed-loop per-region no-clip verdict (each wrapped
/// region either fits its ScrollHost viewport or scrolls in the overflow
/// direction), calls QWidget::grab(), and writes one PNG to outDir. Also
/// writes an index.json describing every case (label, size, scale, png path,
/// grab dimensions, the no-clip pass/fail + any failing regions, and the
/// side-dock visibility).
///
/// outDir defaults to %TEMP%/dve_ui_stress when empty. No GUI interaction;
/// the window is shown (required for a valid grab) but never raised.
///
/// Returns 0 only if every case both grabbed AND passed the no-clip check;
/// 1 otherwise.
int runUiStress(const QString& outDir);

} // namespace DVE
```

- [ ] **Step 2: Create a minimal `UiStress.cpp` skeleton, wire the flag, prove it runs.** Minimal body:

```cpp
#include "UiStress.h"
namespace DVE {
int runUiStress(const QString&) { return 1; }
} // namespace DVE
```

In `src/main.cpp`, add the include (next to `#include "utils/SelfTest.h"`):

```cpp
#include "utils/UiStress.h"
```

Add the options (next to the self-test options):

```cpp
QCommandLineOption optUiStress("ui-stress",
    "Capture a (window size x text-scale) screenshot matrix and exit");
QCommandLineOption optUiStressOut("ui-stress-out",
    "Directory for --ui-stress PNGs + index.json "
    "(default %TEMP%\\dve_ui_stress)", "dir");
```

Register them (next to `parser.addOption(optSelfTestOut);`):

```cpp
parser.addOption(optUiStress);
parser.addOption(optUiStressOut);
```

Add the dispatch (immediately after the self-test dispatch block):

```cpp
// Handle --ui-stress (responsive screenshot matrix, no GUI interaction)
if (parser.isSet(optUiStress)) {
    return DVE::runUiStress(parser.value(optUiStressOut));
}
```

Add `UiStress.cpp` to the `SOURCES +=` block and `UiStress.h` to the `HEADERS +=` block in `DataViewerEnterprise.pro` (one `    src/utils/UiStress.cpp \` / `    src/utils/UiStress.h \` line each, matching the explicit one-per-line style — the `.pro` is NOT a wildcard). Build + smoke:

```bash
python tools/decrypt_via_copy.py --apply
cd build && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8' < /dev/null
./release/DataViewer.exe --ui-stress; echo "exit=$?"
```

Expected: builds clean; `exit=1` (skeleton) — confirms the flag is parsed and dispatched.

- [ ] **Step 3: Commit the skeleton.**

```bash
git add src/utils/UiStress.h src/utils/UiStress.cpp src/main.cpp DataViewerEnterprise.pro
git commit -m "feat(verify): wire --ui-stress / --ui-stress-out CLI flags (skeleton)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 19: Implement the real `--ui-stress` harness

**Files:** `src/utils/UiStress.cpp`.

- [ ] **Step 1: Implement the real harness** (Python rewrite of `src/utils/UiStress.cpp`):

```cpp
#include "UiStress.h"

#include "../MainWindow.h"
#include "../widgets/ScrollHost.h"
#include "AppTheme.h"

#include <QApplication>
#include <QCoreApplication>
#include <QDir>
#include <QDockWidget>
#include <QEventLoop>
#include <QFile>
#include <QFont>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QPixmap>
#include <QSize>
#include <QStringList>
#include <QStyle>
#include <QTextStream>
#include <QTimer>
#include <QWidget>

namespace DVE {

namespace {

// Spin the event loop for ms milliseconds without blocking signal/slot
// delivery (so the responsive debounce fires and child layouts re-flow).
void settle(int ms)
{
    QEventLoop loop;
    QTimer::singleShot(ms, &loop, &QEventLoop::quit);
    loop.exec();
}

// The closed-loop no-clip check, run inside the real DataViewer.exe (where a
// fully-constructed MainWindow with live DB/python context exists -- which is
// why this guarantee is verified HERE rather than in a headless Qt Test). For
// each wrapped region, the content either fits its ScrollHost viewport or the
// host's scrollbar is active in the overflow direction. Returns false (and
// appends a human-readable reason to `failures`) if any region is clipped
// without a scrollbar. Collapsed docks / not-yet-built lazy panels are skipped.
bool regionsFitOrScroll(const MainWindow& w, QStringList& failures)
{
    static const QStringList kRegions = {
        QStringLiteral("central"),
        QStringLiteral("navigator"),
        QStringLiteral("notes"),
        QStringLiteral("ribbonGroups"),
    };
    bool ok = true;
    for (const QString& key : kRegions) {
        ScrollHost* host = w.scrollHostFor(key);
        QWidget* content = w.regionWidget(key);
        if (!host || !content || !host->isVisible())
            continue;   // collapsed dock / lazy panel: nothing on screen to clip
        const QSize vp   = host->viewport()->size();
        const QSize need = content->sizeHint().expandedTo(content->minimumSizeHint());
        const bool hOk = (need.width()  <= vp.width())  || host->scrollbarActive(Qt::Horizontal);
        const bool vOk = (need.height() <= vp.height()) || host->scrollbarActive(Qt::Vertical);
        if (!hOk) {
            ok = false;
            failures << QStringLiteral("%1 H need=%2 vp=%3")
                            .arg(key).arg(need.width()).arg(vp.width());
        }
        if (!vOk) {
            ok = false;
            failures << QStringLiteral("%1 V need=%2 vp=%3")
                            .arg(key).arg(need.height()).arg(vp.height());
        }
    }
    return ok;
}

} // anonymous namespace

int runUiStress(const QString& outDirArg)
{
    const QString outDir = outDirArg.isEmpty()
        ? (QDir::tempPath() + "/dve_ui_stress")
        : outDirArg;
    if (!QDir().mkpath(outDir)) {
        QTextStream(stderr) << "ui-stress: cannot create output dir "
                            << outDir << Qt::endl;
        return 1;
    }

    // The (window-size x text-scale) matrix: corner-snap quarter, half-split,
    // the window floor, and two extreme aspect ratios; scales mimic OS text-
    // scaling at 100/125/150/200%.
    const QList<QSize> sizes = {
        {1920, 1080}, {1280, 800}, {960, 540},
        {800, 600},   {480, 360},  {600, 1200}, {1600, 500},
    };
    const QList<qreal> scales = { 1.0, 1.25, 1.5, 2.0 };

    // Apply the theme ONCE, up front. AppTheme::apply() resets the application
    // font to AppTheme::fontDefault() (9pt), so it must NOT run inside the scale
    // loop -- doing so would clobber the scaled font and make the text-scale axis
    // a no-op. Capture the post-apply() base font as the reference point size.
    AppTheme::apply();
    const QFont baseFont = QApplication::font();
    const qreal basePt = (baseFont.pointSizeF() > 0)
        ? baseFont.pointSizeF()
        : qreal(AppTheme::fontDefault().pointSizeF());

    MainWindow window;
    window.show();
    settle(150);   // a grab needs the widget realized

    QJsonArray cases;
    bool allOk = true;

    for (const qreal scale : scales) {
        // Set the scaled app font AFTER apply() (apply() already ran once, before
        // the loop) so the scaled point size actually takes effect, then re-polish
        // every widget against it. This is the in-process analogue of OS text-
        // scaling and is the whole point of the scale axis.
        QFont f = baseFont;
        f.setPointSizeF(basePt * scale);
        QApplication::setFont(f);
        for (QWidget* top : QApplication::topLevelWidgets())
            QApplication::style()->polish(top);   // refresh derived metrics

        for (const QSize& s : sizes) {
            const QString label = QStringLiteral("%1x%2_x%3")
                .arg(s.width()).arg(s.height())
                .arg(QString::number(scale, 'g', 3));

            window.resize(s);
            settle(120);                       // debounce + re-flow
            QCoreApplication::processEvents();

            const QString pngName = label + QStringLiteral(".png");
            const QString pngPath = outDir + "/" + pngName;

            const QPixmap shot = window.grab();
            const bool grabbed = !shot.isNull() && shot.save(pngPath, "PNG");

            // Closed-loop no-clip pass/fail for this case (the full-window
            // fits-or-scrolls guarantee, verified in-process). A case is `pass`
            // only if every visible region fits or scrolls AND the grab wrote.
            QStringList failures;
            const bool regionsOk = regionsFitOrScroll(window, failures);
            const bool pass = grabbed && regionsOk;
            if (!pass) allOk = false;

            QJsonArray failArr;
            for (const QString& f : failures)
                failArr.append(f);

            // VeryNarrow dock-collapse evidence (Task 11): below 760px both side
            // docks should be hidden. Recorded per-case by real object name so a
            // human (or a JSON assertion) can confirm the auto-collapse fired
            // without instantiating MainWindow inside a headless Qt Test.
            const QDockWidget* navDock   = window.findChild<QDockWidget*>(
                QStringLiteral("navigatorDock"));
            const QDockWidget* notesDock = window.findChild<QDockWidget*>(
                QStringLiteral("notesDock"));

            QJsonObject o;
            o["label"]      = label;
            o["width"]      = s.width();
            o["height"]     = s.height();
            o["scale"]      = scale;
            o["png"]        = pngName;
            o["grab_w"]     = shot.width();
            o["grab_h"]     = shot.height();
            o["grabbed"]    = grabbed;
            o["regions_ok"] = regionsOk;
            o["pass"]       = pass;          // closed-loop per-case verdict
            o["failures"]   = failArr;       // clipped regions, if any
            o["nav_visible"]   = navDock   ? navDock->isVisible()   : false;
            o["notes_visible"] = notesDock ? notesDock->isVisible() : false;
            cases.append(o);

            QTextStream(stderr) << "ui-stress: " << label
                << (pass ? " PASS" : " FAIL")
                << (failures.isEmpty() ? QString()
                                       : QStringLiteral(" [%1]").arg(failures.join(QStringLiteral("; "))))
                << Qt::endl;
        }
    }

    // Restore the un-scaled base font (apply() already ran before the loop;
    // re-running it here just re-asserts fontDefault(), which equals baseFont).
    QApplication::setFont(baseFont);

    QJsonObject root;
    root["version"]    = QApplication::applicationVersion();
    root["out_dir"]    = QDir(outDir).absolutePath();
    root["all_ok"]     = allOk;   // true only if EVERY case grabbed AND no region clipped
    root["case_count"] = cases.size();
    root["cases"]      = cases;

    const QByteArray json = QJsonDocument(root).toJson(QJsonDocument::Indented);
    QFile idx(outDir + "/index.json");
    if (idx.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        idx.write(json);
        idx.close();
    } else {
        allOk = false;
    }

    QTextStream(stdout) << "ui-stress: wrote " << cases.size()
        << " cases + index.json to " << QDir(outDir).absolutePath()
        << Qt::endl;

    return allOk ? 0 : 1;
}

} // namespace DVE
```

(`AppTheme::fontDefault()` and `AppTheme::apply()` already exist. `window.grab()` requires `show()` — done; we deliberately do not `raise()`/`activateWindow()`. The header path `#include "../MainWindow.h"` matches the `src/utils/`→`src/` relative layout used by `SelfTest.cpp`. The per-region no-clip check consumes `MainWindow::scrollHostFor()` / `regionWidget()` from Task 16, so this task lands after Task 16.)

- [ ] **Step 2: Build, run the real harness, inspect output.**

```bash
python tools/decrypt_via_copy.py --apply
cd build && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && mingw32-make -j8' < /dev/null
./release/DataViewer.exe --ui-stress; echo "exit=$?"
ls "$TEMP/dve_ui_stress" 2>/dev/null || ls "$LOCALAPPDATA/Temp/dve_ui_stress"
```

Expected: `exit=0`; 28 PNGs (7 sizes × 4 scales) + `index.json` with `"all_ok": true`, `"case_count": 28`, and every case carrying `"pass": true` (the closed-loop per-region no-clip verdict) plus its `"nav_visible"`/`"notes_visible"` dock state. If any case has `"pass": false`, its `"failures"` array names the clipped region + need/viewport — that is a real responsive bug to fix in the offending component, then re-run. Also open the 480×360/×2.0 and 1600×500/×1.5 PNGs and confirm by eye: scrollbars where content overflows, "View Raw Data" ≤2 lines, no text spilling into the Navigator. The `--ui-stress` JSON pass/fail is the closed-loop gate the agent runs by launching `DataViewer.exe --ui-stress`; the PNGs are the human-eyeball complement. (At very-narrow widths — the 480×360 and 600×1200 cases — confirm `"nav_visible": false` and `"notes_visible": false`, proving Task 11's auto-collapse fired.)

- [ ] **Step 3: Verify `--ui-stress-out`.**

```bash
./build/release/DataViewer.exe --ui-stress --ui-stress-out "$TEMP/dve_stress_custom"; echo "exit=$?"
ls "$TEMP/dve_stress_custom"
```

Expected: PNGs + `index.json` in the custom dir.

- [ ] **Step 4: Confirm `-Werror -Wall -Wextra` clean.** Scan the Step 2 make output for warnings (e.g. an unused include — remove it; fix the underlying code, don't downgrade the flag).

- [ ] **Step 5: Commit.**

```bash
git add src/utils/UiStress.cpp
git commit -m "feat(verify): --ui-stress screenshot matrix harness (size x text-scale -> PNG + index.json)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 20: Document `--ui-stress` in the deployment README

**Files:** `tests/deployment/README.md` (append after the "What's not tested" block, before "## Adding a new check").

- [ ] **Step 1: Add the section** (Edit is fine for a `.md`; fall back to Python rewrite if the read returns ciphertext):

```markdown
## `--ui-stress` — responsive UI screenshot matrix

A closed-loop visual harness for the v2.7.0 responsive-UI work. It opens a
real `MainWindow`, cycles it through a matrix of window sizes × text-scale
factors, and saves one PNG per case plus an `index.json`. Use it to eyeball
every aspect-ratio / DPI case at a glance without owning every monitor.

```powershell
# default output dir: %TEMP%\dve_ui_stress
& "C:\Path\To\DataViewer.exe" --ui-stress

# custom output dir
& "C:\Path\To\DataViewer.exe" --ui-stress --ui-stress-out "$env:TEMP\dve_stress_custom"
```

- **Sizes:** 1920×1080, 1280×800, 960×540 (corner-snap quarter), 800×600,
  480×360 (the window floor), 600×1200 (tall-narrow), 1600×500 (wide-short).
- **Text scales:** 1.0, 1.25, 1.5, 2.0 — applied by multiplying the base
  application font point size (the in-process analogue of OS text-scaling).
- **Output:** 28 PNGs (7 sizes × 4 scales) named `<w>x<h>_x<scale>.png`,
  plus `index.json` listing every case with its size, scale, grab dimensions,
  a closed-loop `pass`/`fail` no-clip verdict (with a `failures` array naming
  any clipped region + its need/viewport), and the `nav_visible`/`notes_visible`
  side-dock state. Exit code `0` only if every case both grabbed AND passed the
  no-clip check.

This is a **closed-loop** harness, not just a screenshot dumper: for each
case it checks every wrapped region (central / navigator / notes / ribbon
group row) and verifies the content fits its ScrollHost viewport OR the
host's scrollbar is active in the overflow direction — so a clipped region
fails the case programmatically. The full-window "nothing clipped without a
scrollbar" guarantee is verified HERE (inside the real `DataViewer.exe`,
where a fully-constructed `MainWindow` with live DB/python context exists)
rather than in a headless Qt Test.

No GUI interaction: the window is shown (required for a valid `grab()`) but
never raised or focused. What to look for in the PNGs: scrollbars appear
wherever content overflows (never clipped-without-scrollbar), the ribbon
"View Raw Data" label stays ≤2 lines, and no text spills into the Navigator
at any scale. The 1280×800 / ×1.0 PNG is the standard-scale regression
baseline — it must look identical to today's UI. At extreme aspect ratios
(600×1200, 1600×500) the guarantee is "scrolls, not reflows": content is
reachable via scroll; splitter-orientation reflow is out of scope for v2.7.0.

The light Qt Tests (`tst_scrollhost`, `tst_ribbonlayout`, `tst_responsivelayout`,
`tst_sizingsweep`) cover the unit-level invariants (scroll behaviour, ≤2-line
ribbon wrap, breakpoints, dialog floors) without a full `MainWindow`;
`--ui-stress` is the full-window closed-loop complement plus a human eyeball.
```

- [ ] **Step 2: Commit.**

```bash
git add tests/deployment/README.md
git commit -m "docs(deploy): document --ui-stress screenshot matrix harness

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Verification & Done Criteria

Run all of the following from the repo root after the final task; every item must hold before the sprint wraps.

- [ ] **Full `-Werror -Wall -Wextra` build clean.**

```bash
python tools/decrypt_via_copy.py --apply
cd build && MSYS_NO_PATHCONV=1 cmd /c 'set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH% && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8' < /dev/null
```

Expected: `DataViewer.exe` produced, zero warnings.

- [ ] **`tests/run-tests.ps1` green.**

```bash
MSYS_NO_PATHCONV=1 cmd /c 'powershell -NoProfile -ExecutionPolicy Bypass -File tests\run-tests.ps1' < /dev/null
```

Expected: all suites pass — including the new `tst_scrollhost`, `tst_ribbonlayout`, `tst_sizingsweep`, and the stabilized `tst_responsivelayout` — **modulo the known pre-existing-flaky `tst_excelreader` / `tst_dataprocessor`** (bundled-Python file-load alternation, not a regression of this work). The Qt Tests stay light: each links only the few sources it exercises (mirroring the existing `tst_responsivelayout.pro`), and none instantiates `MainWindow` (whose ctor opens a DB connection + spawns Python, which is heavy/flaky headless).

- [ ] **The light Qt Tests pin their unit-level invariants.** `tst_scrollhost` proves the wrapper scrolls-not-clips in each direction; `tst_ribbonlayout` pins the ≤2-line "View Raw Data" wrap + grow-not-clip button sizing (linking only RibbonWidget + AppTheme + ScrollHost); `tst_responsivelayout` covers the `VeryNarrow` breakpoint (Task 10) and the singleton-reset flakiness fix (Task 17, plus a 10× repeat run); `tst_sizingsweep` pins the lowered dialog floor + growable controls. None of these drives a full `MainWindow`.

- [ ] **The full-window "no region clipped without a scrollbar" guarantee is verified by `--ui-stress`, not a Qt Test.** Because the guarantee needs a fully-constructed `MainWindow` with live DB/python context, it is checked in-process inside the real `DataViewer.exe`:

```bash
./build/release/DataViewer.exe --ui-stress; echo "exit=$?"
```

Expected: `exit=0`, 28 PNGs + `index.json` (`"all_ok": true`, `"case_count": 28`) in `%TEMP%\dve_ui_stress`, with **every** case carrying `"pass": true` — the closed-loop per-region verdict (each wrapped region central/navigator/notes/ribbonGroups fits its ScrollHost viewport OR its scrollbar is active in the overflow direction). A `"pass": false` case lists the clipped region + need/viewport in its `"failures"` array — a real responsive bug to fix in that component, then re-run. The 480×360 and 600×1200 (very-narrow) cases must also show `"nav_visible": false` / `"notes_visible": false`, confirming Task 11's dock auto-collapse fired. The agent runs this as the closed-loop pass/fail gate by launching `DataViewer.exe --ui-stress` and reading `index.json`; the PNGs are the human-eyeball complement (scrollbars where content overflows, "View Raw Data" ≤2 lines, no Navigator spill — spot-check the floor 480×360/×2.0 and extreme-aspect 1600×500/×1.5, 600×1200/×1.0).

- [ ] **Extreme aspect ratios scroll, they do not reflow (out of scope for v2.7.0).** At 600×1200 and 1600×500 the guarantee is "everything is *reachable* via scroll" — the per-region ScrollHosts scroll the content into view. The spec's optional splitter-orientation reflow hook (rotating the TPM vertical split to horizontal at landscape-extreme ratios, or vice versa) is explicitly **deferred**; `--ui-stress` asserts fits-or-scrolls at those ratios, not a re-layout. `ResponsiveLayout` tracks width only (Task 10), so no aspect-ratio breakpoint/signal exists in this sprint.

- [ ] **No visual regression at standard scale (Goal 5).** The 1280×800/×1.0 `--ui-stress` PNG (and an eyeball at 100% scale on a ≥1100px window) must look identical to v2.6.0 — large buttons 80×76, group titles uncut, no spurious scrollbars. Every fixed→min conversion kept the old value as the *minimum*, so the standard-size look is the floor, not a change.

- [ ] **Wrap to deployable v2.7.0** only after the owner's smoke test passes (do not bump VERSION or touch Synology autonomously — the owner is the prod-ship checkpoint).

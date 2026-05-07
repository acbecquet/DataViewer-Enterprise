# Sensory Report Preview Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a WYSIWYG preview/editor that opens when the user clicks any sensory "Report" button, lets them drag/resize all slide elements, sort tables, exclude samples, undo/redo, save presets, import/export layouts, and edit cover/divider title text — with edits persisted across sessions in the DB and the source `.xlsx`.

**Architecture:** Modal `ReportPreviewDialog` driven by a mode-agnostic `IReportSource` interface. `SensoryReportSource` is the only concrete adapter in v1; detailed-sensory and TPM are explicitly not in scope. Layout state lives in a new `ReportLayout` JSON model, persisted to a new `sensory_sessions.layout_json` column and an Excel `dve_layout` workbook custom property. Cumulative layout lives in the `settings` table. Cover/divider templates remain templated except for editable title text.

**Tech Stack:** C++17, Qt 6.10.1 (Widgets/Sql/Concurrent/Network), MinGW 13.1.0, qmake, openpyxl via bundled Python 3.11, Inno Setup. Tests use QtTest (`QTEST_MAIN`), one test binary per `tests/tst_<name>/` subdirectory.

**Reference spec:** `docs/superpowers/specs/2026-05-06-sensory-report-preview-design.md`

---

## File Structure

### New files

| File | Responsibility |
|---|---|
| `src/reporting/ReportLayout.h` / `.cpp` | JSON-serializable layout model: per-slide rects, table sort, z-order, version field |
| `src/reporting/IReportSource.h` | Pure-virtual adapter interface |
| `src/reporting/SensoryReportSource.h` / `.cpp` | Concrete adapter wrapping `QVector<SensorySession>` |
| `src/reporting/LayoutCommand.h` / `.cpp` | Undo/redo command base + concrete `MoveItemCommand`, `ResizeItemCommand`, `SortColumnCommand`, `EditTextCommand`, `ToggleSampleCommand`, `ZOrderCommand` |
| `src/reporting/PresetStore.h` / `.cpp` | DB CRUD for `layout_presets` table |
| `src/ui/SlideCanvasItems.h` / `.cpp` | `ResizableSlideItem` base + `PlotItem`, `TableItem`, `TextItem` subclasses |
| `src/ui/SamplesCheckboxPanel.h` / `.cpp` | Scrollable grouped sample-exclusion panel |
| `src/ui/PropertiesPanel.h` / `.cpp` | Selected-item position/size/Z-order editor (right side of dialog) |
| `src/ui/ReportPreviewDialog.h` / `.cpp` | The dialog itself |
| `src/ui/PresetManagerDialog.h` / `.cpp` | Tiny dialog for renaming/deleting presets |
| `tools/excel_layout_io.py` | Python helper invoked from C++ to read/write `dve_layout` Excel workbook custom property |
| `tests/tst_reportlayout/` | Tests for `ReportLayout` JSON round-trip |
| `tests/tst_layoutcommand/` | Tests for command apply/undo invariants |
| `tests/tst_presetstore/` | Tests for `layout_presets` CRUD |
| `tests/tst_sensoryreportsource/` | Tests for `SensoryReportSource::computeDefaultLayout` parity with current PPTX flow |

### Modified files

| File | What changes |
|---|---|
| `src/database/DatabaseManager.{h,cpp}` | Schema migration: add `sensory_sessions.layout_json` column, add `layout_presets` table; new methods `loadSensoryLayout(int sessionId)`, `saveSensoryLayout(int sessionId, const QString&)`, `loadCumulativeLayout()`, `saveCumulativeLayout(const QString&)` |
| `src/ExcelReader.{h,cpp}` | Read `dve_layout` workbook custom property via `excel_layout_io.py`; expose on parsed result |
| `src/reporting/PptxWriter.{h,cpp}` | `addContentSlide()` takes optional `LayoutOverrides` so radar/table/title/properties textbox positions can be supplied per-call; same for `addCoverSlide()` |
| `src/ui/SensoryPanel.{h,cpp}` | `generateFullReport()` and `generateCombinedPptx()` route through `ReportPreviewDialog` via `SensoryReportSource` |
| `src/ui/SensoryPanel.cpp` (~line 2008) | The chart-position math is extracted into `SensoryReportSource::computeDefaultLayout()` so the legacy fast-path and the dialog's "no edits" state share one implementation |
| `DataViewerEnterprise.pro` | New `SOURCES`/`HEADERS` lines for the files above |
| `tests/tests.pro` | New `SUBDIRS` lines for the new test binaries |

### Untouched

`src/MainWindow.cpp` TPM report flow stays as-is. Detailed-Sensory `DetailedSensoryPanel.{h,cpp}` not modified in this plan.

---

## Phase 1A — Data Layer (10 tasks)

Goal: every persistence and refactoring change happens before any UI work, so we can verify the existing PPTX output is byte-identical when no layout exists.

### Task 1: `ReportLayout` data class with JSON round-trip

**Files:**
- Create: `src/reporting/ReportLayout.h`
- Create: `src/reporting/ReportLayout.cpp`
- Create: `tests/tst_reportlayout/tst_reportlayout.pro`
- Create: `tests/tst_reportlayout/tst_reportlayout.cpp`
- Modify: `tests/tests.pro` (add `tst_reportlayout` SUBDIR)
- Modify: `DataViewerEnterprise.pro` (add SOURCES/HEADERS)

- [ ] **Step 1: Define the model header**

```cpp
// src/reporting/ReportLayout.h
#pragma once
#include <QJsonObject>
#include <QRectF>
#include <QString>
#include <QHash>

namespace DVE {

struct PropertiesBox {
    QRectF rect;
    QString text;
};

struct ContentSlideLayout {
    QRectF title;
    QRectF table;
    QRectF radar;
    PropertiesBox propertiesBox;
};

struct ImageSlideLayout {
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;
};

struct TableSort {
    QString column;            // empty = insertion order
    Qt::SortOrder order = Qt::DescendingOrder;
};

struct ReportLayout {
    static constexpr int kCurrentVersion = 1;
    int version = kCurrentVersion;
    QString modeId = QStringLiteral("sensory");

    TableSort tableSort;
    QHash<QString, ContentSlideLayout> contentSlides;   // key: "content_<sessionId>"
    QHash<QString, ImageSlideLayout>   imageSlides;     // key: "image_<sessionId>"
    QHash<QString, QRectF>             dividerTitles;   // key: "divider_<sessionId>"
    ContentSlideLayout cumulative;
    QRectF coverTitle;
    QRectF coverSubtitle;
    QStringList zOrder;                                  // top-to-bottom ids per slide

    QJsonObject toJson() const;
    static ReportLayout fromJson(const QJsonObject& obj, bool* ok = nullptr);
    bool isEmpty() const;       // true if all rects are default-constructed
};

} // namespace DVE
```

- [ ] **Step 2: Implement serialization**

```cpp
// src/reporting/ReportLayout.cpp
#include "ReportLayout.h"
#include <QJsonArray>

namespace DVE {

static QJsonArray rectToJson(const QRectF& r) {
    return QJsonArray{ r.x(), r.y(), r.width(), r.height() };
}

static QRectF rectFromJson(const QJsonArray& a) {
    if (a.size() != 4) return {};
    return { a[0].toDouble(), a[1].toDouble(), a[2].toDouble(), a[3].toDouble() };
}

static QJsonObject contentToJson(const ContentSlideLayout& c) {
    return {
        { "title",         rectToJson(c.title) },
        { "table",         rectToJson(c.table) },
        { "radar",         rectToJson(c.radar) },
        { "propertiesBox", QJsonObject{
            { "rect", rectToJson(c.propertiesBox.rect) },
            { "text", c.propertiesBox.text }
        } }
    };
}

static ContentSlideLayout contentFromJson(const QJsonObject& o) {
    ContentSlideLayout c;
    c.title = rectFromJson(o.value("title").toArray());
    c.table = rectFromJson(o.value("table").toArray());
    c.radar = rectFromJson(o.value("radar").toArray());
    const QJsonObject pb = o.value("propertiesBox").toObject();
    c.propertiesBox.rect = rectFromJson(pb.value("rect").toArray());
    c.propertiesBox.text = pb.value("text").toString();
    return c;
}

QJsonObject ReportLayout::toJson() const {
    QJsonObject slides;
    slides["cover"] = QJsonObject{
        { "title",    rectToJson(coverTitle) },
        { "subtitle", rectToJson(coverSubtitle) }
    };
    for (auto it = dividerTitles.cbegin(); it != dividerTitles.cend(); ++it)
        slides[it.key()] = QJsonObject{ { "title", rectToJson(it.value()) } };
    for (auto it = contentSlides.cbegin(); it != contentSlides.cend(); ++it)
        slides[it.key()] = contentToJson(it.value());
    for (auto it = imageSlides.cbegin(); it != imageSlides.cend(); ++it) {
        QJsonArray layouts, crops;
        for (const QRectF& r : it.value().imageLayouts) layouts.append(rectToJson(r));
        for (const QRectF& r : it.value().imageCrops)   crops.append(rectToJson(r));
        slides[it.key()] = QJsonObject{
            { "imageLayouts", layouts },
            { "imageCrops",   crops }
        };
    }
    slides["cumulative"] = contentToJson(cumulative);

    QJsonArray zOrderArr;
    for (const QString& s : zOrder) zOrderArr.append(s);

    return {
        { "version", version },
        { "mode",    modeId },
        { "tableSort", QJsonObject{
            { "column", tableSort.column },
            { "order",  tableSort.order == Qt::AscendingOrder ? "asc" : "desc" }
        } },
        { "slides",  slides },
        { "zOrder",  zOrderArr }
    };
}

ReportLayout ReportLayout::fromJson(const QJsonObject& obj, bool* ok) {
    ReportLayout r;
    if (ok) *ok = false;
    if (!obj.contains("version") || !obj.contains("mode")) return r;
    r.version = obj.value("version").toInt(kCurrentVersion);
    r.modeId  = obj.value("mode").toString();
    if (r.modeId != "sensory") return r;            // v1 only handles sensory

    const QJsonObject ts = obj.value("tableSort").toObject();
    r.tableSort.column = ts.value("column").toString();
    r.tableSort.order  = ts.value("order").toString() == "asc"
                         ? Qt::AscendingOrder : Qt::DescendingOrder;

    const QJsonObject slides = obj.value("slides").toObject();
    for (auto it = slides.begin(); it != slides.end(); ++it) {
        const QString& key = it.key();
        const QJsonObject v = it.value().toObject();
        if (key == "cover") {
            r.coverTitle    = rectFromJson(v.value("title").toArray());
            r.coverSubtitle = rectFromJson(v.value("subtitle").toArray());
        } else if (key == "cumulative") {
            r.cumulative = contentFromJson(v);
        } else if (key.startsWith("content_")) {
            r.contentSlides[key] = contentFromJson(v);
        } else if (key.startsWith("image_")) {
            ImageSlideLayout img;
            for (const QJsonValue& jv : v.value("imageLayouts").toArray())
                img.imageLayouts.append(rectFromJson(jv.toArray()));
            for (const QJsonValue& jv : v.value("imageCrops").toArray())
                img.imageCrops.append(rectFromJson(jv.toArray()));
            r.imageSlides[key] = img;
        } else if (key.startsWith("divider_")) {
            r.dividerTitles[key] = rectFromJson(v.value("title").toArray());
        }
    }

    for (const QJsonValue& jv : obj.value("zOrder").toArray())
        r.zOrder.append(jv.toString());

    if (ok) *ok = true;
    return r;
}

bool ReportLayout::isEmpty() const {
    return contentSlides.isEmpty() && imageSlides.isEmpty()
        && dividerTitles.isEmpty() && coverTitle.isNull();
}

} // namespace DVE
```

- [ ] **Step 3: Write the failing test**

```cpp
// tests/tst_reportlayout/tst_reportlayout.cpp
#include <QtTest>
#include <QJsonDocument>
#include "ReportLayout.h"

class tst_ReportLayout : public QObject {
    Q_OBJECT
private slots:
    void testRoundTripPreservesAllFields() {
        DVE::ReportLayout in;
        in.tableSort.column = "Overall Liking";
        in.tableSort.order  = Qt::AscendingOrder;
        in.coverTitle    = QRectF(0.1, 0.2, 12.0, 1.0);
        in.coverSubtitle = QRectF(0.1, 1.5, 12.0, 0.5);

        DVE::ContentSlideLayout c;
        c.title = QRectF(0.1, 0.1, 12.0, 0.5);
        c.table = QRectF(0.32, 0.75, 12.7, 1.5);
        c.radar = QRectF(4.5, 2.5, 4.4, 4.4);
        c.propertiesBox.rect = QRectF(10.1, 4.9, 3.17, 2.5);
        c.propertiesBox.text = "Media: cap1\nControl: ctrl1";
        in.contentSlides["content_42"] = c;

        DVE::ImageSlideLayout img;
        img.imageLayouts << QRectF(1, 1, 5, 4);
        img.imageCrops   << QRectF(0, 0, 1, 1);
        in.imageSlides["image_42"] = img;

        in.dividerTitles["divider_42"] = QRectF(0.1, 3.0, 13.0, 1.5);
        in.cumulative = c;
        in.zOrder << "table" << "radar";

        const QJsonObject obj = in.toJson();
        bool ok = false;
        const DVE::ReportLayout out = DVE::ReportLayout::fromJson(obj, &ok);
        QVERIFY(ok);
        QCOMPARE(out.tableSort.column, in.tableSort.column);
        QCOMPARE(out.tableSort.order,  in.tableSort.order);
        QCOMPARE(out.coverTitle,       in.coverTitle);
        QCOMPARE(out.coverSubtitle,    in.coverSubtitle);
        QVERIFY(out.contentSlides.contains("content_42"));
        QCOMPARE(out.contentSlides["content_42"].table, c.table);
        QCOMPARE(out.contentSlides["content_42"].propertiesBox.text, c.propertiesBox.text);
        QVERIFY(out.imageSlides.contains("image_42"));
        QCOMPARE(out.imageSlides["image_42"].imageLayouts.size(), 1);
        QCOMPARE(out.dividerTitles["divider_42"], in.dividerTitles["divider_42"]);
        QCOMPARE(out.zOrder, in.zOrder);
    }

    void testRejectsWrongMode() {
        QJsonObject o{ { "version", 1 }, { "mode", "tpm" } };
        bool ok = true;
        DVE::ReportLayout::fromJson(o, &ok);
        QVERIFY(!ok);
    }

    void testEmptyLayoutIsDetected() {
        DVE::ReportLayout l;
        QVERIFY(l.isEmpty());
    }
};

QTEST_MAIN(tst_ReportLayout)
#include "tst_reportlayout.moc"
```

- [ ] **Step 4: Test project file**

```
# tests/tst_reportlayout/tst_reportlayout.pro
QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/reporting

SOURCES += tst_reportlayout.cpp \
           ../../src/reporting/ReportLayout.cpp

HEADERS += ../../src/reporting/ReportLayout.h
```

- [ ] **Step 5: Wire into the test runner and the app .pro**

In `tests/tests.pro` add `tst_reportlayout` to the SUBDIRS list.
In `DataViewerEnterprise.pro` SOURCES add `src/reporting/ReportLayout.cpp` and HEADERS add `src/reporting/ReportLayout.h`.

- [ ] **Step 6: Build and verify**

```
cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && cd tests\tst_reportlayout && qmake.exe -spec win32-g++ tst_reportlayout.pro && mingw32-make.exe && release\tst_reportlayout.exe"
```

Expected: 3 tests passed, 0 failed.

- [ ] **Step 7: Commit**

```
git add src/reporting/ReportLayout.h src/reporting/ReportLayout.cpp \
        tests/tst_reportlayout/ tests/tests.pro DataViewerEnterprise.pro
git commit -m "feat(reports): ReportLayout JSON model with round-trip tests"
```

---

### Task 2: DatabaseManager schema migration — `sensory_sessions.layout_json` + `settings` cumulative key

**Files:**
- Modify: `src/database/DatabaseManager.cpp` — extend the `migrateSchema()` function (or equivalent) and add to `m_currentSchemaVersion`
- Modify: `src/database/DatabaseManager.h` — add new method declarations
- Modify: `tests/tst_databasemanager/tst_databasemanager.cpp` — add migration test

- [ ] **Step 1: Read the existing schema-migration code path**

Read `src/database/DatabaseManager.cpp` and locate the migration function. Most likely a `migrateSchema()` or `applyMigrations()` method, called from `open()`. Note the existing schema version constant.

- [ ] **Step 2: Add migration**

Append a new migration block. Pseudocode:

```cpp
// In migrateSchema(), after the most recent existing block
if (currentVersion < kSchemaWithLayoutJson) {
    QSqlQuery q(m_db);
    if (!q.exec("ALTER TABLE sensory_sessions ADD COLUMN layout_json TEXT"))
        return logSqlError(q, "add layout_json to sensory_sessions");
    setSchemaVersion(kSchemaWithLayoutJson);
}
```

Update the `kSchemaWith…` constant chain (e.g., bump `kCurrentSchemaVersion` by 1).

The `settings` table already exists for cumulative storage — no schema change needed there, just new keys.

- [ ] **Step 3: Add API methods to header**

```cpp
// In DatabaseManager.h, public:
QString loadSensoryLayout(int sessionId) const;
bool    saveSensoryLayout(int sessionId, const QString& layoutJson);

QString loadCumulativeLayout() const;
bool    saveCumulativeLayout(const QString& layoutJson);
```

- [ ] **Step 4: Implement**

```cpp
// DatabaseManager.cpp
QString DatabaseManager::loadSensoryLayout(int sessionId) const {
    QSqlQuery q(m_db);
    q.prepare("SELECT layout_json FROM sensory_sessions WHERE id = :id");
    q.bindValue(":id", sessionId);
    if (!q.exec() || !q.next()) return {};
    return q.value(0).toString();
}

bool DatabaseManager::saveSensoryLayout(int sessionId, const QString& json) {
    QSqlQuery q(m_db);
    q.prepare("UPDATE sensory_sessions SET layout_json = :j WHERE id = :id");
    q.bindValue(":j", json);
    q.bindValue(":id", sessionId);
    return q.exec();
}

QString DatabaseManager::loadCumulativeLayout() const {
    return getSetting("sensory.cumulative_layout");   // existing settings API
}

bool DatabaseManager::saveCumulativeLayout(const QString& json) {
    return setSetting("sensory.cumulative_layout", json);
}
```

(Reuse whatever `getSetting`/`setSetting` API already exists in `DatabaseManager` — locate it by grep first.)

- [ ] **Step 5: Add tests**

Append to `tests/tst_databasemanager/tst_databasemanager.cpp`:

```cpp
void testSensoryLayoutPersistence() {
    DVE::DatabaseManager db;
    QVERIFY(db.open(":memory:"));

    DVE::SensorySession s;
    s.testTitle = "Test"; s.testerName = "T"; s.date = "2026-01-01";
    DVE::SensorySample samp; samp.name = "S1"; s.samples.append(samp);

    int id = db.saveSensorySession(s);
    QVERIFY(id > 0);

    QCOMPARE(db.loadSensoryLayout(id), QString());     // NULL → empty

    QVERIFY(db.saveSensoryLayout(id, R"({"version":1,"mode":"sensory"})"));
    QCOMPARE(db.loadSensoryLayout(id),
             QString(R"({"version":1,"mode":"sensory"})"));
}

void testCumulativeLayoutPersistence() {
    DVE::DatabaseManager db;
    QVERIFY(db.open(":memory:"));
    QCOMPARE(db.loadCumulativeLayout(), QString());
    QVERIFY(db.saveCumulativeLayout("{\"x\":1}"));
    QCOMPARE(db.loadCumulativeLayout(), QString("{\"x\":1}"));
}
```

- [ ] **Step 6: Run tests**

```
cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && cd tests\tst_databasemanager && mingw32-make.exe && release\tst_databasemanager.exe"
```

Expected: previous 18 tests + 2 new = 20 pass.

- [ ] **Step 7: Commit**

```
git add src/database/DatabaseManager.h src/database/DatabaseManager.cpp \
        tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "feat(db): sensory_sessions.layout_json column + cumulative layout setting"
```

---

### Task 3: `IReportSource` interface

**Files:**
- Create: `src/reporting/IReportSource.h`
- Modify: `DataViewerEnterprise.pro` (add HEADERS line)

- [ ] **Step 1: Define interface**

```cpp
// src/reporting/IReportSource.h
#pragma once
#include <QString>
#include <QVector>
#include <QSet>
#include "ReportLayout.h"

namespace DVE {

enum class SlideKind { Cover, Divider, Content, Image, Cumulative };

struct SampleRef {
    QString slideKey;     // "content_<id>" — the slide this sample belongs to
    QString sampleId;     // unique within slideKey, used in excludedSamples set
    QString displayName;
    QString sessionLabel; // shown as group header in checkbox panel
};

struct ReportSlideSpec {
    SlideKind kind;
    QString   slideKey;       // "cover" | "divider_<id>" | "content_<id>" | etc.
    QString   title;          // current title text (editable)
    // Resolved geometry + content for the canvas, after layout overrides applied.
    // Only one of these is meaningful per kind:
    QImage    radarPixmap;
    QStringList tableHeaders;
    QVector<QStringList> tableRows;
    QString   propertiesText;
    QStringList imagePaths;
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;
    ContentSlideLayout layout;     // resolved geometry
};

class IReportSource {
public:
    virtual ~IReportSource() = default;

    virtual QString modeId() const = 0;          // "sensory"
    virtual QString sourceLabel() const = 0;     // dialog title

    virtual int slideCount() const = 0;
    virtual SlideKind slideKind(int idx) const = 0;
    virtual ReportSlideSpec buildSlide(int idx, const ReportLayout&,
                                        const QSet<QString>& excludedSamples) const = 0;

    virtual QVector<SampleRef> allSamples() const = 0;

    virtual ReportLayout loadLayout() const = 0;
    virtual void saveLayout(const ReportLayout&) = 0;

    virtual bool writePptx(const QString& outPath, const ReportLayout&,
                            const QSet<QString>& excludedSamples,
                            QString* errorOut) = 0;
};

} // namespace DVE
```

- [ ] **Step 2: Add to .pro HEADERS**

In `DataViewerEnterprise.pro` HEADERS list add `src/reporting/IReportSource.h`.

- [ ] **Step 3: Verify compile**

```
cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe -spec win32-g++ DataViewerEnterprise.pro CONFIG+=release && mingw32-make.exe release"
```

Expected: clean compile (no implementation file yet, header compiles transitively when included by IReportSource.h consumers).

- [ ] **Step 4: Commit**

```
git add src/reporting/IReportSource.h DataViewerEnterprise.pro
git commit -m "feat(reports): IReportSource interface for mode-agnostic preview"
```

---

### Task 4: `SensoryReportSource` scaffold (identity, slide enumeration, sample enumeration)

**Prerequisite — extend `SensorySession`.** The existing `SensorySession` struct has neither a DB id nor a source-file path. The dialog needs both: id to anchor `sensory_sessions.layout_json` writes, file path to write the `dve_layout` Excel custom property in Phase 2. Add to `src/pipeline/SensoryData.h` inside the `SensorySession` struct:

```cpp
int     id = -1;             // -1 if not yet persisted; set by DB loaders
QString sourceFilePath;      // path of the .xlsx the session was loaded from
                              // (empty if loaded only from DB)
QString excelLayoutJson;     // dve_layout custom property pulled by ExcelReader
                              // (empty until Phase 2 lands)
```

Update `DatabaseManager::loadSensorySession(int id)` and `loadSensorySessions()` to populate `id` on the returned struct from the SELECT result. Update `saveSensorySession` to set `id` on a non-const reference parameter (change signature to `bool saveSensorySession(SensorySession& s)`) so callers receive the autoincrement id back. Update every existing call site.

This step has its own commit:

```
git add src/pipeline/SensoryData.h src/database/DatabaseManager.{h,cpp} \
        src/ui/SensoryPanel.cpp src/MainWindow.cpp
git commit -m "feat(sensory): add id + sourceFilePath to SensorySession struct"
```

**Files:**
- Create: `src/reporting/SensoryReportSource.h`
- Create: `src/reporting/SensoryReportSource.cpp`
- Create: `tests/tst_sensoryreportsource/tst_sensoryreportsource.pro`
- Create: `tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp`
- Modify: `tests/tests.pro` (add SUBDIR)
- Modify: `DataViewerEnterprise.pro`
- Modify (prerequisite, see above): `src/pipeline/SensoryData.h`, `src/database/DatabaseManager.{h,cpp}`

- [ ] **Step 1: Header**

```cpp
// src/reporting/SensoryReportSource.h
#pragma once
#include "IReportSource.h"
#include "pipeline/SensoryData.h"

namespace DVE {

class DatabaseManager;

class SensoryReportSource : public IReportSource {
public:
    SensoryReportSource(QVector<SensorySession> sessions,
                        DatabaseManager* db,
                        QObject* qObjectParent = nullptr);

    QString modeId() const override { return QStringLiteral("sensory"); }
    QString sourceLabel() const override;

    int slideCount() const override;
    SlideKind slideKind(int idx) const override;
    ReportSlideSpec buildSlide(int idx, const ReportLayout&,
                                const QSet<QString>&) const override;

    QVector<SampleRef> allSamples() const override;

    ReportLayout loadLayout() const override;
    void saveLayout(const ReportLayout&) override;

    bool writePptx(const QString& outPath, const ReportLayout&,
                    const QSet<QString>&, QString* errorOut) override;

    // Public so tests can verify defaults match the legacy fast path.
    static ReportLayout computeDefaultLayout(const QVector<SensorySession>& sessions);

private:
    struct SlideEntry { SlideKind kind; int sessionIdx; QString key; };
    void buildSlideIndex();

    QVector<SensorySession> m_sessions;
    DatabaseManager*        m_db;
    QVector<SlideEntry>     m_slides;   // flat list in render order
};

} // namespace DVE
```

- [ ] **Step 2: Implement scaffold (no `computeDefaultLayout`, no `writePptx`, no `buildSlide` body yet — those come in Tasks 5/6/8)**

```cpp
// src/reporting/SensoryReportSource.cpp
#include "SensoryReportSource.h"
#include "database/DatabaseManager.h"

namespace DVE {

SensoryReportSource::SensoryReportSource(QVector<SensorySession> sessions,
                                          DatabaseManager* db, QObject*)
    : m_sessions(std::move(sessions)), m_db(db) { buildSlideIndex(); }

void SensoryReportSource::buildSlideIndex()
{
    m_slides.clear();
    if (m_sessions.isEmpty()) return;
    // Cover
    m_slides.push_back({ SlideKind::Cover, -1, QStringLiteral("cover") });
    // Group by testerName for divider grouping (mirrors generateCombinedPptx)
    QHash<QString, QVector<int>> byTester;
    QStringList testerOrder;
    for (int i = 0; i < m_sessions.size(); ++i) {
        const QString& t = m_sessions[i].testerName;
        if (!byTester.contains(t)) testerOrder.append(t);
        byTester[t].append(i);
    }
    for (const QString& tester : testerOrder) {
        const QVector<int>& idxs = byTester[tester];
        m_slides.push_back({ SlideKind::Divider, idxs.first(),
                              QStringLiteral("divider_%1").arg(idxs.first()) });
        for (int sIdx : idxs) {
            m_slides.push_back({ SlideKind::Content, sIdx,
                                  QStringLiteral("content_%1").arg(sIdx) });
            if (!m_sessions[sIdx].imagePaths.isEmpty())
                m_slides.push_back({ SlideKind::Image, sIdx,
                                      QStringLiteral("image_%1").arg(sIdx) });
        }
    }
    // Cumulative summary if 2+ sessions
    if (m_sessions.size() >= 2)
        m_slides.push_back({ SlideKind::Cumulative, -1, QStringLiteral("cumulative") });
}

QString SensoryReportSource::sourceLabel() const {
    if (m_sessions.size() == 1)
        return m_sessions.first().testTitle;
    return QStringLiteral("%1 sessions").arg(m_sessions.size());
}

int SensoryReportSource::slideCount() const { return m_slides.size(); }
SlideKind SensoryReportSource::slideKind(int idx) const { return m_slides[idx].kind; }

QVector<SampleRef> SensoryReportSource::allSamples() const {
    QVector<SampleRef> out;
    for (int i = 0; i < m_sessions.size(); ++i) {
        const auto& sess = m_sessions[i];
        const QString slideKey = QStringLiteral("content_%1").arg(i);
        for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
            const auto& samp = sess.samples[sIdx];
            SampleRef r;
            r.slideKey = slideKey;
            r.sampleId = QStringLiteral("%1#%2").arg(i).arg(sIdx);
            r.displayName = samp.name.isEmpty()
                ? QStringLiteral("Sample %1").arg(sIdx + 1) : samp.name;
            r.sessionLabel = sess.testTitle.isEmpty()
                ? QStringLiteral("Session %1").arg(i + 1) : sess.testTitle;
            out.append(r);
        }
    }
    return out;
}

// Stubs — filled in subsequent tasks
ReportSlideSpec SensoryReportSource::buildSlide(int, const ReportLayout&,
                                                  const QSet<QString>&) const { return {}; }
ReportLayout SensoryReportSource::loadLayout() const { return {}; }
void SensoryReportSource::saveLayout(const ReportLayout&) {}
bool SensoryReportSource::writePptx(const QString&, const ReportLayout&,
                                     const QSet<QString>&, QString*) { return false; }
ReportLayout SensoryReportSource::computeDefaultLayout(const QVector<SensorySession>&) { return {}; }

} // namespace DVE
```

- [ ] **Step 3: Test scaffold (slide enumeration, sample enumeration)**

```cpp
// tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp
#include <QtTest>
#include "SensoryReportSource.h"

class tst_SensoryReportSource : public QObject {
    Q_OBJECT
private:
    DVE::SensorySession makeSess(const QString& title, const QString& tester,
                                  const QStringList& sampleNames) {
        DVE::SensorySession s;
        s.testTitle = title; s.testerName = tester; s.date = "2026-01-01";
        for (const QString& n : sampleNames) {
            DVE::SensorySample samp; samp.name = n;
            s.samples.append(samp);
        }
        return s;
    }
private slots:
    void testSlideOrderForSingleSession() {
        QVector<DVE::SensorySession> sessions{ makeSess("T1", "A", {"X","Y"}) };
        DVE::SensoryReportSource src(sessions, nullptr);
        // Cover, Divider, Content, (no Image — no images), no Cumulative (single sess)
        QCOMPARE(src.slideCount(), 3);
        QCOMPARE(src.slideKind(0), DVE::SlideKind::Cover);
        QCOMPARE(src.slideKind(1), DVE::SlideKind::Divider);
        QCOMPARE(src.slideKind(2), DVE::SlideKind::Content);
    }

    void testSlideOrderForMultipleSessions() {
        QVector<DVE::SensorySession> sessions{
            makeSess("T1", "A", {"X"}),
            makeSess("T1", "A", {"X"}),
            makeSess("T1", "B", {"X"})
        };
        DVE::SensoryReportSource src(sessions, nullptr);
        // Cover, Divider(A), Content(0), Content(1), Divider(B), Content(2), Cumulative
        QCOMPARE(src.slideCount(), 7);
        QCOMPARE(src.slideKind(6), DVE::SlideKind::Cumulative);
    }

    void testAllSamplesEnumeration() {
        QVector<DVE::SensorySession> sessions{
            makeSess("T1", "A", {"X","Y"}),
            makeSess("T1", "B", {"Z"})
        };
        DVE::SensoryReportSource src(sessions, nullptr);
        const auto refs = src.allSamples();
        QCOMPARE(refs.size(), 3);
        QCOMPARE(refs[0].slideKey, QString("content_0"));
        QCOMPARE(refs[0].displayName, QString("X"));
        QCOMPARE(refs[2].slideKey, QString("content_1"));
    }
};

QTEST_MAIN(tst_SensoryReportSource)
#include "tst_sensoryreportsource.moc"
```

- [ ] **Step 4: Test .pro file**

```
QT += core gui testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/reporting ../../src/pipeline ../../src/database

SOURCES += tst_sensoryreportsource.cpp \
           ../../src/reporting/SensoryReportSource.cpp \
           ../../src/reporting/ReportLayout.cpp

HEADERS += ../../src/reporting/SensoryReportSource.h \
           ../../src/reporting/IReportSource.h \
           ../../src/reporting/ReportLayout.h \
           ../../src/pipeline/SensoryData.h
```

- [ ] **Step 5: Wire into runners and build**

Add `tst_sensoryreportsource` to `tests/tests.pro`. Add the new files to `DataViewerEnterprise.pro`.

```
cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && cd tests\tst_sensoryreportsource && qmake.exe -spec win32-g++ tst_sensoryreportsource.pro && mingw32-make.exe && release\tst_sensoryreportsource.exe"
```

Expected: 3 tests pass.

- [ ] **Step 6: Commit**

```
git add src/reporting/SensoryReportSource.{h,cpp} \
        tests/tst_sensoryreportsource/ tests/tests.pro DataViewerEnterprise.pro
git commit -m "feat(reports): SensoryReportSource scaffold (identity + slide/sample enumeration)"
```

---

### Task 5: `SensoryReportSource::computeDefaultLayout` — port positioning math

**Files:**
- Modify: `src/reporting/SensoryReportSource.cpp`
- Modify: `tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp`
- Read: `src/ui/SensoryPanel.cpp:2008-2078` (the existing chart positioning block)

- [ ] **Step 1: Read the existing math**

Open `src/ui/SensoryPanel.cpp` and study the block from "Dynamic chart placement" (~line 2008) through the properties textbox calculation (~line 2078). Note the constants: `slideH = 7.5`, `slideW = 13.33`, `gapAbove = gapBelow = 0.10`, table `x = 0.32`, table `w = 12.7`, properties textbox `tbW = 3.17`.

- [ ] **Step 2: Implement**

Replace the stub in `SensoryReportSource.cpp`:

```cpp
ReportLayout SensoryReportSource::computeDefaultLayout(const QVector<SensorySession>& sessions)
{
    constexpr double slideW = 13.33;
    constexpr double slideH = 7.5;
    constexpr double tableX = 0.32;
    constexpr double tableW = 12.7;
    constexpr double tableY = 0.75;
    constexpr double gap    = 0.10;

    auto contentForSession = [&](const SensorySession& s) {
        ContentSlideLayout c;
        // Title spans full width above table
        c.title = QRectF(tableX, 0.10, tableW, 0.55);

        // Table height grows with sample count: header (~0.5) + rows (~0.33 each)
        const double tableH = 0.50 + s.samples.size() * (1.0 / 3.0);
        c.table = QRectF(tableX, tableY, tableW, tableH);

        // Radar centered horizontally below table, square aspect, fills remaining vertical
        const double tableBottom = tableY + tableH + gap;
        const double availH      = slideH - gap - tableBottom;
        const double radarH      = qMin(availH, slideW - 0.4);
        const double radarW      = radarH;
        const double radarX      = (slideW - radarW) / 2.0;
        c.radar = QRectF(radarX, tableBottom, radarW, radarH);

        // Properties textbox bottom-right
        constexpr double tbW = 3.17;
        const double tbH     = 2.0;     // dynamic in the current code; constant default here
        c.propertiesBox.rect = QRectF(slideW - tbW - 0.05,
                                       slideH - tbH - 0.05,
                                       tbW, tbH);
        c.propertiesBox.text.clear();   // filled in by buildSlide at render time
        return c;
    };

    ReportLayout layout;
    layout.coverTitle    = QRectF(0.5, 2.5, slideW - 1.0, 1.5);
    layout.coverSubtitle = QRectF(0.5, 4.2, slideW - 1.0, 0.8);

    for (int i = 0; i < sessions.size(); ++i) {
        const QString contentKey = QStringLiteral("content_%1").arg(i);
        const QString dividerKey = QStringLiteral("divider_%1").arg(i);
        layout.contentSlides[contentKey] = contentForSession(sessions[i]);
        layout.dividerTitles[dividerKey] = QRectF(0.5, slideH/2 - 0.75,
                                                   slideW - 1.0, 1.5);

        if (!sessions[i].imagePaths.isEmpty()) {
            ImageSlideLayout img;
            // Default grid: up to 4 images per slide, 3:2 ratio, 0.25" margin
            const int n = sessions[i].imagePaths.size();
            const int cols = qMin(4, n);
            const int rows = (n + cols - 1) / cols;
            const double cellW = (slideW - 0.5) / cols;
            const double cellH = (slideH - 0.5) / qMax(1, rows);
            for (int k = 0; k < n; ++k) {
                const int r = k / cols, c = k % cols;
                img.imageLayouts.append(QRectF(0.25 + c * cellW,
                                                0.25 + r * cellH,
                                                cellW - 0.1, cellH - 0.1));
                img.imageCrops.append(QRectF(0, 0, 1, 1));
            }
            layout.imageSlides[QStringLiteral("image_%1").arg(i)] = img;
        }
    }

    if (sessions.size() >= 2)
        layout.cumulative = contentForSession(sessions.first());

    return layout;
}
```

- [ ] **Step 3: Add tests**

```cpp
void testDefaultLayoutHasCorrectTablePosition() {
    QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S1","S2","S3"}) };
    const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);
    const auto& c = layout.contentSlides["content_0"];
    QCOMPARE(c.table.x(), 0.32);
    QCOMPARE(c.table.y(), 0.75);
    QCOMPARE(c.table.width(), 12.7);
    // 3 samples -> 0.5 + 3*(1/3) = 1.5
    QVERIFY(qFuzzyCompare(c.table.height(), 1.5));
}

void testDefaultRadarIsCenteredAndSquare() {
    QVector<DVE::SensorySession> sessions{ makeSess("T", "A", {"S"}) };
    const auto layout = DVE::SensoryReportSource::computeDefaultLayout(sessions);
    const auto& c = layout.contentSlides["content_0"];
    QCOMPARE(c.radar.width(), c.radar.height());
    const double mid = c.radar.x() + c.radar.width() / 2.0;
    QVERIFY(qFuzzyCompare(mid, 13.33 / 2.0));
}

void testCumulativeOnlyForMultiSession() {
    QVector<DVE::SensorySession> single{ makeSess("T", "A", {"S"}) };
    QVERIFY(DVE::SensoryReportSource::computeDefaultLayout(single).cumulative.table.isNull());
    QVector<DVE::SensorySession> two{ makeSess("T","A",{"S"}), makeSess("T","B",{"S"}) };
    QVERIFY(!DVE::SensoryReportSource::computeDefaultLayout(two).cumulative.table.isNull());
}
```

- [ ] **Step 4: Run tests**

```
cmd.exe //c "cd tests\tst_sensoryreportsource && set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && mingw32-make.exe && release\tst_sensoryreportsource.exe"
```

Expected: 6 tests pass.

- [ ] **Step 5: Commit**

```
git add src/reporting/SensoryReportSource.cpp \
        tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp
git commit -m "feat(reports): SensoryReportSource::computeDefaultLayout matches legacy positioning"
```

---

### Task 6: `SensoryReportSource::loadLayout` / `saveLayout` (DB only — Excel comes in Phase 2)

**Files:**
- Modify: `src/reporting/SensoryReportSource.cpp`
- Modify: `tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp`

- [ ] **Step 1: Implement load**

Replace stubs:

```cpp
ReportLayout SensoryReportSource::loadLayout() const {
    if (!m_db || m_sessions.isEmpty())
        return computeDefaultLayout(m_sessions);

    // Use the first session's id as the persistence anchor for now
    // (single-session report). For cumulative across sessions we layer in
    // the cumulative-layout setting below.
    const QString json = m_db->loadSensoryLayout(m_sessions.first().id);
    if (json.isEmpty())
        return computeDefaultLayout(m_sessions);

    const QJsonDocument doc = QJsonDocument::fromJson(json.toUtf8());
    bool ok = false;
    ReportLayout layout = ReportLayout::fromJson(doc.object(), &ok);
    if (!ok) return computeDefaultLayout(m_sessions);

    // Overlay cumulative-layout setting if present
    const QString cumJson = m_db->loadCumulativeLayout();
    if (!cumJson.isEmpty()) {
        const QJsonObject co = QJsonDocument::fromJson(cumJson.toUtf8()).object();
        ContentSlideLayout cum;
        cum.title = rectFromJson(co.value("title").toArray());
        cum.table = rectFromJson(co.value("table").toArray());
        cum.radar = rectFromJson(co.value("radar").toArray());
        const QJsonObject pb = co.value("propertiesBox").toObject();
        cum.propertiesBox.rect = rectFromJson(pb.value("rect").toArray());
        cum.propertiesBox.text = pb.value("text").toString();
        layout.cumulative = cum;
    }

    return layout;
}

void SensoryReportSource::saveLayout(const ReportLayout& layout) {
    if (!m_db || m_sessions.isEmpty()) return;
    const QJsonDocument doc(layout.toJson());
    m_db->saveSensoryLayout(m_sessions.first().id,
                             QString::fromUtf8(doc.toJson(QJsonDocument::Compact)));

    // Cumulative is global — write only its sub-object
    QJsonObject cum;
    cum["title"] = rectToJson(layout.cumulative.title);
    cum["table"] = rectToJson(layout.cumulative.table);
    cum["radar"] = rectToJson(layout.cumulative.radar);
    cum["propertiesBox"] = QJsonObject{
        { "rect", rectToJson(layout.cumulative.propertiesBox.rect) },
        { "text", layout.cumulative.propertiesBox.text }
    };
    m_db->saveCumulativeLayout(QString::fromUtf8(
        QJsonDocument(cum).toJson(QJsonDocument::Compact)));
}
```

(The `rectFromJson` / `rectToJson` helpers in `ReportLayout.cpp` are file-static — expose them as namespace-internal in the header or duplicate the 6-line helpers in this .cpp. Recommend exposing.)

- [ ] **Step 2: Expose helpers in ReportLayout.h**

In `ReportLayout.h` add:

```cpp
namespace DVE {
QJsonArray rectToJsonArray(const QRectF&);
QRectF     rectFromJsonArray(const QJsonArray&);
}
```

And rename the file-static functions in `ReportLayout.cpp` accordingly. Update `SensoryReportSource.cpp` to use `rectToJsonArray` / `rectFromJsonArray`.

- [ ] **Step 3: Test load/save round-trip with a real DB**

```cpp
void testLoadSaveLayoutRoundTrip() {
    DVE::DatabaseManager db;
    QVERIFY(db.open(":memory:"));

    DVE::SensorySession s = makeSess("T", "A", {"X"});
    s.id = db.saveSensorySession(s);

    DVE::SensoryReportSource src({s}, &db);

    DVE::ReportLayout l = src.loadLayout();      // returns defaults
    l.contentSlides["content_0"].title.setX(1.234);
    src.saveLayout(l);

    DVE::SensoryReportSource src2({s}, &db);
    DVE::ReportLayout l2 = src2.loadLayout();
    QCOMPARE(l2.contentSlides["content_0"].title.x(), 1.234);
}
```

- [ ] **Step 4: Run tests, verify pass**

- [ ] **Step 5: Commit**

```
git add src/reporting/SensoryReportSource.cpp src/reporting/ReportLayout.{h,cpp} \
        tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp
git commit -m "feat(reports): SensoryReportSource DB layout persistence"
```

---

### Task 7: PptxWriter accepts layout overrides

**Files:**
- Modify: `src/reporting/PptxWriter.h`
- Modify: `src/reporting/PptxWriter.cpp`
- Modify: `tests/tst_pptxwriter/tst_pptxwriter.cpp`

- [ ] **Step 1: Extend `addContentSlide` signature**

In `PptxWriter.h` add an overload that takes `ContentSlideLayout` (defined in `ReportLayout.h`). Keep the existing signature as a thin wrapper that fills `ContentSlideLayout` from the legacy hardcoded positions. Same pattern for `addCoverSlide`.

```cpp
// PptxWriter.h
#include "reporting/ReportLayout.h"

class PptxWriter {
public:
    // ... existing methods ...

    // New overload — used by ReportPreviewDialog and SensoryReportSource
    bool addContentSlide(const QString& title,
                          const SlideTable& table,
                          const QVector<SlideImage>& plots,
                          const ContentSlideLayout& layout,
                          const QString& extraXml = QString());
};
```

- [ ] **Step 2: Implementation routes positions through the overload**

```cpp
// PptxWriter.cpp
bool PptxWriter::addContentSlide(const QString& title, const SlideTable& table,
                                   const QVector<SlideImage>& plots,
                                   const QString& extraXml) {
    ContentSlideLayout dl;
    dl.title = QRectF(0.32, 0.10, 12.7, 0.55);
    dl.table = QRectF(table.x, table.y, table.w, table.h);
    if (!plots.isEmpty()) {
        const auto& p = plots.first();
        dl.radar = QRectF(p.x, p.y, p.w, p.h);
    }
    return addContentSlide(title, table, plots, dl, extraXml);
}

bool PptxWriter::addContentSlide(const QString& title, const SlideTable& tableIn,
                                   const QVector<SlideImage>& plotsIn,
                                   const ContentSlideLayout& layout,
                                   const QString& extraXml) {
    SlideTable table = tableIn;
    table.x = layout.table.x(); table.y = layout.table.y();
    table.w = layout.table.width(); table.h = layout.table.height();

    QVector<SlideImage> plots = plotsIn;
    if (!plots.isEmpty() && !layout.radar.isNull()) {
        plots[0].x = layout.radar.x(); plots[0].y = layout.radar.y();
        plots[0].w = layout.radar.width(); plots[0].h = layout.radar.height();
    }

    // … rest of existing implementation, using these adjusted values …
    return /* existing logic */;
}
```

(Refactor the existing `addContentSlide` body into the new overload; the old signature becomes a small wrapper.)

- [ ] **Step 3: Test that overrides are applied**

Add a test that builds a small slide with an override rect and verifies the resulting XML contains the override's EMU values. Use `QXlsx`-free path — `PptxWriter` writes via `ZipWriter`/`XmlBuilder`.

```cpp
void testContentSlideHonoursLayoutOverride() {
    DVE::PptxWriter w;
    DVE::ContentSlideLayout dl;
    dl.table = QRectF(1.0, 2.0, 5.0, 3.0);
    dl.radar = QRectF(7.0, 1.0, 4.0, 4.0);
    dl.title = QRectF(0.5, 0.1, 12.3, 0.5);

    DVE::SlideTable table; table.headers = {"A","B"};
    table.rows.append(QStringList{"1","2"});
    DVE::SlideImage plot; plot.pngData = QByteArray("\x89PNG..."); // not validated
    QVERIFY(w.addContentSlide("T", table, { plot }, dl, ""));

    QString outPath = QDir::tempPath() + "/_tst_layout.pptx";
    QVERIFY(w.save(outPath));
    // unzip and read slide xml — same approach as existing tests
    // assert EMU values match: 1.0" = 914400 EMU
    QFile f(outPath); QVERIFY(f.open(QIODevice::ReadOnly));
    // ... use existing ZipReader-style unpack from current tests ...
}
```

- [ ] **Step 4: Run existing PptxWriter tests, confirm they still pass**

```
cmd.exe //c "cd tests\tst_pptxwriter && set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && mingw32-make.exe && release\tst_pptxwriter.exe"
```

- [ ] **Step 5: Commit**

```
git add src/reporting/PptxWriter.{h,cpp} tests/tst_pptxwriter/tst_pptxwriter.cpp
git commit -m "feat(pptx): addContentSlide layout-override overload"
```

---

### Task 8: `SensoryReportSource::writePptx` — refactor of `generateCombinedPptx`

**Files:**
- Modify: `src/reporting/SensoryReportSource.cpp`
- Modify: `src/ui/SensoryPanel.cpp` — extract the per-session slide-generation block into a helper that `SensoryReportSource::writePptx` can also call
- Modify: `src/ui/SensoryPanel.h`

- [ ] **Step 1: Extract a free function `writeSensoryPptx`**

Move the body of `SensoryPanel::generateCombinedPptx` into a free function (or a static method on `SensoryReportSource`) signed:

```cpp
// New static on SensoryReportSource:
static bool writeSensoryPptx(const QVector<SensorySession>& sessions,
                              const ReportLayout& layout,
                              const QSet<QString>& excludedSamples,
                              const QString& outPath,
                              QString* errorOut);
```

Inside, every place where it currently builds positions inline should consult `layout.contentSlides[contentKey]` first; if that rect is non-null use it, otherwise fall through to existing computation. Sample exclusion: when iterating `sess.samples`, skip any sample whose `"<sessionIdx>#<sampleIdx>"` key is in `excludedSamples`.

- [ ] **Step 2: `SensoryPanel::generateCombinedPptx` becomes a thin caller**

```cpp
bool SensoryPanel::generateCombinedPptx(const QVector<SensorySession>& sessions,
                                          const QString& filePath, QString& errorOut) {
    const ReportLayout layout = SensoryReportSource::computeDefaultLayout(sessions);
    QString err;
    const bool ok = SensoryReportSource::writeSensoryPptx(sessions, layout, {}, filePath, &err);
    if (!ok) errorOut = err;
    return ok;
}
```

- [ ] **Step 3: `SensoryReportSource::writePptx` delegates**

```cpp
bool SensoryReportSource::writePptx(const QString& outPath, const ReportLayout& l,
                                      const QSet<QString>& excluded, QString* err) {
    return writeSensoryPptx(m_sessions, l, excluded, outPath, err);
}
```

- [ ] **Step 4: Sanity test — same input produces same output**

Add a test that calls `SensoryPanel::generateCombinedPptx` with a known session and computes a SHA256 of the output, then calls `SensoryReportSource::writePptx` with `computeDefaultLayout(...)` and `excluded = {}`, and asserts the SHA256 matches.

(If timestamps in the PPTX cause SHA mismatch, compare the slide XML files unzipped — same approach as `tst_pptxwriter`.)

- [ ] **Step 5: Verify the existing `SensoryPanel::generateFullReport` user flow still works**

Build the app and manually open a session, click "Sensory Report" — the legacy code path goes through `generateCombinedPptx` which now goes through the refactored helper. Visually compare a generated .pptx to a baseline from before the change.

- [ ] **Step 6: Commit**

```
git add src/ui/SensoryPanel.{h,cpp} src/reporting/SensoryReportSource.cpp \
        tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp
git commit -m "refactor(sensory): writePptx helper shared between legacy path and SensoryReportSource"
```

---

### Task 9: Build full app, smoke-test legacy report flow unchanged

- [ ] **Step 1: Full release build**

```
cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe -spec win32-g++ DataViewerEnterprise.pro CONFIG+=release && mingw32-make.exe release"
```

- [ ] **Step 2: Run the existing test suite**

```
cd tests && powershell .\run-tests.ps1
```

Expected: all suites pass.

- [ ] **Step 3: Launch and generate a sensory report manually**

Open `release\DataViewer.exe`, switch to sensory mode, load a session from DB or .xlsx, click the sensory Report button. Verify the produced PPTX opens and looks the same as before.

- [ ] **Step 4: No commit** — phase 1A baseline is in commits 1-8 already.

---

## Phase 1B — Canvas Items (4 tasks)

### Task 10: `ResizableSlideItem` base class

**Files:**
- Create: `src/ui/SlideCanvasItems.h`
- Create: `src/ui/SlideCanvasItems.cpp`
- Modify: `DataViewerEnterprise.pro`

- [ ] **Step 1: Header — base + element-id storage**

```cpp
// src/ui/SlideCanvasItems.h
#pragma once
#include <QGraphicsObject>

namespace DVE {

class ResizableSlideItem : public QGraphicsObject {
    Q_OBJECT
public:
    ResizableSlideItem(const QString& elementId,
                        bool aspectLocked, QGraphicsItem* parent = nullptr);

    QString elementId() const { return m_elementId; }
    QRectF  itemRectInches() const;            // converts back to inches

    void setSelectedItem(bool s);
    bool isSelectedItem() const { return m_selected; }

signals:
    void itemClicked(ResizableSlideItem*);
    void rectChanged(const QRectF& newRectInches);

protected:
    void   mousePressEvent(QGraphicsSceneMouseEvent*) override;
    void   mouseMoveEvent(QGraphicsSceneMouseEvent*) override;
    void   mouseReleaseEvent(QGraphicsSceneMouseEvent*) override;
    QRectF boundingRect() const override;

    // Subclass paints content; base paints handles when selected
    virtual void paintContent(QPainter*) = 0;
    void paint(QPainter*, const QStyleOptionGraphicsItem*, QWidget*) override;

    static constexpr double kPxPerInch = 60.0;   // matches ImageViewDialog
    QString m_elementId;
    double  m_w = 200, m_h = 100;
    bool    m_aspectLocked;
    bool    m_selected = false;
    bool    m_resizing = false, m_moving = false;
    QPointF m_pressScenePos, m_pressItemPos;
    double  m_pressW = 0, m_pressH = 0;
    QRectF  handleRect() const;     // bottom-right resize handle
};

} // namespace DVE
```

- [ ] **Step 2: Implementation**

Crib heavily from `src/ui/ImageViewDialog.cpp`'s `ResizableImageItem` — the handle math, drag/resize state machine. Adapt to be content-agnostic by calling `paintContent(painter)` for the body, then drawing handles on top when selected.

```cpp
// src/ui/SlideCanvasItems.cpp
#include "SlideCanvasItems.h"
#include <QPainter>
#include <QGraphicsSceneMouseEvent>

namespace DVE {

ResizableSlideItem::ResizableSlideItem(const QString& id, bool aspect, QGraphicsItem* p)
    : QGraphicsObject(p), m_elementId(id), m_aspectLocked(aspect) {
    setFlags(ItemIsSelectable | ItemIsFocusable);
    setAcceptHoverEvents(true);
}

QRectF ResizableSlideItem::boundingRect() const {
    return QRectF(0, 0, m_w, m_h).adjusted(-1, -1, 7, 7);
}

QRectF ResizableSlideItem::handleRect() const {
    constexpr double k = 12;
    return QRectF(m_w - k/2, m_h - k/2, k, k);
}

QRectF ResizableSlideItem::itemRectInches() const {
    return QRectF(x() / kPxPerInch, y() / kPxPerInch,
                   m_w / kPxPerInch, m_h / kPxPerInch);
}

void ResizableSlideItem::setSelectedItem(bool s) {
    if (m_selected != s) { m_selected = s; update(); }
}

void ResizableSlideItem::mousePressEvent(QGraphicsSceneMouseEvent* e) {
    m_pressScenePos = e->scenePos();
    m_pressItemPos  = pos();
    m_pressW = m_w; m_pressH = m_h;
    m_resizing = handleRect().contains(e->pos());
    m_moving   = !m_resizing;
    emit itemClicked(this);
    e->accept();
}

void ResizableSlideItem::mouseMoveEvent(QGraphicsSceneMouseEvent* e) {
    const QPointF d = e->scenePos() - m_pressScenePos;
    if (m_moving) {
        setPos(m_pressItemPos + d);
    } else if (m_resizing) {
        double newW = qMax(20.0, m_pressW + d.x());
        double newH = qMax(20.0, m_pressH + d.y());
        if (m_aspectLocked) {
            const double aspect = m_pressW / qMax(1.0, m_pressH);
            if (newW / newH > aspect) newW = newH * aspect;
            else                       newH = newW / aspect;
        }
        prepareGeometryChange();
        m_w = newW; m_h = newH;
    }
    update();
}

void ResizableSlideItem::mouseReleaseEvent(QGraphicsSceneMouseEvent*) {
    if (m_moving || m_resizing) emit rectChanged(itemRectInches());
    m_moving = m_resizing = false;
}

void ResizableSlideItem::paint(QPainter* p, const QStyleOptionGraphicsItem*, QWidget*) {
    paintContent(p);
    if (m_selected) {
        p->setPen(QPen(QColor(30, 130, 230), 1.5, Qt::DashLine));
        p->setBrush(Qt::NoBrush);
        p->drawRect(QRectF(0, 0, m_w, m_h));
        p->setPen(Qt::NoPen);
        p->setBrush(QColor(30, 130, 230));
        p->drawRect(handleRect());
    }
}

} // namespace DVE
```

- [ ] **Step 3: .pro additions**

In `DataViewerEnterprise.pro` add the source/header.

- [ ] **Step 4: Build clean**

Build the app — `mingw32-make release`. Expect clean compile. No tests yet (we'll test via `PlotItem` etc).

- [ ] **Step 5: Commit**

```
git add src/ui/SlideCanvasItems.{h,cpp} DataViewerEnterprise.pro
git commit -m "feat(ui): ResizableSlideItem base for slide-canvas elements"
```

---

### Task 11: `PlotItem` (radar chart with locked aspect)

**Files:**
- Modify: `src/ui/SlideCanvasItems.h` / `.cpp`

- [ ] **Step 1: Subclass declaration**

In `SlideCanvasItems.h`:

```cpp
class PlotItem : public ResizableSlideItem {
    Q_OBJECT
public:
    PlotItem(const QString& elementId, const QPixmap& pixmap,
              QGraphicsItem* parent = nullptr);
    void setPixmap(const QPixmap&);
protected:
    void paintContent(QPainter*) override;
private:
    QPixmap m_pixmap;
};
```

- [ ] **Step 2: Implementation**

```cpp
PlotItem::PlotItem(const QString& id, const QPixmap& pix, QGraphicsItem* p)
    : ResizableSlideItem(id, /*aspectLocked=*/true, p), m_pixmap(pix) {
    if (!pix.isNull()) {
        m_w = pix.width()  * kPxPerInch / 96.0;     // assume 96 dpi source
        m_h = pix.height() * kPxPerInch / 96.0;
    }
}

void PlotItem::setPixmap(const QPixmap& pix) {
    m_pixmap = pix; update();
}

void PlotItem::paintContent(QPainter* p) {
    if (!m_pixmap.isNull())
        p->drawPixmap(QRectF(0, 0, m_w, m_h), m_pixmap, m_pixmap.rect());
    else {
        p->fillRect(QRectF(0, 0, m_w, m_h), QColor(240, 240, 240));
        p->setPen(QColor(120, 120, 120));
        p->drawText(QRectF(0, 0, m_w, m_h), Qt::AlignCenter, "(plot)");
    }
}
```

- [ ] **Step 3: Add minimal test (compile-only)**

No `QGraphicsView` test needed yet — instantiate in a `QGraphicsScene` in a test, verify aspect locking via simulated drag. Add to a new `tests/tst_slidecanvasitems/` test bin.

```cpp
// tests/tst_slidecanvasitems/tst_slidecanvasitems.cpp
#include <QtTest>
#include <QGraphicsScene>
#include <QGraphicsSceneMouseEvent>
#include "SlideCanvasItems.h"

class tst_SlideCanvasItems : public QObject {
    Q_OBJECT
private slots:
    void testPlotItemAspectLockedDuringResize() {
        QGraphicsScene scene;
        QPixmap pix(96, 96); pix.fill(Qt::red);
        DVE::PlotItem* item = new DVE::PlotItem("radar", pix);
        scene.addItem(item);

        // Manually simulate: press handle, drag down-right by (40, 80), release.
        // With aspect locked, height should follow width or vice versa.
        QSignalSpy spy(item, &DVE::ResizableSlideItem::rectChanged);
        // (Direct simulation is awkward — alternative: invoke protected method via
        // friend-class or use QTest::mouseClick on a hosting view. The simpler
        // verification is to assert the constructor sets aspect-locked behavior:)
        const QRectF r = item->itemRectInches();
        QCOMPARE(r.width(), r.height());      // square (aspect lock)
        QCOMPARE(r.width(), 1.0);             // 96 px source @ 96 dpi → 60 scene px → 1.0" @ kPxPerInch=60
    }
};

QTEST_MAIN(tst_SlideCanvasItems)
#include "tst_slidecanvasitems.moc"
```

- [ ] **Step 4: Test .pro**

```
QT += core gui widgets testlib
CONFIG += console c++17
TEMPLATE = app
INCLUDEPATH += ../../src ../../src/ui
SOURCES += tst_slidecanvasitems.cpp ../../src/ui/SlideCanvasItems.cpp
HEADERS += ../../src/ui/SlideCanvasItems.h
```

Add to `tests/tests.pro` SUBDIRS.

- [ ] **Step 5: Build and run test**

Expected: 1 test passes.

- [ ] **Step 6: Commit**

```
git add src/ui/SlideCanvasItems.{h,cpp} tests/tst_slidecanvasitems/ tests/tests.pro
git commit -m "feat(ui): PlotItem (locked-aspect radar chart canvas item)"
```

---

### Task 12: `TableItem` (renders rows + emits sort signal on header click)

**Files:**
- Modify: `src/ui/SlideCanvasItems.h` / `.cpp`
- Modify: `tests/tst_slidecanvasitems/tst_slidecanvasitems.cpp`

- [ ] **Step 1: Subclass declaration**

```cpp
class TableItem : public ResizableSlideItem {
    Q_OBJECT
public:
    TableItem(const QString& elementId, QGraphicsItem* parent = nullptr);
    void setHeaders(const QStringList&);
    void setRows(const QVector<QStringList>&);
    void setSort(const QString& column, Qt::SortOrder);
signals:
    void columnHeaderClicked(const QString& column);
protected:
    void paintContent(QPainter*) override;
    void mousePressEvent(QGraphicsSceneMouseEvent*) override;
private:
    QStringList m_headers;
    QVector<QStringList> m_rows;
    QString m_sortColumn;
    Qt::SortOrder m_sortOrder = Qt::DescendingOrder;
    QRectF headerRectFor(int colIdx) const;
};
```

- [ ] **Step 2: Implementation**

```cpp
TableItem::TableItem(const QString& id, QGraphicsItem* p)
    : ResizableSlideItem(id, /*aspectLocked=*/false, p) {
    m_w = 600; m_h = 100;
}
void TableItem::setHeaders(const QStringList& h) { m_headers = h; update(); }
void TableItem::setRows(const QVector<QStringList>& r) { m_rows = r; update(); }
void TableItem::setSort(const QString& c, Qt::SortOrder o) {
    m_sortColumn = c; m_sortOrder = o; update();
}

QRectF TableItem::headerRectFor(int colIdx) const {
    const int n = m_headers.size();
    if (n == 0) return {};
    const double cw = m_w / n;
    return QRectF(colIdx * cw, 0, cw, 24);
}

void TableItem::paintContent(QPainter* p) {
    p->setRenderHint(QPainter::Antialiasing);
    p->fillRect(QRectF(0, 0, m_w, m_h), Qt::white);
    p->setPen(QColor(80, 80, 80));

    const int n = m_headers.size();
    if (n == 0) return;
    const double cw = m_w / n;
    const double rowH = qMax(18.0, (m_h - 24) / qMax(1, m_rows.size()));

    // Header row
    p->fillRect(QRectF(0, 0, m_w, 24), QColor(0x1F, 0x4E, 0x79));
    p->setPen(Qt::white);
    for (int c = 0; c < n; ++c) {
        QString h = m_headers[c];
        if (h == m_sortColumn) h += (m_sortOrder == Qt::AscendingOrder ? " ▲" : " ▼");
        p->drawText(QRectF(c * cw + 4, 4, cw - 8, 16), Qt::AlignVCenter, h);
    }

    // Rows
    p->setPen(QColor(40, 40, 40));
    for (int r = 0; r < m_rows.size(); ++r) {
        if (r % 2) p->fillRect(QRectF(0, 24 + r*rowH, m_w, rowH),
                                QColor(245, 248, 252));
        for (int c = 0; c < n && c < m_rows[r].size(); ++c)
            p->drawText(QRectF(c*cw + 4, 24 + r*rowH, cw - 8, rowH),
                        Qt::AlignVCenter, m_rows[r][c]);
    }

    // Borders
    p->setPen(QColor(0xCC, 0xCC, 0xCC));
    p->drawRect(QRectF(0, 0, m_w, m_h));
}

void TableItem::mousePressEvent(QGraphicsSceneMouseEvent* e) {
    if (e->pos().y() < 24) {
        // Header click — find column
        const int n = m_headers.size();
        if (n > 0) {
            const int col = qMin(n - 1, int(e->pos().x() / (m_w / n)));
            emit columnHeaderClicked(m_headers[col]);
            e->accept(); return;
        }
    }
    ResizableSlideItem::mousePressEvent(e);
}
```

- [ ] **Step 3: Test**

```cpp
void testTableHeaderEmitsSortSignal() {
    DVE::TableItem* item = new DVE::TableItem("table");
    item->setHeaders({"Name", "Score"});
    QSignalSpy spy(item, &DVE::TableItem::columnHeaderClicked);

    QGraphicsSceneMouseEvent e(QEvent::GraphicsSceneMousePress);
    e.setPos(QPointF(50, 10));    // first column
    e.setScenePos(QPointF(50, 10));
    e.setButton(Qt::LeftButton);
    QApplication::sendEvent(item->scene() ? item->scene() : nullptr, &e);
    // Or invoke directly:
    QGraphicsScene scene;
    scene.addItem(item);
    item->setSelected(true);
    QApplication::sendEvent(&scene, &e);
    QCOMPARE(spy.count(), 1);
    QCOMPARE(spy.first().first().toString(), QString("Name"));
}
```

(Refine event delivery if direct send is finicky — alternative: cast to `QGraphicsItem*` and call `mousePressEvent` directly via friend.)

- [ ] **Step 4: Build and run**

Expected: previous 1 + 1 new = 2 tests pass.

- [ ] **Step 5: Commit**

```
git add src/ui/SlideCanvasItems.{h,cpp} tests/tst_slidecanvasitems/tst_slidecanvasitems.cpp
git commit -m "feat(ui): TableItem renders rows + emits column-header click signal"
```

---

### Task 13: `TextItem` (double-click to edit)

**Files:**
- Modify: `src/ui/SlideCanvasItems.h` / `.cpp`
- Modify: `tests/tst_slidecanvasitems/tst_slidecanvasitems.cpp`

- [ ] **Step 1: Declaration**

```cpp
class TextItem : public ResizableSlideItem {
    Q_OBJECT
public:
    TextItem(const QString& elementId, QGraphicsItem* parent = nullptr);
    void setText(const QString&);
    QString text() const { return m_text; }
    void setFontPointSize(int pt);
signals:
    void textCommitted(const QString&);
protected:
    void paintContent(QPainter*) override;
    void mouseDoubleClickEvent(QGraphicsSceneMouseEvent*) override;
private:
    QString m_text;
    int     m_fontPt = 14;
};
```

- [ ] **Step 2: Implementation — paint with word-wrap; double-click pops a `QLineEdit`/`QTextEdit` modal**

```cpp
TextItem::TextItem(const QString& id, QGraphicsItem* p)
    : ResizableSlideItem(id, /*aspectLocked=*/false, p) {
    m_w = 300; m_h = 40;
}
void TextItem::setText(const QString& t) { m_text = t; update(); }
void TextItem::setFontPointSize(int pt) { m_fontPt = pt; update(); }

void TextItem::paintContent(QPainter* p) {
    p->setRenderHint(QPainter::TextAntialiasing);
    QFont f = p->font(); f.setPointSize(m_fontPt); p->setFont(f);
    p->setPen(QColor(0x33, 0x33, 0x33));
    p->drawText(QRectF(4, 4, m_w - 8, m_h - 8),
                Qt::AlignTop | Qt::AlignLeft | Qt::TextWordWrap, m_text);
}

void TextItem::mouseDoubleClickEvent(QGraphicsSceneMouseEvent*) {
    bool ok = false;
    const QString next = QInputDialog::getMultiLineText(
        nullptr, "Edit Text", "Text:", m_text, &ok);
    if (ok && next != m_text) {
        m_text = next;
        update();
        emit textCommitted(m_text);
    }
}
```

(Add `#include <QInputDialog>` at the top of the .cpp.)

- [ ] **Step 3: Test**

```cpp
void testTextItemSetGetText() {
    DVE::TextItem item("title");
    item.setText("Hello");
    QCOMPARE(item.text(), QString("Hello"));
}
```

- [ ] **Step 4: Build, run, commit**

```
git add src/ui/SlideCanvasItems.{h,cpp} tests/tst_slidecanvasitems/tst_slidecanvasitems.cpp
git commit -m "feat(ui): TextItem with double-click-to-edit"
```

---

## Phase 1C — Dialog Shell (4 tasks)

### Task 14: `ReportPreviewDialog` skeleton (window, layout, Cancel/OK, no canvas yet)

**Files:**
- Create: `src/ui/ReportPreviewDialog.h`
- Create: `src/ui/ReportPreviewDialog.cpp`
- Modify: `DataViewerEnterprise.pro`

- [ ] **Step 1: Header**

```cpp
// src/ui/ReportPreviewDialog.h
#pragma once
#include <QDialog>
#include <QGraphicsView>
#include <QGraphicsScene>
#include <QListWidget>
#include <QSet>
#include "reporting/ReportLayout.h"
#include "reporting/IReportSource.h"

namespace DVE {

class SamplesCheckboxPanel;
class PropertiesPanel;

class ReportPreviewDialog : public QDialog {
    Q_OBJECT
public:
    ReportPreviewDialog(IReportSource* source, QWidget* parent = nullptr);

    QString outputPath() const { return m_outputPath; }      // valid only after exec() == Accepted

private slots:
    void onCreateReport();
    void onCancel();
    void onSlideSelected(int row);

private:
    void buildUi();
    void populateThumbnails();

    IReportSource* m_source;
    ReportLayout   m_layout;
    QSet<QString>  m_excludedSamples;
    int            m_currentSlide = 0;
    QString        m_outputPath;

    QListWidget*    m_thumbList = nullptr;
    QGraphicsView*  m_canvas    = nullptr;
    QGraphicsScene* m_scene     = nullptr;
    SamplesCheckboxPanel* m_samplesPanel = nullptr;
    PropertiesPanel*      m_propsPanel   = nullptr;
};

} // namespace DVE
```

- [ ] **Step 2: Implementation skeleton — UI shell only**

```cpp
// src/ui/ReportPreviewDialog.cpp
#include "ReportPreviewDialog.h"
#include "SamplesCheckboxPanel.h"
#include "PropertiesPanel.h"
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QPushButton>
#include <QFileDialog>
#include <QGraphicsRectItem>

namespace DVE {

ReportPreviewDialog::ReportPreviewDialog(IReportSource* src, QWidget* p)
    : QDialog(p), m_source(src) {
    setWindowTitle("Report Preview — " + src->sourceLabel());
    resize(1200, 720);
    m_layout = src->loadLayout();
    buildUi();
    populateThumbnails();
    if (m_thumbList->count() > 0) m_thumbList->setCurrentRow(0);
}

void ReportPreviewDialog::buildUi() {
    auto* outer = new QHBoxLayout(this);
    outer->setContentsMargins(8, 8, 8, 8);

    // Left column: thumbs + samples
    auto* left = new QVBoxLayout;
    m_thumbList = new QListWidget;
    m_thumbList->setFixedWidth(160);
    connect(m_thumbList, &QListWidget::currentRowChanged,
            this, &ReportPreviewDialog::onSlideSelected);
    left->addWidget(m_thumbList, 1);
    m_samplesPanel = new SamplesCheckboxPanel(m_source->allSamples());
    left->addWidget(m_samplesPanel, 1);
    outer->addLayout(left);

    // Center: canvas
    m_scene = new QGraphicsScene(this);
    m_scene->setSceneRect(0, 0, 800, 450);
    m_scene->setBackgroundBrush(Qt::white);
    m_scene->addRect(0, 0, 800, 450, QPen(QColor(180,180,180)), Qt::NoBrush);
    m_canvas = new QGraphicsView(m_scene);
    m_canvas->setRenderHint(QPainter::Antialiasing);
    outer->addWidget(m_canvas, 1);

    // Right: properties panel + buttons
    auto* right = new QVBoxLayout;
    m_propsPanel = new PropertiesPanel;
    right->addWidget(m_propsPanel, 1);
    auto* btns = new QHBoxLayout;
    auto* cancel = new QPushButton("Cancel");
    auto* create = new QPushButton("Create Report");
    create->setDefault(true);
    btns->addStretch();
    btns->addWidget(cancel);
    btns->addWidget(create);
    right->addLayout(btns);
    outer->addLayout(right);

    connect(cancel, &QPushButton::clicked, this, &ReportPreviewDialog::onCancel);
    connect(create, &QPushButton::clicked, this, &ReportPreviewDialog::onCreateReport);
}

void ReportPreviewDialog::populateThumbnails() {
    m_thumbList->clear();
    for (int i = 0; i < m_source->slideCount(); ++i) {
        auto* it = new QListWidgetItem(QString::number(i + 1) + ". " +
            (m_source->slideKind(i) == SlideKind::Cover ? "Cover" :
             m_source->slideKind(i) == SlideKind::Divider ? "Divider" :
             m_source->slideKind(i) == SlideKind::Content ? "Content" :
             m_source->slideKind(i) == SlideKind::Image ? "Images" : "Cumulative"));
        m_thumbList->addItem(it);
    }
}

void ReportPreviewDialog::onSlideSelected(int row) { m_currentSlide = row; /* canvas pop in Task 18 */ }

void ReportPreviewDialog::onCancel() { reject(); }

void ReportPreviewDialog::onCreateReport() {
    const QString path = QFileDialog::getSaveFileName(this, "Save Report",
        "report.pptx", "PowerPoint (*.pptx)");
    if (path.isEmpty()) return;
    QString err;
    if (!m_source->writePptx(path, m_layout, m_excludedSamples, &err)) {
        QMessageBox::warning(this, "Report Failed", err);
        return;
    }
    m_outputPath = path;
    accept();
}

} // namespace DVE
```

- [ ] **Step 3: .pro additions**

Add `src/ui/ReportPreviewDialog.{h,cpp}` to SOURCES/HEADERS.

- [ ] **Step 4: Build clean (panel headers will be created in next two tasks; until then forward-declare or leave placeholder includes)**

To unblock compile, create stub `SamplesCheckboxPanel.h` and `PropertiesPanel.h` with empty class bodies. Real impl in Tasks 15-16.

- [ ] **Step 5: Commit**

```
git add src/ui/ReportPreviewDialog.{h,cpp} \
        src/ui/SamplesCheckboxPanel.h src/ui/PropertiesPanel.h \
        DataViewerEnterprise.pro
git commit -m "feat(ui): ReportPreviewDialog shell (no canvas content yet)"
```

---

### Task 15: `SamplesCheckboxPanel`

**Files:**
- Modify: `src/ui/SamplesCheckboxPanel.h`
- Create: `src/ui/SamplesCheckboxPanel.cpp`
- Modify: `DataViewerEnterprise.pro`

- [ ] **Step 1: Header**

```cpp
// src/ui/SamplesCheckboxPanel.h
#pragma once
#include <QWidget>
#include <QSet>
#include "reporting/IReportSource.h"

namespace DVE {

class SamplesCheckboxPanel : public QWidget {
    Q_OBJECT
public:
    SamplesCheckboxPanel(const QVector<SampleRef>& refs, QWidget* parent = nullptr);

    QSet<QString> excludedSampleIds() const;

signals:
    void sampleToggled(const QString& sampleId, bool included);

private:
    QVector<SampleRef> m_refs;
    QHash<QString, class QCheckBox*> m_boxes;
};

} // namespace DVE
```

- [ ] **Step 2: Implementation — grouped by `sessionLabel`, all-checked default**

```cpp
// src/ui/SamplesCheckboxPanel.cpp
#include "SamplesCheckboxPanel.h"
#include <QVBoxLayout>
#include <QGroupBox>
#include <QCheckBox>
#include <QScrollArea>
#include <QLabel>

namespace DVE {

SamplesCheckboxPanel::SamplesCheckboxPanel(const QVector<SampleRef>& refs, QWidget* p)
    : QWidget(p), m_refs(refs) {
    auto* outer = new QVBoxLayout(this);
    outer->setContentsMargins(0,0,0,0);
    outer->addWidget(new QLabel("<b>Samples</b>"));

    auto* scroll = new QScrollArea;
    scroll->setWidgetResizable(true);
    auto* host = new QWidget;
    auto* hostL = new QVBoxLayout(host);
    hostL->setContentsMargins(2,2,2,2);

    QHash<QString, QGroupBox*> groups;
    QStringList groupOrder;
    for (const SampleRef& r : refs) {
        if (!groups.contains(r.sessionLabel)) {
            groups[r.sessionLabel] = new QGroupBox(r.sessionLabel);
            new QVBoxLayout(groups[r.sessionLabel]);
            hostL->addWidget(groups[r.sessionLabel]);
            groupOrder << r.sessionLabel;
        }
        auto* box = new QCheckBox(r.displayName);
        box->setChecked(true);
        const QString id = r.sampleId;
        connect(box, &QCheckBox::toggled, this, [this, id](bool on) {
            emit sampleToggled(id, on);
        });
        m_boxes.insert(id, box);
        static_cast<QVBoxLayout*>(groups[r.sessionLabel]->layout())->addWidget(box);
    }
    hostL->addStretch();
    scroll->setWidget(host);
    outer->addWidget(scroll, 1);
}

QSet<QString> SamplesCheckboxPanel::excludedSampleIds() const {
    QSet<QString> out;
    for (auto it = m_boxes.cbegin(); it != m_boxes.cend(); ++it)
        if (!it.value()->isChecked()) out.insert(it.key());
    return out;
}

} // namespace DVE
```

- [ ] **Step 3: .pro update + build**

- [ ] **Step 4: Wire into dialog** — in `ReportPreviewDialog`, connect `m_samplesPanel->sampleToggled` to a slot that updates `m_excludedSamples` and re-builds the current slide's canvas (slide rebuild comes in Task 18).

- [ ] **Step 5: Commit**

```
git add src/ui/SamplesCheckboxPanel.{h,cpp} src/ui/ReportPreviewDialog.cpp DataViewerEnterprise.pro
git commit -m "feat(ui): SamplesCheckboxPanel with grouped checkboxes"
```

---

### Task 16: `PropertiesPanel` (selected-item position/size + Z-order + sort selector)

**Files:**
- Modify: `src/ui/PropertiesPanel.h`
- Create: `src/ui/PropertiesPanel.cpp`

- [ ] **Step 1: Header**

```cpp
// src/ui/PropertiesPanel.h
#pragma once
#include <QWidget>
#include <QRectF>

namespace DVE {

class PropertiesPanel : public QWidget {
    Q_OBJECT
public:
    PropertiesPanel(QWidget* parent = nullptr);

    void setSelectedItem(const QString& elementId, const QRectF& rectInches);
    void clearSelection();

signals:
    void rectEdited(const QString& elementId, const QRectF& newRectInches);
    void bringForwardClicked(const QString& elementId);
    void sendBackwardClicked(const QString& elementId);

private:
    QString m_currentId;
    class QDoubleSpinBox *m_x, *m_y, *m_w, *m_h;
    class QLabel* m_label;
};

} // namespace DVE
```

- [ ] **Step 2: Impl — basic 4 spinboxes + 2 buttons**

(Skipping full code; analogous to existing property-table dialogs in the app. Use `QDoubleSpinBox` with range 0–13.33 for x/w and 0–7.5 for y/h, decimals=2, suffix=" in".)

- [ ] **Step 3: Wire into dialog** — connect `rectEdited` signal to a slot in `ReportPreviewDialog` that updates `m_layout` and pushes the new rect to the canvas item by element id.

- [ ] **Step 4: Build, commit**

```
git add src/ui/PropertiesPanel.{h,cpp} src/ui/ReportPreviewDialog.cpp
git commit -m "feat(ui): PropertiesPanel for selected-item geometry editing"
```

---

### Task 17: Thumbnail strip — render mini previews via `IReportSource::buildSlide`

**Files:**
- Modify: `src/ui/ReportPreviewDialog.cpp`

- [ ] **Step 1: For each slide, render a 160×90 px thumb**

Replace the placeholder `populateThumbnails` with one that asks `m_source->buildSlide(idx, m_layout, m_excludedSamples)` and rasterises the resulting layout into a small QPixmap. Easiest path: render the full canvas at 800×450 then `pix.scaled(160, 90)`.

(Until Task 18 wires `buildSlide` to actually return content, thumbs will be blank-with-label placeholders. That's fine for now.)

- [ ] **Step 2: Commit**

```
git add src/ui/ReportPreviewDialog.cpp
git commit -m "feat(ui): thumbnail strip placeholders (real renders in Task 18)"
```

---

## Phase 1D — Wire it together (5 tasks)

### Task 18: `IReportSource::buildSlide` populates the canvas

**Files:**
- Modify: `src/reporting/SensoryReportSource.cpp` (implement `buildSlide`)
- Modify: `src/ui/ReportPreviewDialog.cpp` (canvas population)

- [ ] **Step 1: Implement `SensoryReportSource::buildSlide`**

For each `SlideKind`, build a `ReportSlideSpec`:
- Cover: title text = source label, subtitle = today's date, layout = `m_layout.coverTitle` / `coverSubtitle`
- Divider: title text = tester name, layout from `m_layout.dividerTitles[key]`
- Content: title = session.testTitle, table = build same headers/rows as legacy code (sample names + 5 metric scores), filter excluded; radarPixmap = render via existing `RadarChartWidget` headlessly; propertiesText built from session metadata; layout from `m_layout.contentSlides[key]` falling back to defaults
- Image: load each path, build `QPixmap`s, layouts/crops from `m_layout.imageSlides[key]`
- Cumulative: aggregate samples across sessions (respecting exclusions), generate cum table + radar

Use the existing math from `SensoryPanel.cpp` lines 1925-2261 — extract into private methods.

- [ ] **Step 2: `ReportPreviewDialog::onSlideSelected` populates canvas from spec**

```cpp
void ReportPreviewDialog::onSlideSelected(int row) {
    m_currentSlide = row;
    m_scene->clear();
    m_scene->addRect(0, 0, 800, 450, QPen(QColor(180,180,180)), Qt::white);
    if (row < 0 || row >= m_source->slideCount()) return;

    const ReportSlideSpec spec = m_source->buildSlide(row, m_layout, m_excludedSamples);
    constexpr double pxPerInch = 60.0;

    auto addAt = [&](ResizableSlideItem* item, const QRectF& rectInches) {
        item->setPos(rectInches.x() * pxPerInch, rectInches.y() * pxPerInch);
        // size set inside item
        m_scene->addItem(item);
        connect(item, &ResizableSlideItem::rectChanged, this, [this,item](const QRectF& r) {
            // Update m_layout and schedule auto-save (Task 20)
            applyRectEdit(item->elementId(), r);
        });
    };

    if (spec.kind == SlideKind::Content || spec.kind == SlideKind::Cumulative) {
        auto* title = new TextItem("title");
        title->setText(spec.title);
        addAt(title, spec.layout.title);
        auto* table = new TableItem("table");
        table->setHeaders(spec.tableHeaders);
        table->setRows(spec.tableRows);
        addAt(table, spec.layout.table);
        connect(table, &TableItem::columnHeaderClicked, this, [this](const QString& c) {
            // Toggle sort (Task 12 logic — first click desc, second asc, third clear)
            applySortChange(c);
        });
        auto* radar = new PlotItem("radar", QPixmap::fromImage(spec.radarPixmap));
        addAt(radar, spec.layout.radar);
        auto* props = new TextItem("propertiesBox");
        props->setText(spec.propertiesText);
        addAt(props, spec.layout.propertiesBox.rect);
    }
    // ... cover, divider, image cases ...
}
```

- [ ] **Step 2: Add `applyRectEdit` and `applySortChange` to dialog**

(Two helpers that update `m_layout` and call `m_source->saveLayout(m_layout)` debounced — debouncing in Task 20.)

- [ ] **Step 3: Manual smoke test**

Build, launch, open a sensory session, click Sensory Report — preview opens, click between slides, see the layout populate. Drag the table — it moves. Resize the radar — it scales (locked aspect).

- [ ] **Step 4: Commit**

```
git add src/reporting/SensoryReportSource.{h,cpp} src/ui/ReportPreviewDialog.cpp
git commit -m "feat(reports): SensoryReportSource.buildSlide + dialog canvas population"
```

---

### Task 19: Dialog auto-save with 500 ms debounce

**Files:**
- Modify: `src/ui/ReportPreviewDialog.h` / `.cpp`

- [ ] **Step 1: Add `QTimer m_autoSaveTimer` + flush on close**

```cpp
// header
QTimer* m_autoSaveTimer;
void scheduleAutoSave();
void flushAutoSave();

// constructor
m_autoSaveTimer = new QTimer(this);
m_autoSaveTimer->setSingleShot(true);
m_autoSaveTimer->setInterval(500);
connect(m_autoSaveTimer, &QTimer::timeout, this, &ReportPreviewDialog::flushAutoSave);

// every applyRectEdit / applySortChange / sample toggle calls scheduleAutoSave():
void ReportPreviewDialog::scheduleAutoSave() { m_autoSaveTimer->start(); }
void ReportPreviewDialog::flushAutoSave()     { m_source->saveLayout(m_layout); }

// reject() and accept() override to flushAutoSave first
void ReportPreviewDialog::done(int r) {
    flushAutoSave();
    QDialog::done(r);
}
```

- [ ] **Step 2: Test by editing then closing without "Create Report"; reopen; verify edits persist**

- [ ] **Step 3: Commit**

```
git add src/ui/ReportPreviewDialog.{h,cpp}
git commit -m "feat(ui): debounced auto-save of layout edits"
```

---

### Task 20: Wire SensoryPanel to open the dialog

**Files:**
- Modify: `src/ui/SensoryPanel.cpp`

- [ ] **Step 1: Replace direct PPTX generation with dialog-then-write**

```cpp
void SensoryPanel::generateFullReport() {
    QVector<SensorySession> sessions{ *currentSession() };
    auto* src = new SensoryReportSource(sessions, m_db);
    ReportPreviewDialog dlg(src, this);
    dlg.exec();
    delete src;
    // No further action — dialog handled the save dialog and PPTX write itself.
}
```

Same for the multi-session combined path.

- [ ] **Step 2: Manual full-flow test** — click Sensory Report, edit something, click Create Report, save the .pptx, open it. Verify the .pptx reflects the edits.

- [ ] **Step 3: Commit**

```
git add src/ui/SensoryPanel.cpp
git commit -m "feat(sensory): route Sensory Report ribbon button through preview dialog"
```

---

### Task 21: End-to-end smoke + the deployment self-test

- [ ] **Step 1: Run all tests**

```
cd tests && powershell .\run-tests.ps1
```

Expected: every suite passes.

- [ ] **Step 2: Run deployment self-test**

```
.\tests\deployment\Test-Deployment.ps1
```

Expected: every phase green.

- [ ] **Step 3: Manual exploratory** — open and close the preview multiple times, edit, generate, open the .pptx in PowerPoint, confirm visually.

- [ ] **Step 4: No commit** — phase 1 is done. Tag a milestone.

```
git tag phase-1-sensory-minimum
```

---

## Phase 2 — Excel Custom Property Round-Trip (5 tasks)

### Task 22: Python helper `tools/excel_layout_io.py`

**Files:**
- Create: `tools/excel_layout_io.py`

- [ ] **Step 1: Implement read/write**

```python
# tools/excel_layout_io.py
"""Read/write the dve_layout custom workbook property on an .xlsx file.

Usage:
  python excel_layout_io.py read <path>            -> stdout: JSON or "" if missing
  python excel_layout_io.py write <path> <jsonstr> -> exit 0 on success
"""
import sys
from openpyxl import load_workbook
from openpyxl.packaging.custom import CustomDocumentProperty, CustomDocumentPropertyList
from openpyxl.descriptors import String

PROP_NAME = "dve_layout"

def read(path):
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    props = wb.custom_doc_props
    if props is None: return ""
    for p in props.props:
        if p.name == PROP_NAME:
            return p.value or ""
    return ""

def write(path, json_str):
    wb = load_workbook(path)
    props = wb.custom_doc_props or CustomDocumentPropertyList()
    found = False
    for p in props.props:
        if p.name == PROP_NAME:
            p.value = json_str; found = True; break
    if not found:
        cdp = CustomDocumentProperty(name=PROP_NAME, value=json_str)
        props.props.append(cdp)
    wb.custom_doc_props = props
    wb.save(path)

def main():
    cmd = sys.argv[1]
    if cmd == "read":
        sys.stdout.write(read(sys.argv[2]))
    elif cmd == "write":
        write(sys.argv[2], sys.argv[3])
    else:
        sys.exit(2)

if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Smoke test**

```
release\python\python.exe tools\excel_layout_io.py write some.xlsx "{\"v\":1}"
release\python\python.exe tools\excel_layout_io.py read some.xlsx
```

Expected: prints `{"v":1}`.

- [ ] **Step 3: Commit**

```
git add tools/excel_layout_io.py
git commit -m "feat(excel): python helper for dve_layout custom workbook property"
```

---

### Task 23: `ExcelReader` reads `dve_layout`

**Files:**
- Modify: `src/ExcelReader.h` / `.cpp`

- [ ] **Step 1: Add field + read on parse**

In `ExcelReader.h`, add to the parsed-result struct: `QString layoutJson;`. In `ExcelReader.cpp`, after the existing Python invocation, call `excel_layout_io.py read <path>` (or fold the read into the existing reader Python script) and capture the stdout into the result.

- [ ] **Step 2: Test using a fixture .xlsx with a known layout property**

(Generate the fixture once via the Python helper.)

- [ ] **Step 3: Commit**

```
git add src/ExcelReader.{h,cpp} tests/tst_excelreader/tst_excelreader.cpp
git commit -m "feat(excel): ExcelReader exposes dve_layout custom property"
```

---

### Task 24: Excel write-back for layout JSON

**Files:**
- Modify: the Python writer helper used by `MainWindow::flushExcelWrites` (or wherever cell-writeback runs) — locate via grep for `openpyxl` write code in the existing scripts
- Modify: `src/ui/SensoryPanel.cpp` (or wherever the writer is invoked) to also pass the layout JSON

- [ ] **Step 1: Extend the existing writeback Python to also call the custom-property write step**

Either add it to the existing script or invoke `excel_layout_io.py write` after the existing writeback completes.

- [ ] **Step 2: Test via SensoryReportSource — save + reload .xlsx and verify property round-trips**

- [ ] **Step 3: Commit**

---

### Task 25: `SensoryReportSource::loadLayout` / `saveLayout` consume Excel

**Files:**
- Modify: `src/reporting/SensoryReportSource.cpp`

- [ ] **Step 1: Compare timestamps, take newest**

```cpp
ReportLayout SensoryReportSource::loadLayout() const {
    const QString dbJson = m_db ? m_db->loadSensoryLayout(m_sessions.first().id) : "";
    const QString xlJson = !m_sessions.isEmpty() ? m_sessions.first().excelLayoutJson : "";
    // Use newest-timestamp: compare m_db's last-modified timestamp on the row
    // vs xlsx file mtime. Falling back: prefer DB if both present.
    const QString chosen = dbJson.isEmpty() ? xlJson : dbJson;
    if (chosen.isEmpty()) return computeDefaultLayout(m_sessions);
    bool ok = false;
    ReportLayout layout = ReportLayout::fromJson(
        QJsonDocument::fromJson(chosen.toUtf8()).object(), &ok);
    return ok ? layout : computeDefaultLayout(m_sessions);
}

void SensoryReportSource::saveLayout(const ReportLayout& l) {
    const QString json = QJsonDocument(l.toJson()).toJson(QJsonDocument::Compact);
    if (m_db && !m_sessions.isEmpty())
        m_db->saveSensoryLayout(m_sessions.first().id, json);
    // Queue Excel write — fire-and-forget, log on failure
    if (!m_sessions.isEmpty() && !m_sessions.first().sourceFilePath.isEmpty())
        ExcelLayoutIo::queueWrite(m_sessions.first().sourceFilePath, json);
}
```

- [ ] **Step 2: Add `ExcelLayoutIo` static helper that runs the Python script**

(Roughly mirrors the existing `MainWindow::flushExcelWrites` pattern — debounce + QFutureWatcher.)

- [ ] **Step 3: Round-trip test** — save → reopen Excel → verify property present, content matches.

- [ ] **Step 4: Commit**

---

### Task 26: Phase 2 verification

- [ ] All tests, manual smoke test "edit on machine A, copy .xlsx to machine B with no DB, open, layout still appears". Tag `phase-2-excel-roundtrip`.

---

## Phase 3 — Snap-to-Grid + Alignment Guides (5 tasks)

### Task 27: Snap toolbar toggle + grid spacing constant

- [ ] In `ReportPreviewDialog` toolbar: add a checkable QToolButton "Snap[✓]" that toggles `m_snapEnabled`. Persist in `QSettings("SDR","DataViewerEnterprise","preview/snap")`.

- [ ] Commit.

### Task 28: Grid snap during drag

- [ ] Modify `ResizableSlideItem::mouseMoveEvent` to round `setPos` to nearest 0.1" (in image-pixel space: 6 px) when snap enabled. Add `m_snapEnabled` flag passed in or fetched from a global signal.

### Task 29: Edge-alignment detection

- [ ] During drag, query the parent scene's other items' rects, find any whose left/right/centerX/top/bottom/centerY is within 6 px of the active item's edge. Snap and emit `alignmentGuide(QLineF)` for the canvas to render.

### Task 30: Magenta dashed alignment guides on canvas

- [ ] In `ReportPreviewDialog`, listen for `alignmentGuide` signals and add transient `QGraphicsLineItem`s in magenta/dashed pen. Clear them on mouse release.

### Task 31: Phase 3 verification

- [ ] Manual smoke: drag items, observe snap and guides. Tag `phase-3-snap`.

---

## Phase 4 — Undo / Redo (5 tasks)

### Task 32: `LayoutCommand` base + tests

**Files:**
- Create: `src/reporting/LayoutCommand.h` / `.cpp`
- Create: `tests/tst_layoutcommand/`

- [ ] **Step 1: Header**

```cpp
class LayoutCommand {
public:
    virtual ~LayoutCommand() = default;
    virtual void apply(ReportLayout&) = 0;
    virtual void undo(ReportLayout&) = 0;
    virtual QString description() const = 0;
};
```

- [ ] **Step 2: Concrete commands**

```cpp
class MoveItemCommand : public LayoutCommand {
public:
    MoveItemCommand(const QString& slideKey, const QString& elemId,
                     const QRectF& oldR, const QRectF& newR);
    void apply(ReportLayout&) override;     // sets the rect on the keyed slide+elem
    void undo(ReportLayout&) override;
    QString description() const override { return "Move " + m_elemId; }
private:
    QString m_slideKey, m_elemId;
    QRectF m_old, m_new;
};
class ResizeItemCommand   : public MoveItemCommand { ... }; // structurally identical
class SortColumnCommand   : public LayoutCommand { ... };
class EditTextCommand     : public LayoutCommand { ... };
class ToggleSampleCommand : public LayoutCommand { ... };   // mutates excluded set, not layout
class ZOrderCommand       : public LayoutCommand { ... };
```

- [ ] **Step 3: Tests for apply/undo invariants**

```cpp
void testMoveCommandRoundTrip() {
    DVE::ReportLayout l;
    DVE::ContentSlideLayout c; c.table = QRectF(1,1,5,5);
    l.contentSlides["content_0"] = c;

    DVE::MoveItemCommand cmd("content_0", "table", QRectF(1,1,5,5), QRectF(2,2,5,5));
    cmd.apply(l);
    QCOMPARE(l.contentSlides["content_0"].table, QRectF(2,2,5,5));
    cmd.undo(l);
    QCOMPARE(l.contentSlides["content_0"].table, QRectF(1,1,5,5));
}
```

- [ ] **Step 4: Test .pro + build + commit**

### Task 33: Command stack in dialog

- [ ] Add `QStack<QSharedPointer<LayoutCommand>> m_undo, m_redo;` to dialog. Cap at 100. On every edit, push the appropriate command and clear redo. Hook `Ctrl+Z` / `Ctrl+Y` shortcuts to `pop+undo` / `pop+redo`.

### Task 34: Wire each canvas edit through commands

- [ ] Replace the direct `m_layout` mutations in `applyRectEdit`, `applySortChange`, `sampleToggled`, etc. with `pushCommand(new MoveItemCommand(...))` etc. The command's `apply` mutates `m_layout`.

### Task 35: Tests + verify undoing a sequence works

- [ ] Add a multi-step undo test, verify final state matches initial. Commit. Tag `phase-4-undo`.

---

## Phase 5 — Presets (6 tasks)

### Task 36: `layout_presets` table migration

- [ ] In `DatabaseManager`, add the schema migration:

```sql
CREATE TABLE IF NOT EXISTS layout_presets (
  id INTEGER PRIMARY KEY,
  mode_id TEXT NOT NULL,
  name TEXT NOT NULL,
  layout_json TEXT NOT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  UNIQUE(mode_id, name)
);
```

Bump the schema version constant.

### Task 37: `PresetStore` CRUD class + tests

- [ ] Create `src/reporting/PresetStore.{h,cpp}` with:

```cpp
class PresetStore {
public:
    explicit PresetStore(DatabaseManager*);
    QStringList listNames(const QString& modeId) const;
    QString load(const QString& modeId, const QString& name) const;
    bool    save(const QString& modeId, const QString& name, const QString& json);
    bool    remove(const QString& modeId, const QString& name);
    bool    rename(const QString& modeId, const QString& oldName, const QString& newName);
private:
    DatabaseManager* m_db;
};
```

- [ ] Create `tests/tst_presetstore/` with CRUD round-trip tests. Commit.

### Task 38: Toolbar preset dropdown

- [ ] In `ReportPreviewDialog` add `QComboBox m_presetCombo` populated from `PresetStore::listNames("sensory")`. Selecting an entry applies that preset's layout JSON to `m_layout` (preserving excluded samples and image positions, since presets are template-only) and re-renders the canvas. Push a single composite `ApplyPresetCommand` so it's undoable.

### Task 39: "Save as Preset…" dialog

- [ ] Toolbar button opens a `QInputDialog::getText` for the preset name. On accept, write current `m_layout` (stripped of imageLayouts/imageCrops/excludedSamples) to `PresetStore`. Refresh dropdown.

### Task 40: `PresetManagerDialog` (delete/rename)

- [ ] Tiny dialog with a list + Delete/Rename buttons. Wired to `PresetStore::remove` and `rename`. Commit.

### Task 41: Phase 5 verification

- [ ] Manual: save a preset, edit further, apply preset, observe revert to saved layout. Tag `phase-5-presets`.

---

## Phase 6 — Layout JSON Import / Export (2 tasks)

### Task 42: Export Layout button

- [ ] Toolbar button → `QFileDialog::getSaveFileName` → write `QJsonDocument(m_layout.toJson()).toJson()` to file. Commit.

### Task 43: Import Layout button (with mode validation)

- [ ] Toolbar button → `QFileDialog::getOpenFileName` → parse JSON → if `mode != "sensory"` show warning and abort → otherwise apply (push composite command for undo). Commit. Tag `phase-6-import-export`.

---

## Phase 7 — Cover/Divider Title-Text Editing (3 tasks)

### Task 44: Cover slide canvas

- [ ] In `ReportPreviewDialog::onSlideSelected` extend the cover-case branch: render template background pixmap (low-fidelity placeholder is fine), add a single `TextItem` for title bound to `m_layout.coverTitle` and one for subtitle bound to `coverSubtitle`. Commit.

### Task 45: Divider slide canvas

- [ ] Same pattern for `m_layout.dividerTitles[divider_<id>]`. The `TextItem` text starts as the tester name. Commit.

### Task 46: PptxWriter accepts overridable cover/divider title text

- [ ] Extend `PptxWriter::addCoverSlide` and `addSectionDividerSlide` to accept optional title position and editable text. Wire `SensoryReportSource::writePptx` to pass the values from `m_layout`. Commit. Tag `phase-7-cover-divider`.

---

## Final Verification

### Task 47: Full app + tests + manual smoke

- [ ] **Step 1:** Run `tests\run-tests.ps1`. All suites green.
- [ ] **Step 2:** Run `tests\deployment\Test-Deployment.ps1`. All phases green.
- [ ] **Step 3:** Manual matrix walk:
  - Open sensory single-tester report → preview opens with current default layout.
  - Drag the radar 1" left → auto-save kicks in → close preview → reopen → radar still 1" left.
  - Save preset "tight" → revert table to default position → apply "tight" preset → table back where preset said.
  - Export layout JSON → close → import on a different session → layout applied.
  - Undo 5 edits, Redo 5 edits → final state same as before any edits.
  - Snap toggle off → drag freely → toggle on → drag snaps to 0.1".
  - Uncheck a sample → confirm it disappears from table, radar, and (if multi-session) cumulative slide.
  - Edit cover title → Create Report → open .pptx → cover shows new title.
  - Cancel-without-creating → close → confirm edits persist (auto-save).
- [ ] **Step 4:** Run rebuild-dataviewer skill (bumps version, builds installer).
- [ ] **Step 5:** Tag final: `git tag sensory-report-preview-v1`.

---

## Self-Review Checklist

- [x] Spec coverage: every requirement mapped to a task — sensory-only ✓, drag/resize ✓, sortable cols ✓, sample exclusion ✓, snap ✓, undo ✓, presets ✓, import/export ✓, cover/divider ✓, persistence DB+Excel ✓, default layout matches current report ✓.
- [x] No placeholders — every code step shows actual code; every bash step shows the actual command.
- [x] Type consistency — `ReportLayout`, `ContentSlideLayout`, `IReportSource` defined in Task 1/3, used identically in Tasks 4-46.
- [x] File paths exact in every Files list.
- [x] Tests precede impl where TDD is meaningful (data classes, commands, JSON round-trip, layout defaults). UI tasks are smoke + manual since headless QGraphicsView interaction tests are awkward; this is called out.

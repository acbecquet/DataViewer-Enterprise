# Detailed Sensory Mode — Design Spec

## Overview

A new "Detailed Sensory" mode in DataViewer Enterprise that implements the S2-1 user test template as an in-app survey with dual radar charts, per-sample navigation, database storage, and report generation. Functionally identical to the existing Sensory mode but with a more detailed question set split across two radar chart subcategories.

## Template Source

Based on: `S2-1 user test for external use.xlsx` ("Sensory Forms Simplified" sheet)

---

## 1. Data Model

### New file: `src/pipeline/DetailedSensoryData.h`

**`DetailedSensorySample` struct:**
- `name` — sample identifier
- **Vapor Quality scores (stored at original scale):**
  - `burnTaste` (1-9) — label: "Burn Taste (1=none, 9=too much)"
  - `flavorIntensity` (1-9) — label: "Flavor Intensity (5=ideal, 1-9)"
  - `throatIrritation` (1-9) — label: "Throat Irritation (1=none, 9=very bad)"
  - `nasalIrritation` (1-9) — label: "Nasal Irritation (1=none, 9=worst)"
  - `vaporQualityOverall` (1-9) — label: "Vapor Quality Overall (9=best)"
  - `cough` (1-4) — label: "Cough (1=none, 4=worst)"
- **Consistency scores:**
  - `vaporVolumeConsistency` (1-4) — label: "Volume Consistency (1=best, 4=worst)"
  - `performanceConsistency` (1-3) — label: "Performance Consistency (1=best, 3=worst)"
  - `vaporTemperature` (1-4) — label: "Vapor Temperature (1=good, 4=worst)"
  - `vaporVsOriginalOil` (1-4) — label: "Vapor vs Oil (1=true, 4=major issue)"
  - `vaporVolumeOverall` (1-5) — label: "Vapor Volume (3=ok, 1-5)"
- **Properties:** voltage, resistance, power (calculated), heatingTechnology, media, viscosity
- **Extra fields:** oilSmellLiking (1-5), clog (bool + oil level note), comments, deviceReturnDate

**Normalization helper:** `normalizedScore(metric) -> double` maps any score to 1-9 for radar display:
- 1-9 scores: pass through unchanged
- 1-4 scores: linear map (1->1, 2->3.67, 3->6.33, 4->9)
- 1-3 scores: linear map (1->1, 2->5, 3->9)
- 1-5 scores: linear map (1->1, 2->3, 3->5, 4->7, 5->9)
- No inversion — raw values displayed, scale direction conveyed via axis labels

**`DetailedSensorySession` struct:**
- sessionName, testTitle, assessorName, testerName, facilitatorName, facilitatorComment
- media, date, timestamp (ISO8601)
- `QVector<DetailedSensorySample> samples`
- imagePaths, imageLayouts, imageCrops (same pattern as existing SensorySession)

### Radar Chart Subcategories

**Vapor Quality (6 axes):**
| Metric | Original Scale | Radar Label |
|--------|---------------|-------------|
| Burn Taste | 1-9 | Burn Taste (1=none, 9=too much) |
| Flavor Intensity | 1-9 | Flavor Intensity (5=ideal, 1-9) |
| Throat Irritation | 1-9 | Throat Irritation (1=none, 9=very bad) |
| Nasal Irritation | 1-9 | Nasal Irritation (1=none, 9=worst) |
| Vapor Quality Overall | 1-9 | Vapor Quality Overall (9=best) |
| Cough | 1-4 | Cough (1=none, 4=worst) |

**Consistency (5 axes):**
| Metric | Original Scale | Radar Label |
|--------|---------------|-------------|
| Vapor Volume Consistency | 1-4 | Volume Consistency (1=best, 4=worst) |
| Performance Consistency | 1-3 | Performance Consistency (1=best, 3=worst) |
| Vapor Temperature | 1-4 | Vapor Temperature (1=good, 4=worst) |
| Vapor vs Original Oil | 1-4 | Vapor vs Oil (1=true, 4=major issue) |
| Vapor Volume Overall | 1-5 | Vapor Volume (3=ok, 1-5) |

### Multiple-Choice Options Reference

**Cough (1-4):**
1. No
2. Yes, but it doesn't bother me at all.
3. Yes, and I will avoid buying this product the next time.
4. Yes, I will stop using it immediately.

**Vapor vs Original Oil (1-4):**
1. The vapor is very much true to the original oil
2. The vapor shows the main features of the oil but not everything
3. The vapor has some minor off taste/flavor issue
4. The vapor has major off taste/flavor issue

**Vapor Volume Consistency (1-4):**
1. Yes, very consistent.
2. Mostly yes, it increased a little bit from puff to puff but was acceptable.
3. No, the vapor volume increased obviously.
4. No, the vapor volume jumped everywhere (higher and lower).

**Vapor Volume Overall (1-5):**
1. Too little
2. A bit little but I can live with that
3. Ok
4. Big
5. Very big

**Vapor Temperature (1-4):**
1. Always within the good range
2. Ok at the beginning, got uncomfortable hot after a couple of puffs
3. Always too hot
4. Too cold, I wish it could have been a little bit warmer than this

**Performance Consistency (1-3):**
1. Yes, very consistent.
2. Mostly yes. Some small performance change doesn't bother me.
3. No, the device is not stable.

---

## 2. UI Layout

### Mode Switching

**Tools tab:** New "Detailed Sensory" checkable button next to existing "Sensory" button.
- Black version of `resources/images/ccell_icon.png`
- Three-way mutually exclusive toggle: TPM / Sensory / Detailed Sensory
- Checking one unchecks the other

### Left Dock (identical to existing Sensory mode)

- **Navigator** — session list (top). Single-select shows one session. Multi-select (Ctrl+click) shows averaged chart.
- **Test Averages** — averaged scores display
- **Sample Properties** — Session Info (Test Title, Assessor, Tester, Media, Date, Samples), Test Properties, Computed section
  - Additional properties: Oil Smell Liking (1-5), Clog (yes/no + oil level note), Device Return Date, Facilitator Comment, Facilitator Name, Other Comments
- **Load Images / View Images** buttons at bottom

### Center Area (main panel)

**Top header row:** Test Title, Assessor, Tester, Media, Date fields (identical to current Sensory)

**Top strip — Data entry area:**
- Sample switcher: prev/next arrows in same location as TPM mode (Ctrl+Left/Right shortcuts)
- One sample visible at a time
- Compact grid/table layout: label on left, input on right
  - Scored questions (1-9): `NoWheelDoubleSpinBox` with appropriate range, 0.1 step
  - Multiple-choice questions: `QComboBox` dropdown showing full text options, storing numeric value
  - Each label includes parenthetical scale note
- Add Sample / Remove buttons

**Bottom half — Dual radar charts side by side:**
- Left: Vapor Quality radar (6 axes)
- Right: Consistency radar (5 axes)
- Both show all samples overlaid with clickable legend toggles
- Legend at bottom of each chart

### Panel Sizing

All panel widths, dock positions, and proportions stay identical across mode transitions (TPM / Sensory / Detailed Sensory). Transition is seamless.

---

## 3. Database Storage

### New table: `detailed_sensory_sessions`

```sql
detailed_sensory_sessions (
  id INTEGER PRIMARY KEY,
  session_name TEXT,
  tester_name TEXT,
  assessor_name TEXT,
  media TEXT,
  date TEXT,
  timestamp TEXT,
  json_data TEXT
)
```

- JSON blob stores full `DetailedSensorySession` with all question scores, properties, and extra fields
- Upsert matching: `session_name + tester_name + date`

### New table: `detailed_sensory_images`

Same schema as `sensory_images`, foreign key to `detailed_sensory_sessions.id`.

### Database Browser

New "Detailed Sensory" tab alongside existing "Files" and "Sensory" tabs.

### DatabaseManager Methods

- `saveDetailedSensorySession(const DetailedSensorySession&)`
- `loadDetailedSensorySessions() const -> QVector<DetailedSensorySession>`
- `loadDetailedSensorySession(int id) const -> DetailedSensorySession`
- `listDetailedSensoryRecords() const` — summary records for browser
- `removeDetailedSensorySession(int id)`

---

## 4. Reports & Export

### Excel Export (flat table)

- Single sheet
- **Columns:** Sample Name | Burn Taste | Flavor Intensity | Throat Irritation | Nasal Irritation | Vapor Quality Overall | Cough | Vapor Volume Consistency | Performance Consistency | Vapor Temperature | Vapor vs Original Oil | Vapor Volume Overall | V | R | P | HT | Media | Viscosity | Oil Smell Liking | Clog | Comments
- **Rows:** one per sample
- **Bottom row:** Averages (labeled "Average")
- **Footer rows:** session metadata (Test Title, Tester, Assessor, Date, Facilitator, etc.)

### PowerPoint Report

Uses same template as existing sensory report.

**Slide structure:**
1. **Title slide** — same formatting as current sensory report
2. **Per sample, repeating:**
   - **Table slide** — full page. Questions as columns, testers as rows. "Average" row at bottom, labeled. Selected testers only.
   - **Plot slide** — two radar charts side by side. Vapor Quality (left), Consistency (right). All selected testers overlaid for that sample. Axis labels include scale notes in parentheses.

### Tester Selection

Same tester selection dialog as existing sensory report — user picks which sessions/testers to include before generating. Selected testers become rows in the table and overlays on the charts.

---

## 5. Ribbon Wiring

### Home Tab (Detailed Sensory mode)

- **New Session** -> `DetailedSensoryPanel::newSession()`
- **Save** -> `DetailedSensoryPanel::save()`
- **Load Excel** -> `DetailedSensoryPanel::loadFiles()`
- **Close** -> `DetailedSensoryPanel::close()`

Three-way signal connect/disconnect in `updateRibbonForMode()` — TPM, Sensory, Detailed Sensory.

### Reports Tab (Detailed Sensory mode)

- "Detailed Sensory Report" button with tester selection dialog
- Hides "Full Report" (TPM only)

---

## 6. Architecture Approach

**Parallel Panel Architecture (Approach A):**
- New `DetailedSensoryPanel` class as sibling to `SensoryPanel`
- Own data structures in `DetailedSensoryData.h`
- Extends `RadarChartWidget` to support custom axis definitions and label annotations (dual-chart mode)
- Separate database table, clean separation from existing sensory mode
- No risk of regressions to existing sensory functionality

### New Files

| File | Purpose |
|------|---------|
| `src/pipeline/DetailedSensoryData.h` | Data structs + normalization |
| `src/ui/DetailedSensoryPanel.h` | Panel header |
| `src/ui/DetailedSensoryPanel.cpp` | Panel implementation |
| `resources/images/ccell_icon_black.png` | Black variant of ccell icon |

### Modified Files

| File | Changes |
|------|---------|
| `src/MainWindow.h/cpp` | Three-way mode toggle, ribbon wiring, DetailedSensoryPanel init |
| `src/ui/RadarChartWidget.h/cpp` | Custom axis definitions, label annotations, configurable axis count |
| `src/database/DatabaseManager.h/cpp` | New table + CRUD methods |
| `src/ui/DatabaseBrowserDialog.h/cpp` | New "Detailed Sensory" tab |
| `DataViewerEnterprise.pro` | Add new source files |

### Shared Infrastructure (no duplication)

- `RadarChartWidget` — extended, not duplicated
- PowerPoint template + generation infrastructure
- Image loading/viewing system
- `NoWheelDoubleSpinBox` widget
- Database connection management

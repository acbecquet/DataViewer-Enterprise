# Legacy TPM Template Backwards Compatibility — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enable DataViewer Enterprise to load TPM Excel files from 4 legacy template formats (A through D), normalizing them into the current Format E (Dec 2025) column layout so the existing processing pipeline works unchanged.

**Architecture:** Add a normalization layer in `ExcelReader` that detects which legacy format a sheet uses (by inspecting header strings in row 4 and metadata labels in rows 1-3), then rewrites the in-memory `SheetData::cells` to match the Format E 12-column layout before any metadata extraction or sample parsing occurs. This keeps all changes localized to `ExcelReader` — `SheetProcessors`, `DataProcessor`, and all downstream code remain untouched.

**Tech Stack:** C++ / Qt 6, QVariant 2D arrays, existing ExcelReader/SheetData infrastructure.

---

## File Structure

| File | Action | Responsibility |
|------|--------|---------------|
| `src/ExcelReader.h` | Modify | Add `LegacyFormat` enum, `detectLegacyFormat()`, `normalizeSheet()` declarations |
| `src/ExcelReader.cpp` | Modify | Implement format detection + normalization; update `detectTemplateVersion()`, `countSamples()`, `extractMetadata()` |

No new files are created. All changes are in the ExcelReader translation layer — the pipeline sees normalized Format E data regardless of source format.

---

## Format Reference (for implementer context)

All formats must be normalized to **Format E** (12 cols/sample):

```
Col 0: puffs
Col 1: Before weight (g)
Col 2: After weight (g)
Col 3: Draw Pressure (kpa)
Col 4: Resistance (Ω)
Col 5: Smell
Col 6: Clog
Col 7: Notes
Col 8: TPM (mg/puff)
Col 9: TPM Power Density (mg/(W*s))
Col 10: Variation in TPM (%)
Col 11: Oil Consumed (Cumulative, g)
```

Format E metadata (rows 0-2, relative to sample offset):
```
Row 0: [0]=TestName  [2]="Date:" [3]=date  [4]="Sample ID:" [5]=sampleID  [6]="Heating Technology:" [7]=heatingTech
Row 1: [0]="Media:" [1]=media  [2]="Resistance (Ω):" [3]=resistance  [4]="Power:" [5]=power  [6]="Puffing Regime:" [7]=regime
Row 2: [0]="Viscosity:" [1]=viscosity  [2]="Tester:" [3]=tester  [4]="Voltage:" [5]=voltage  [6]="Initial Oil Mass:" [7]=oilMass
```

### Format A — "Comparison Test" (13 cols/sample)
- **Detection**: Row 3 (header) col 3 = "PV1" AND col 10 = "Clog?" (13 cols)
- **Data cols**: puffs | Before weight/g | After weight/g | PV1 | PV2 | PV3 | PV4 | PV5 | Resistance | Smell (0-4) | Clog? | Notes | TPM (mg/puff)
- **Metadata**: Row 0: A="Date:" | Row 1: A="Cart #" B=sampleID, C="Ri (Ohms)" D=resistance, E="Power", G="Viscosity" H=viscosity, I="Did this clog?", K="Initial Oil Mass" L=oilMass | Row 2: A="Media" B=media, E="Puff Regime", G="Voltage" H=voltage, K="Usage Efficiency"
- **Column mapping**: PV1→Draw Pressure, PV2-PV5→discard, Clog?→Clog, remaining cols shift

### Format B — "M1 Extended" (12 cols/sample, row 0 empty)
- **Detection**: Row 0 is entirely empty AND Row 3 (header) col 3 = "PV1" AND col 10 = "Notes" (no Clog column, 12 cols)
- **Data cols**: puffs | Before weight/g | After weight/g | PV1 | PV2 | PV3 | PV4 | PV5 | Resistance | Smell (0-4) | Notes | TPM (mg/puff)
- **Metadata**: Row 0: empty. Row 1: A="Cart #" B=sampleID, C="Ri (Ohms)" D=resistance, E="Power", G="Viscosity" H=viscosity, I="Temperature". Row 2: A="Media" B=media, E="Puff Regime", G="Voltage" H=voltage.
- **Column mapping**: PV1→Draw Pressure, PV2-PV5→discard, insert empty Clog column, remap remaining

### Format C — "Gembox/T58G" (12 cols/sample)
- **Detection**: Row 0 col 0 starts with "TPM" AND Row 1 col 2 = "Ri (Ohms)"
- **Data cols**: Already 12 cols matching Format E data positions: puffs | Before weight/g | After weight/g | Draw Pressure | Resistance | Smell | Clog | Notes | TPM | TPM Power Density | TPM Consistency | Rolling Average TPM
- **Metadata remapping needed**: Row 0: A="TPM" (testName), B="Date:" C=date, D="Tester:" E=tester, F="Project:" G=project, H="Sample:" I=sampleID → combine G+I for sampleID. Row 1: A="Media:" B=media, C="Ri (Ohms)" D=resistance, E="Power:" F=power, G="Puffing Regime:" H=regime. Row 2: A="Viscosity:" B=viscosity, E="Voltage:" F=voltage, G="Initial Oil Mass:" H=oilMass.

### Format D — "Jan 2025 Standard" (11-12 cols/sample)
- **Detection**: Row 1 col 2 contains "Resistance (Ohms)" AND Row 0 col 0 is NOT "TPM" AND row 0 is NOT empty
- **Data cols**: Same 12-col layout as Format E for most sheets. Long Puff Test variant has 11 cols (missing Rolling Average TPM / Oil Consumed).
- **Metadata**: Nearly identical to Format E except missing Heating Technology at row 0 col 6/7. Uses "Resistance (Ohms):" label.
- **Normalization**: For 11-col sheets, pad to 12 cols. No other data remapping needed. Metadata is already in the right positions (just no heating tech).

### Format E — "Dec 2025 New" (already supported, no changes needed)
- **Detection**: Row 1 col 2 contains "Resistance (Ω)" or sheet names contain "Long Puff Lifetime Test" etc.

---

## Task 1: Add LegacyFormat enum and detectLegacyFormat()

**Files:**
- Modify: `src/ExcelReader.h:19-107`
- Modify: `src/ExcelReader.cpp:312-350`

- [ ] **Step 1: Add the LegacyFormat enum to ExcelReader.h**

In `ExcelReader.h`, add the enum inside the class (after the `SampleData` struct, before the public constructor):

```cpp
    // Legacy template format identifiers.
    // normalizeSheet() rewrites in-memory cells to Format E before extraction.
    enum class LegacyFormat {
        FormatE,       // Dec 2025 "new" — already native, no transform
        FormatD,       // Jan 2025 "old" — minor metadata gaps, 11-col variant
        FormatC,       // Gembox/T58G — different metadata layout
        FormatB,       // M1 Extended — PV1-5, no Clog, row 0 empty
        FormatA,       // Comparison Test — PV1-5, 13 cols, Clog?
        Unknown        // unrecognised — pass through unchanged
    };
```

Also add the private method declarations at the bottom of the private section:

```cpp
    LegacyFormat detectLegacyFormat()      const;
    void         normalizeSheet(LegacyFormat fmt);
```

- [ ] **Step 2: Implement detectLegacyFormat() in ExcelReader.cpp**

Add the following method after the existing `detectTemplateVersion()` function (after line 350):

```cpp
ExcelReader::LegacyFormat ExcelReader::detectLegacyFormat() const
{
    const SheetData* sd = currentSheetData();
    if (!sd || sd->cells.size() < 4)
        return LegacyFormat::Unknown;

    // Helper: get a trimmed string from a cell
    auto cell = [&](int row, int col) -> QString {
        if (row < 0 || row >= sd->cells.size()) return {};
        const auto& r = sd->cells[row];
        if (col < 0 || col >= r.size()) return {};
        const QVariant& v = r[col];
        return v.isNull() ? QString() : v.toString().trimmed();
    };

    // Check row 3 (header row, 0-based) for PV1 — indicates Format A or B
    QString col3Header = cell(3, 3);  // col D of first sample
    bool hasPV1 = col3Header.compare("PV1", Qt::CaseInsensitive) == 0;

    if (hasPV1) {
        // Distinguish A (13 cols, has Clog?) from B (12 cols, no Clog)
        // Format A: col 10 = "Clog?" or "Clog"
        QString col10Header = cell(3, 10);
        if (col10Header.contains("Clog", Qt::CaseInsensitive)) {
            debugPrint("Detected legacy Format A (Comparison Test, 13 cols)");
            return LegacyFormat::FormatA;
        }
        // Format B: col 10 = "Notes" (no Clog column)
        debugPrint("Detected legacy Format B (M1 Extended, 12 cols, PV1-5)");
        return LegacyFormat::FormatB;
    }

    // Check for Format C: Row 0 col 0 starts with "TPM" and Row 1 col 2 = "Ri (Ohms)"
    QString r0c0 = cell(0, 0);
    QString r1c2 = cell(1, 2);
    if (r0c0.startsWith("TPM", Qt::CaseInsensitive) &&
        r1c2.compare("Ri (Ohms)", Qt::CaseInsensitive) == 0) {
        debugPrint("Detected legacy Format C (Gembox/T58G)");
        return LegacyFormat::FormatC;
    }

    // Check for Format D: Row 1 col 2 contains "Resistance (Ohms)"
    // (Format E uses "Resistance (Ω)" instead)
    if (r1c2.contains("Resistance (Ohms)", Qt::CaseInsensitive)) {
        debugPrint("Detected legacy Format D (Jan 2025 standard)");
        return LegacyFormat::FormatD;
    }

    // Check for Format E: Row 1 col 2 contains Ω or already detected as "new"
    if (r1c2.contains(QChar(0x03A9)) || r1c2.contains("Resistance", Qt::CaseInsensitive)) {
        debugPrint("Detected Format E (Dec 2025 new) — no normalization needed");
        return LegacyFormat::FormatE;
    }

    debugPrint("Could not detect legacy format — passing through unchanged");
    return LegacyFormat::Unknown;
}
```

- [ ] **Step 3: Update detectTemplateVersion() to account for legacy formats**

The existing `detectTemplateVersion()` returns "new" or "old" and controls metadata extraction paths. After normalization, all legacy formats will look like "new" (Format E). Modify the method to also check if the file was normalized from a legacy format.

In `ExcelReader.cpp`, replace the existing `detectTemplateVersion()` (lines 312-350) with:

```cpp
QString ExcelReader::detectTemplateVersion() const
{
    debugPrint("Detecting template version");
    if (m_sheets.isEmpty()) return QStringLiteral("unknown");

    QStringList newIndicators = {
        "Long Puff lifetime Test",
        "Rapid Puff Lifetime Test",
        "Temperature Cycling Test #1",
        "Temperature Cycling Test #2"
    };
    bool hasAnyData = false;
    for (const auto& s : m_sheets) {
        for (const auto& ind : newIndicators) {
            if (s.name.contains(ind, Qt::CaseInsensitive)) {
                debugPrint("Detected new template (December 2025)");
                return QStringLiteral("new");
            }
        }
        if (!hasAnyData) {
            for (int r = 0; r < qMin(5, (int)s.cells.size()); ++r) {
                for (const QVariant& cell : s.cells[r]) {
                    if (cell.isValid() && !cell.isNull() && !cell.toString().trimmed().isEmpty()) {
                        hasAnyData = true;
                        break;
                    }
                }
                if (hasAnyData) break;
            }
        }
    }
    if (!hasAnyData) {
        debugPrint("Sheets exist but contain no data — empty template");
        return QStringLiteral("unknown");
    }

    // If any sheet has been normalized from a legacy format, treat the file
    // as "new" so extractMetadata uses the Format E column positions.
    const SheetData* sd = currentSheetData();
    if (sd && !sd->cells.isEmpty()) {
        // After normalization, row 1 col 2 will contain "Resistance (Ω):"
        if (sd->cells.size() > 1 && sd->cells[1].size() > 2) {
            QString r1c2 = sd->cells[1][2].toString().trimmed();
            if (r1c2.contains(QChar(0x03A9))) {  // Ω
                debugPrint("Sheet has been normalized to Format E — treating as 'new'");
                return QStringLiteral("new");
            }
        }
    }

    debugPrint("Detected old template (January 2025)");
    return QStringLiteral("old");
}
```

- [ ] **Step 4: Commit**

```bash
git add src/ExcelReader.h src/ExcelReader.cpp
git commit -m "feat: add legacy TPM format detection (Formats A-D)"
```

---

## Task 2: Implement normalizeSheet() for Format A (Comparison Test, 13 cols)

**Files:**
- Modify: `src/ExcelReader.cpp`

Format A has 13 cols per sample:
```
Old:  puffs | Before weight/g | After weight/g | PV1 | PV2 | PV3 | PV4 | PV5 | Resistance | Smell (0-4) | Clog? | Notes | TPM (mg/puff)
Index:  0         1                 2             3     4     5     6     7       8            9             10      11      12
```
Target (Format E, 12 cols):
```
New:  puffs | Before weight (g) | After weight (g) | Draw Pressure | Resistance (Ω) | Smell | Clog | Notes | TPM | TPM Power Density | Variation | Oil Consumed
Index:  0         1                   2                 3               4               5       6      7      8      9                  10          11
```
Mapping: old[0]→new[0], old[1]→new[1], old[2]→new[2], old[3](PV1)→new[3](Draw Pressure), old[4-7](PV2-5)→discard, old[8](Resistance)→new[4], old[9](Smell)→new[5], old[10](Clog?)→new[6], old[11](Notes)→new[7], old[12](TPM)→new[8], new[9-11]→empty (will be recalculated).

Metadata remapping (per sample, from 13-col positions to 12-col positions):
```
Row 0: old positions → new positions
  A0="Date:"      → keep but need to restructure entirely
  B1(Cart#)→sampleID  needs to go to row0[5]
  B2(Media)→media      needs to go to row1[1]
  ...
```

Full metadata rewrite for each sample block:
```
OLD Row 0: [0]=Date:    [2]=Coil Material [4]=Thermal Conductivity [6]=Column inner diameter [8]=Total Oil Puffed [10]=Did this burn? [12]=Average TPM
OLD Row 1: [0]=Cart#    [1]=sampleID      [2]=Ri(Ohms) [3]=resistance [4]=Power [6]=Viscosity [7]=viscosity [8]=Did this clog? [10]=Initial Oil Mass [11]=oilMass
OLD Row 2: [0]=Media    [1]=media         [2]=Rf(Ohms)               [4]=Puff Regime         [6]=Voltage [7]=voltage [8]=Did this leak? [10]=Usage Efficiency

NEW Row 0: [0]=testName [2]=Date: [3]=date [4]=Sample ID: [5]=sampleID [6]=Heating Technology: [7]=(empty)
NEW Row 1: [0]=Media: [1]=media [2]=Resistance (Ω): [3]=resistance [4]=Power: [5]=power [6]=Puffing Regime: [7]=regime
NEW Row 2: [0]=Viscosity: [1]=viscosity [2]=Tester: [3]=(empty) [4]=Voltage: [5]=voltage [6]=Initial Oil Mass: [7]=oilMass
```

- [ ] **Step 1: Write normalizeSheetFormatA() helper**

Add this private method to `ExcelReader.cpp` (before `normalizeSheet()`):

```cpp
// ---------------------------------------------------------------------------
// normalizeSheetFormatA — Comparison Test, 13 cols/sample → 12 cols/sample
// ---------------------------------------------------------------------------
static void normalizeSheetFormatA(QVector<QVector<QVariant>>& cells)
{
    if (cells.size() < 4) return;

    const int OLD_COLS = 13;
    const int NEW_COLS = 12;

    // Determine number of samples from header row width
    int totalOldCols = cells[3].size();
    int nSamples = totalOldCols / OLD_COLS;
    if (nSamples <= 0) return;

    int totalNewCols = nSamples * NEW_COLS;

    // --- Remap data rows (row 4+) first: 13→12 per sample ---
    // Mapping per sample: old→new
    //   old[0]→new[0]  puffs
    //   old[1]→new[1]  Before weight
    //   old[2]→new[2]  After weight
    //   old[3]→new[3]  PV1 → Draw Pressure
    //   old[4..7]      PV2-PV5 → discard
    //   old[8]→new[4]  Resistance
    //   old[9]→new[5]  Smell
    //   old[10]→new[6] Clog?
    //   old[11]→new[7] Notes
    //   old[12]→new[8] TPM
    //   new[9..11]     empty (recalculated by pipeline)
    static const int dataMap[] = { 0, 1, 2, 3, 8, 9, 10, 11, 12 }; // old indices → new[0..8]
    static const int dataMapLen = 9; // fills new[0] through new[8]

    for (int row = 4; row < cells.size(); ++row) {
        QVector<QVariant> oldRow = cells[row];
        // Pad oldRow if needed
        while (oldRow.size() < nSamples * OLD_COLS)
            oldRow.append(QVariant());

        QVector<QVariant> newRow(totalNewCols);
        for (int s = 0; s < nSamples; ++s) {
            int oldOff = s * OLD_COLS;
            int newOff = s * NEW_COLS;
            for (int i = 0; i < dataMapLen; ++i)
                newRow[newOff + i] = oldRow[oldOff + dataMap[i]];
            // new[9..11] remain default QVariant() (null) — will be recalculated
        }
        cells[row] = newRow;
    }

    // --- Rewrite header row (row 3) ---
    {
        QVector<QVariant> newHeaders(totalNewCols);
        QStringList hdr = { "puffs", "Before weight (g)", "After weight (g)",
                            "Draw Pressure (kpa)", QString("Resistance (%1)").arg(QChar(0x03A9)),
                            "Smell", "Clog", "Notes", "TPM (mg/puff)",
                            "TPM Power Density (mg/(W*s))", "Variation in TPM (%)",
                            "Oil Consumed (Cumulative, g)" };
        for (int s = 0; s < nSamples; ++s) {
            int off = s * NEW_COLS;
            for (int c = 0; c < NEW_COLS; ++c)
                newHeaders[off + c] = hdr[c];
        }
        cells[3] = newHeaders;
    }

    // --- Rewrite metadata rows (0-2) per sample ---
    // Pad rows 0-2 to old width if needed
    for (int r = 0; r < 3; ++r) {
        while (cells[r].size() < nSamples * OLD_COLS)
            cells[r].append(QVariant());
    }

    QVector<QVariant> newRow0(totalNewCols);
    QVector<QVariant> newRow1(totalNewCols);
    QVector<QVariant> newRow2(totalNewCols);

    for (int s = 0; s < nSamples; ++s) {
        int oldOff = s * OLD_COLS;
        int newOff = s * NEW_COLS;

        // Extract old metadata values
        QVariant date      = cells[0][oldOff + 1];   // B1 (value next to "Date:" label at A1)
        // Date label is at old[0], value could be at old[1] or sometimes empty
        // In Comparison Test, Date: is at col A with value at col B (but often empty)
        QVariant sampleID  = cells[1][oldOff + 1];   // B2 (Cart # value)
        QVariant resistance= cells[1][oldOff + 3];   // D2 (Ri value)
        QVariant viscosity = cells[1][oldOff + 7];   // H2 (Viscosity value)
        QVariant oilMass   = cells[1][oldOff + 11];  // L2 (Initial Oil Mass value)
        QVariant media     = cells[2][oldOff + 1];   // B3 (Media value)
        QVariant puffRegime= cells[2][oldOff + 5];   // F3 (Puff Regime value)
        QVariant voltage   = cells[2][oldOff + 7];   // H3 (Voltage value)

        // Burn/Clog/Leak from old metadata
        QVariant burnStatus = cells[0][oldOff + 11];  // L1 (Did this burn? answer)
        QVariant clogStatus = cells[1][oldOff + 9];   // J2 (Did this clog? answer)
        QVariant leakStatus = cells[2][oldOff + 9];   // J3 (Did this leak? answer)

        // Build test name from "Comparison Test" or sheet context
        QString testName = QStringLiteral("Comparison Test");

        // Write Format E metadata
        // Row 0: testName, _, "Date:", date, "Sample ID:", sampleID, "Heating Technology:", _
        newRow0[newOff + 0] = testName;
        newRow0[newOff + 2] = QStringLiteral("Date:");
        newRow0[newOff + 3] = date;
        newRow0[newOff + 4] = QStringLiteral("Sample ID:");
        newRow0[newOff + 5] = sampleID;
        newRow0[newOff + 6] = QStringLiteral("Heating Technology:");
        // newRow0[newOff + 7] stays empty — no heating tech in Format A
        // Burn status: store at [9] for detection by SheetProcessors
        newRow0[newOff + 9] = QStringLiteral("Did this burn?");
        newRow0[newOff + 10] = burnStatus;

        // Row 1: "Media:", media, "Resistance (Ω):", resistance, "Power:", _, "Puffing Regime:", regime
        newRow1[newOff + 0] = QStringLiteral("Media:");
        newRow1[newOff + 1] = media;
        newRow1[newOff + 2] = QString("Resistance (%1):").arg(QChar(0x03A9));
        newRow1[newOff + 3] = resistance;
        newRow1[newOff + 4] = QStringLiteral("Power:");
        // newRow1[newOff + 5] = power formula — leave empty, extractMetadata calculates it
        newRow1[newOff + 6] = QStringLiteral("Puffing Regime:");
        newRow1[newOff + 7] = puffRegime;
        newRow1[newOff + 9] = QStringLiteral("Did this clog?");
        newRow1[newOff + 10] = clogStatus;

        // Row 2: "Viscosity:", viscosity, "Tester:", _, "Voltage:", voltage, "Initial Oil Mass:", oilMass
        newRow2[newOff + 0] = QStringLiteral("Viscosity:");
        newRow2[newOff + 1] = viscosity;
        newRow2[newOff + 2] = QStringLiteral("Tester:");
        // newRow2[newOff + 3] — no tester in Format A
        newRow2[newOff + 4] = QStringLiteral("Voltage:");
        newRow2[newOff + 5] = voltage;
        newRow2[newOff + 6] = QStringLiteral("Initial Oil Mass:");
        newRow2[newOff + 7] = oilMass;
        newRow2[newOff + 9] = QStringLiteral("Did this leak?");
        newRow2[newOff + 10] = leakStatus;
    }

    cells[0] = newRow0;
    cells[1] = newRow1;
    cells[2] = newRow2;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/ExcelReader.cpp
git commit -m "feat: implement Format A normalization (Comparison Test, 13→12 cols)"
```

---

## Task 3: Implement normalizeSheet() for Format B (M1 Extended, 12 cols, PV1-5)

**Files:**
- Modify: `src/ExcelReader.cpp`

Format B has 12 cols but PV1-5 instead of Draw Pressure, and no Clog column:
```
Old:  puffs | Before weight/g | After weight/g | PV1 | PV2 | PV3 | PV4 | PV5 | Resistance | Smell (0-4) | Notes | TPM
Index:  0         1                 2             3     4     5     6     7       8            9             10      11
```
Target (Format E, 12 cols):
```
New:  puffs | Before weight (g) | After weight (g) | Draw Pressure | Resistance | Smell | Clog | Notes | TPM | TPM PD | Variation | Oil Consumed
Index:  0         1                   2                 3              4           5       6      7      8     9        10          11
```
Mapping: old[0]→new[0], old[1]→new[1], old[2]→new[2], old[3](PV1)→new[3], old[4-7]→discard, old[8]→new[4], old[9]→new[5], (empty)→new[6](Clog), old[10]→new[7](Notes), old[11]→new[8](TPM), new[9-11]→empty.

Row 0 is completely empty. Metadata starts at row 1.

- [ ] **Step 1: Write normalizeSheetFormatB() helper**

```cpp
// ---------------------------------------------------------------------------
// normalizeSheetFormatB — M1 Extended, 12 cols with PV1-5, no Clog, row 0 empty
// ---------------------------------------------------------------------------
static void normalizeSheetFormatB(QVector<QVector<QVariant>>& cells)
{
    if (cells.size() < 4) return;

    const int OLD_COLS = 12;
    const int NEW_COLS = 12;

    int totalCols = cells[3].size();
    int nSamples = totalCols / OLD_COLS;
    if (nSamples <= 0) return;

    int totalNewCols = nSamples * NEW_COLS;

    // --- Remap data rows (row 4+) ---
    // old[0]→new[0], old[1]→new[1], old[2]→new[2], old[3](PV1)→new[3](DrawPressure)
    // old[4..7](PV2-5)→discard
    // old[8]→new[4](Resistance), old[9]→new[5](Smell), (empty)→new[6](Clog)
    // old[10]→new[7](Notes), old[11]→new[8](TPM)
    // new[9..11] = empty
    for (int row = 4; row < cells.size(); ++row) {
        QVector<QVariant> oldRow = cells[row];
        while (oldRow.size() < nSamples * OLD_COLS)
            oldRow.append(QVariant());

        QVector<QVariant> newRow(totalNewCols);
        for (int s = 0; s < nSamples; ++s) {
            int off = s * OLD_COLS;  // same width before and after for B
            newRow[off + 0] = oldRow[off + 0];  // puffs
            newRow[off + 1] = oldRow[off + 1];  // before weight
            newRow[off + 2] = oldRow[off + 2];  // after weight
            newRow[off + 3] = oldRow[off + 3];  // PV1 → Draw Pressure
            newRow[off + 4] = oldRow[off + 8];  // Resistance
            newRow[off + 5] = oldRow[off + 9];  // Smell
            newRow[off + 6] = QVariant();        // Clog (not in Format B)
            newRow[off + 7] = oldRow[off + 10]; // Notes
            newRow[off + 8] = oldRow[off + 11]; // TPM
            // new[9..11] remain null — recalculated
        }
        cells[row] = newRow;
    }

    // --- Rewrite header row (row 3) ---
    {
        QVector<QVariant> newHeaders(totalNewCols);
        QStringList hdr = { "puffs", "Before weight (g)", "After weight (g)",
                            "Draw Pressure (kpa)", QString("Resistance (%1)").arg(QChar(0x03A9)),
                            "Smell", "Clog", "Notes", "TPM (mg/puff)",
                            "TPM Power Density (mg/(W*s))", "Variation in TPM (%)",
                            "Oil Consumed (Cumulative, g)" };
        for (int s = 0; s < nSamples; ++s) {
            int off = s * NEW_COLS;
            for (int c = 0; c < NEW_COLS; ++c)
                newHeaders[off + c] = hdr[c];
        }
        cells[3] = newHeaders;
    }

    // --- Rewrite metadata rows 0-2 ---
    for (int r = 0; r < 3; ++r) {
        while (cells[r].size() < nSamples * OLD_COLS)
            cells[r].append(QVariant());
    }

    QVector<QVariant> newRow0(totalNewCols);
    QVector<QVariant> newRow1(totalNewCols);
    QVector<QVariant> newRow2(totalNewCols);

    for (int s = 0; s < nSamples; ++s) {
        int off = s * OLD_COLS;

        // Old metadata (row 0 is empty, metadata in rows 1-2):
        // Row 1: [0]=Cart# [1]=sampleID [2]=Ri(Ohms) [3]=resistance [4]=Power [6]=Viscosity [7]=viscosity [8]=Temperature
        // Row 2: [0]=Media [1]=media [2]=Rf(Ohms) [4]=Puff Regime [6]=Voltage [7]=voltage
        QVariant sampleID  = cells[1][off + 1];
        QVariant resistance= cells[1][off + 3];
        QVariant viscosity = cells[1][off + 7];  // H2 per user correction
        QVariant media     = cells[2][off + 1];
        QVariant puffRegime= cells[2][off + 5];  // F3 (value next to "Puff Regime" at E3)
        QVariant voltage   = cells[2][off + 7];  // H3

        QString testName = QStringLiteral("Extended TPM Test");

        // Row 0: testName, _, "Date:", _, "Sample ID:", sampleID, "Heating Technology:", _
        newRow0[off + 0] = testName;
        newRow0[off + 2] = QStringLiteral("Date:");
        // No date in Format B
        newRow0[off + 4] = QStringLiteral("Sample ID:");
        newRow0[off + 5] = sampleID;
        newRow0[off + 6] = QStringLiteral("Heating Technology:");

        // Row 1: "Media:", media, "Resistance (Ω):", resistance, "Power:", _, "Puffing Regime:", regime
        newRow1[off + 0] = QStringLiteral("Media:");
        newRow1[off + 1] = media;
        newRow1[off + 2] = QString("Resistance (%1):").arg(QChar(0x03A9));
        newRow1[off + 3] = resistance;
        newRow1[off + 4] = QStringLiteral("Power:");
        newRow1[off + 6] = QStringLiteral("Puffing Regime:");
        newRow1[off + 7] = puffRegime;

        // Row 2: "Viscosity:", viscosity, "Tester:", _, "Voltage:", voltage, "Initial Oil Mass:", _
        newRow2[off + 0] = QStringLiteral("Viscosity:");
        newRow2[off + 1] = viscosity;
        newRow2[off + 2] = QStringLiteral("Tester:");
        newRow2[off + 4] = QStringLiteral("Voltage:");
        newRow2[off + 5] = voltage;
        newRow2[off + 6] = QStringLiteral("Initial Oil Mass:");
    }

    cells[0] = newRow0;
    cells[1] = newRow1;
    cells[2] = newRow2;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/ExcelReader.cpp
git commit -m "feat: implement Format B normalization (M1 Extended, PV1-5 remap)"
```

---

## Task 4: Implement normalizeSheet() for Format C (Gembox/T58G)

**Files:**
- Modify: `src/ExcelReader.cpp`

Format C has 12 cols matching Format E data positions — no data column remapping needed. Only metadata rows 0-2 need to be rewritten.

```
OLD Row 0: [0]="TPM" [1]="Date:" [2]=date [3]="Tester:" [4]=tester [5]="Project:" [6]=project [7]="Sample:" [8]=sampleID [9]="Did this burn?" [10]=burnAnswer
OLD Row 1: [0]="Media:" [1]=media [2]="Ri (Ohms)" [3]=resistance [4]="Power:" [5]=power [6]="Puffing Regime:" [7]=regime [8]="Usage Efficiency" [9]="Did this clog?" [10]=clogAnswer
OLD Row 2: [0]="Viscosity:" [1]=viscosity [2]="Rf (Ohms)" [3]=rfValue [4]="Voltage:" [5]=voltage [6]="Initial Oil Mass:" [7]=oilMass [8]=effValue [9]="Did this leak?" or "Puff Time" [10]=leakAnswer or puffTimeVal

NEW Row 0: [0]=testName [2]="Date:" [3]=date [4]="Sample ID:" [5]="project sampleID" [6]="Heating Technology:" [7]=_
NEW Row 1: [0]="Media:" [1]=media [2]="Resistance (Ω):" [3]=resistance [4]="Power:" [5]=power [6]="Puffing Regime:" [7]=regime
NEW Row 2: [0]="Viscosity:" [1]=viscosity [2]="Tester:" [3]=tester [4]="Voltage:" [5]=voltage [6]="Initial Oil Mass:" [7]=oilMass
```

- [ ] **Step 1: Write normalizeSheetFormatC() helper**

```cpp
// ---------------------------------------------------------------------------
// normalizeSheetFormatC — Gembox/T58G, 12 cols, metadata remap only
// ---------------------------------------------------------------------------
static void normalizeSheetFormatC(QVector<QVector<QVariant>>& cells)
{
    if (cells.size() < 4) return;

    const int COLS = 12;
    int totalCols = cells[3].size();
    int nSamples = totalCols / COLS;
    if (nSamples <= 0) return;

    int totalNewCols = nSamples * COLS;

    // Pad metadata rows
    for (int r = 0; r < 3; ++r) {
        while (cells[r].size() < nSamples * COLS)
            cells[r].append(QVariant());
    }

    QVector<QVariant> newRow0(totalNewCols);
    QVector<QVariant> newRow1(totalNewCols);
    QVector<QVariant> newRow2(totalNewCols);

    for (int s = 0; s < nSamples; ++s) {
        int off = s * COLS;

        // Extract old metadata
        // Row 0: [1]="Date:" [2]=date [3]="Tester:" [4]=tester [5]="Project:" [6]=project [7]="Sample:" [8]=sampleID
        QVariant date      = cells[0][off + 2];
        QVariant tester    = cells[0][off + 4];
        QString  project   = cells[0][off + 6].toString().trimmed();
        QString  sampleRaw = cells[0][off + 8].toString().trimmed();

        // Combine Project + Sample for the sampleID
        QString combinedID;
        if (!project.isEmpty() && !sampleRaw.isEmpty())
            combinedID = project + " " + sampleRaw;
        else if (!sampleRaw.isEmpty())
            combinedID = sampleRaw;
        else
            combinedID = project;

        // Row 0 burn status: [9]="Did this burn?" [10]=answer
        QVariant burnAnswer = cells[0][off + 10];

        // Row 1: [1]=media [3]=resistance [5]=power [7]=regime
        QVariant media      = cells[1][off + 1];
        QVariant resistance = cells[1][off + 3];
        QVariant power      = cells[1][off + 5];
        QVariant regime     = cells[1][off + 7];
        QVariant clogAnswer = cells[1][off + 10];

        // Row 2: [1]=viscosity [5]=voltage [7]=oilMass
        QVariant viscosity  = cells[2][off + 1];
        QVariant voltage    = cells[2][off + 5];
        QVariant oilMass    = cells[2][off + 7];

        // Check if row 2 col 9 is "Did this leak?" or "Puff Time" (ignore Puff Time)
        QString r2c9Label = cells[2][off + 9].toString().trimmed();
        QVariant leakAnswer;
        if (r2c9Label.contains("leak", Qt::CaseInsensitive))
            leakAnswer = cells[2][off + 10];

        // Use "TPM" from old row 0 col 0 as a generic test name;
        // the sheet name provides better context so keep it generic.
        QString testName = QStringLiteral("TPM Test");

        // Write Format E metadata
        newRow0[off + 0] = testName;
        newRow0[off + 2] = QStringLiteral("Date:");
        newRow0[off + 3] = date;
        newRow0[off + 4] = QStringLiteral("Sample ID:");
        newRow0[off + 5] = combinedID;
        newRow0[off + 6] = QStringLiteral("Heating Technology:");
        // No heating tech in Format C
        newRow0[off + 9] = QStringLiteral("Did this burn?");
        newRow0[off + 10]= burnAnswer;

        newRow1[off + 0] = QStringLiteral("Media:");
        newRow1[off + 1] = media;
        newRow1[off + 2] = QString("Resistance (%1):").arg(QChar(0x03A9));
        newRow1[off + 3] = resistance;
        newRow1[off + 4] = QStringLiteral("Power:");
        newRow1[off + 5] = power;
        newRow1[off + 6] = QStringLiteral("Puffing Regime:");
        newRow1[off + 7] = regime;
        newRow1[off + 9] = QStringLiteral("Did this clog?");
        newRow1[off + 10]= clogAnswer;

        newRow2[off + 0] = QStringLiteral("Viscosity:");
        newRow2[off + 1] = viscosity;
        newRow2[off + 2] = QStringLiteral("Tester:");
        newRow2[off + 3] = tester;
        newRow2[off + 4] = QStringLiteral("Voltage:");
        newRow2[off + 5] = voltage;
        newRow2[off + 6] = QStringLiteral("Initial Oil Mass:");
        newRow2[off + 7] = oilMass;
        newRow2[off + 9] = QStringLiteral("Did this leak?");
        newRow2[off + 10]= leakAnswer;
    }

    cells[0] = newRow0;
    cells[1] = newRow1;
    cells[2] = newRow2;

    // Rewrite header row to match Format E naming
    QVector<QVariant> newHeaders(totalNewCols);
    QStringList hdr = { "puffs", "Before weight (g)", "After weight (g)",
                        "Draw Pressure (kpa)", QString("Resistance (%1)").arg(QChar(0x03A9)),
                        "Smell", "Clog", "Notes", "TPM (mg/puff)",
                        "TPM Power Density (mg/(W*s))", "Variation in TPM (%)",
                        "Oil Consumed (Cumulative, g)" };
    for (int s = 0; s < nSamples; ++s) {
        int off = s * COLS;
        for (int c = 0; c < COLS; ++c)
            newHeaders[off + c] = hdr[c];
    }
    cells[3] = newHeaders;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/ExcelReader.cpp
git commit -m "feat: implement Format C normalization (Gembox/T58G metadata remap)"
```

---

## Task 5: Implement normalizeSheet() for Format D (Jan 2025 Standard)

**Files:**
- Modify: `src/ExcelReader.cpp`

Format D is nearly identical to Format E. Two issues:
1. Missing Heating Technology field (row 0 cols 6/7 are empty or used for something else)
2. Some sheets (Long Puff Test) have only 11 cols per sample (missing the 12th column)
3. Uses "Resistance (Ohms):" instead of "Resistance (Ω):"

The data columns are already in the correct positions. Only need to:
- Pad 11-col sheets to 12 cols
- Update the resistance label so detectTemplateVersion() recognizes it as normalized

- [ ] **Step 1: Write normalizeSheetFormatD() helper**

```cpp
// ---------------------------------------------------------------------------
// normalizeSheetFormatD — Jan 2025 Standard, 11-12 cols, pad + label fix
// ---------------------------------------------------------------------------
static void normalizeSheetFormatD(QVector<QVector<QVariant>>& cells)
{
    if (cells.size() < 4) return;

    const int NEW_COLS = 12;

    // Determine actual columns per sample by finding the second "puffs" in header row
    int headerCols = cells[3].size();
    int oldColsPerSample = 12; // default

    // Find the second occurrence of "puffs" in the header row
    int puffsCount = 0;
    for (int c = 0; c < headerCols; ++c) {
        QString h = cells[3][c].toString().trimmed().toLower();
        if (h == "puffs") {
            puffsCount++;
            if (puffsCount == 2) {
                oldColsPerSample = c; // second puffs is at start of sample 2
                break;
            }
        }
    }

    int nSamples = (oldColsPerSample > 0) ? headerCols / oldColsPerSample : 0;
    if (nSamples <= 0) {
        // Single sample — use total width
        nSamples = 1;
        oldColsPerSample = headerCols;
    }

    int totalNewCols = nSamples * NEW_COLS;

    if (oldColsPerSample == NEW_COLS) {
        // Already 12 cols — only need to fix resistance label in metadata
        // Update row 1 resistance label for each sample
        if (cells[1].size() >= totalNewCols) {
            for (int s = 0; s < nSamples; ++s) {
                int off = s * NEW_COLS;
                if (off + 2 < cells[1].size()) {
                    QString label = cells[1][off + 2].toString().trimmed();
                    if (label.contains("Resistance (Ohms)", Qt::CaseInsensitive)) {
                        cells[1][off + 2] = QString("Resistance (%1):").arg(QChar(0x03A9));
                    }
                }
                // Add Heating Technology label at row 0 if missing
                if (off + 6 < cells[0].size()) {
                    QString r0c6 = cells[0][off + 6].toString().trimmed();
                    if (r0c6.isEmpty()) {
                        cells[0][off + 6] = QStringLiteral("Heating Technology:");
                    }
                }
            }
        }
        // Update header row labels to Format E naming
        QStringList hdr = { "puffs", "Before weight (g)", "After weight (g)",
                            "Draw Pressure (kpa)", QString("Resistance (%1)").arg(QChar(0x03A9)),
                            "Smell", "Clog", "Notes", "TPM (mg/puff)",
                            "TPM Power Density (mg/(W*s))", "Variation in TPM (%)",
                            "Oil Consumed (Cumulative, g)" };
        for (int s = 0; s < nSamples; ++s) {
            int off = s * NEW_COLS;
            for (int c = 0; c < NEW_COLS && (off + c) < cells[3].size(); ++c)
                cells[3][off + c] = hdr[c];
        }
        return;
    }

    // 11-col variant (e.g., Long Puff Test) — pad each sample to 12 cols
    // Data columns 0-10 are in the same positions, just need to insert col 11 (Oil Consumed)
    for (int row = 4; row < cells.size(); ++row) {
        QVector<QVariant> oldRow = cells[row];
        while (oldRow.size() < nSamples * oldColsPerSample)
            oldRow.append(QVariant());

        QVector<QVariant> newRow(totalNewCols);
        for (int s = 0; s < nSamples; ++s) {
            int oldOff = s * oldColsPerSample;
            int newOff = s * NEW_COLS;
            for (int c = 0; c < oldColsPerSample && c < NEW_COLS; ++c)
                newRow[newOff + c] = oldRow[oldOff + c];
            // col 11 stays null — will be recalculated
        }
        cells[row] = newRow;
    }

    // Pad and fix metadata rows 0-2
    for (int r = 0; r < 3; ++r) {
        QVector<QVariant> oldRow = cells[r];
        while (oldRow.size() < nSamples * oldColsPerSample)
            oldRow.append(QVariant());

        QVector<QVariant> newRow(totalNewCols);
        for (int s = 0; s < nSamples; ++s) {
            int oldOff = s * oldColsPerSample;
            int newOff = s * NEW_COLS;
            for (int c = 0; c < oldColsPerSample && c < NEW_COLS; ++c)
                newRow[newOff + c] = oldRow[oldOff + c];
        }
        cells[r] = newRow;
    }

    // Fix resistance label and add Heating Technology
    for (int s = 0; s < nSamples; ++s) {
        int off = s * NEW_COLS;
        QString label = cells[1][off + 2].toString().trimmed();
        if (label.contains("Resistance (Ohms)", Qt::CaseInsensitive))
            cells[1][off + 2] = QString("Resistance (%1):").arg(QChar(0x03A9));
        QString r0c6 = cells[0][off + 6].toString().trimmed();
        if (r0c6.isEmpty())
            cells[0][off + 6] = QStringLiteral("Heating Technology:");
    }

    // Rewrite header row
    QVector<QVariant> newHeaders(totalNewCols);
    QStringList hdr = { "puffs", "Before weight (g)", "After weight (g)",
                        "Draw Pressure (kpa)", QString("Resistance (%1)").arg(QChar(0x03A9)),
                        "Smell", "Clog", "Notes", "TPM (mg/puff)",
                        "TPM Power Density (mg/(W*s))", "Variation in TPM (%)",
                        "Oil Consumed (Cumulative, g)" };
    for (int s = 0; s < nSamples; ++s) {
        int off = s * NEW_COLS;
        for (int c = 0; c < NEW_COLS; ++c)
            newHeaders[off + c] = hdr[c];
    }
    cells[3] = newHeaders;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/ExcelReader.cpp
git commit -m "feat: implement Format D normalization (Jan 2025 std, 11-col padding)"
```

---

## Task 6: Wire up normalizeSheet() dispatcher and integrate into data flow

**Files:**
- Modify: `src/ExcelReader.cpp`

The normalization must run **after** a sheet is selected but **before** any metadata extraction or sample counting. The best integration point is in `selectSheet()` — normalize immediately after selecting, so all subsequent calls (`getSampleCount`, `extractMetadata`, `getAllSamples`) see Format E data.

- [ ] **Step 1: Implement the normalizeSheet() dispatcher**

Add after the Format D helper:

```cpp
void ExcelReader::normalizeSheet(LegacyFormat fmt)
{
    if (fmt == LegacyFormat::FormatE || fmt == LegacyFormat::Unknown)
        return;  // nothing to do

    if (m_currentSheetIdx < 0 || m_currentSheetIdx >= m_sheets.size())
        return;

    QVector<QVector<QVariant>>& cells = m_sheets[m_currentSheetIdx].cells;

    switch (fmt) {
    case LegacyFormat::FormatA:
        normalizeSheetFormatA(cells);
        break;
    case LegacyFormat::FormatB:
        normalizeSheetFormatB(cells);
        break;
    case LegacyFormat::FormatC:
        normalizeSheetFormatC(cells);
        break;
    case LegacyFormat::FormatD:
        normalizeSheetFormatD(cells);
        break;
    default:
        break;
    }

    debugPrint("Sheet normalized to Format E layout");
}
```

- [ ] **Step 2: Call normalization from selectSheet()**

In `ExcelReader::selectSheet()` (around line 217-232), add the normalization call after the sheet is found:

Replace the existing `selectSheet` method:

```cpp
bool ExcelReader::selectSheet(const QString& sheetName)
{
    debugPrint("Selecting sheet: " + sheetName);
    for (int i = 0; i < m_sheets.size(); ++i) {
        if (m_sheets[i].name == sheetName) {
            m_currentSheetIdx = i;
            m_currentSheet    = sheetName;

            // Detect and normalize legacy formats before any data extraction.
            LegacyFormat fmt = detectLegacyFormat();
            if (fmt != LegacyFormat::FormatE && fmt != LegacyFormat::Unknown) {
                debugPrint("Normalizing legacy format for sheet: " + sheetName);
                normalizeSheet(fmt);
            }

            debugPrint("Sheet selected successfully");
            debugPrint("Sample count: " + QString::number(getSampleCount()));
            return true;
        }
    }
    m_lastError = "Sheet not found: " + sheetName;
    debugPrint("ERROR: " + m_lastError);
    return false;
}
```

- [ ] **Step 3: Add a guard against double-normalization**

If `selectSheet()` is called multiple times for the same sheet, we must not normalize again (the data is already rewritten). Add a tracking set.

In `ExcelReader.h`, add to the private section:

```cpp
    QSet<int> m_normalizedSheets;  // indices of sheets already normalized
```

Add `#include <QSet>` at the top of ExcelReader.h.

Then update `selectSheet()`:

```cpp
bool ExcelReader::selectSheet(const QString& sheetName)
{
    debugPrint("Selecting sheet: " + sheetName);
    for (int i = 0; i < m_sheets.size(); ++i) {
        if (m_sheets[i].name == sheetName) {
            m_currentSheetIdx = i;
            m_currentSheet    = sheetName;

            // Detect and normalize legacy formats — only once per sheet.
            if (!m_normalizedSheets.contains(i)) {
                LegacyFormat fmt = detectLegacyFormat();
                if (fmt != LegacyFormat::FormatE && fmt != LegacyFormat::Unknown) {
                    debugPrint("Normalizing legacy format for sheet: " + sheetName);
                    normalizeSheet(fmt);
                }
                m_normalizedSheets.insert(i);
            }

            debugPrint("Sheet selected successfully");
            debugPrint("Sample count: " + QString::number(getSampleCount()));
            return true;
        }
    }
    m_lastError = "Sheet not found: " + sheetName;
    debugPrint("ERROR: " + m_lastError);
    return false;
}
```

Also clear it in `closeFile()`:

```cpp
void ExcelReader::closeFile()
{
    debugPrint("Closing file");
    m_sheets.clear();
    m_normalizedSheets.clear();
    m_currentSheetIdx = -1;
    m_filePath.clear();
    m_currentSheet.clear();
}
```

- [ ] **Step 4: Commit**

```bash
git add src/ExcelReader.h src/ExcelReader.cpp
git commit -m "feat: wire up legacy format normalization in selectSheet()"
```

---

## Task 7: Handle special sheets (User Test Simulation, User Test - Full Cycle)

**Files:**
- Modify: `src/ExcelReader.cpp`

The "User Test Simulation" sheet in Format D (CPS2920 file) has 8 cols with a `Chronology` column and completely different metadata layout. The existing `isDeprecatedUserTestSimulation()` already detects this.

The "User Test - Full Cycle" sheet is a tracking table (not TPM data) — it should be passed through as-is (already handled by the "no samples found" path).

For the GTI ODM file (Format E), "User Test Simulation" uses the standard 12-col Format E layout — that's already handled correctly.

We need to ensure `detectLegacyFormat()` does NOT attempt normalization on 8-col User Test Simulation sheets. Add a guard:

- [ ] **Step 1: Add guard in detectLegacyFormat() for special sheets**

At the beginning of `detectLegacyFormat()`, add:

```cpp
    // Don't attempt legacy detection on sheets that have their own handling
    if (m_currentSheet.contains("User Test - Full Cycle", Qt::CaseInsensitive))
        return LegacyFormat::Unknown;

    // 8-col User Test Simulation is a deprecated format handled separately
    if (isDeprecatedUserTestSimulation())
        return LegacyFormat::Unknown;
```

- [ ] **Step 2: Add guard for non-TPM sheets**

Several sheets like "Test Plan", "Initial State Inspection", "Aerosol Temperature", "Anti-Burn Protection Test", "Temperature Cycling Test", "Quick Sensory Test", "Off-odor Score", "Sensory Consistency", "Heavy Metal Leaching Test", "High T High Humidity Test" have completely different layouts. They don't have "puffs" in row 3 col 0. Add:

```cpp
    // Only normalize sheets that look like TPM data sheets
    // (header row must start with "puffs" at col 0)
    QString headerCol0 = cell(3, 0);
    if (headerCol0.compare("puffs", Qt::CaseInsensitive) != 0 &&
        headerCol0.compare("Chronology", Qt::CaseInsensitive) != 0) {
        return LegacyFormat::Unknown;
    }
```

Place this after the `isDeprecatedUserTestSimulation` check, before the PV1 check.

- [ ] **Step 3: Commit**

```bash
git add src/ExcelReader.cpp
git commit -m "feat: guard legacy normalization against non-TPM special sheets"
```

---

## Task 8: Build and test with all 7 legacy files

**Files:**
- No file changes — validation only

- [ ] **Step 1: Build the project**

```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
"C:/Qt/6.10.1/mingw_64/bin/qmake.exe" DataViewerEnterprise.pro -spec win32-g++
"C:/Qt/Tools/mingw1310_64/bin/windres.exe" \
  "--preprocessor=C:/Qt/Tools/mingw1310_64/bin/gcc.exe" \
  "--preprocessor-arg=-E" "--preprocessor-arg=-xc-header" \
  "--preprocessor-arg=-DRC_INVOKED" \
  -i DataViewer_resource.rc -o release/DataViewer_resource_res.o \
  --include-dir=. -DUNICODE -D_UNICODE -DWIN32
"C:/Qt/Tools/mingw1310_64/bin/mingw32-make.exe" -f Makefile.Release \
  CXX="C:/Qt/Tools/mingw1310_64/bin/g++.exe" \
  LINKER="C:/Qt/Tools/mingw1310_64/bin/g++.exe"
```

Expected: Clean build with no errors.

- [ ] **Step 2: Test each legacy file loads correctly**

Launch the built application and load each of the 7 files from `C:/Users/S1134987/Documents/Weekly Reports 3-23-2026 snapshot - local/example good tpm files/`. For each file, verify:

1. File loads without errors
2. Sheet tabs appear correctly
3. Sample data is populated in the table
4. TPM values are reasonable (not 0 or NaN)
5. Metadata (sample ID, media, voltage, resistance) displays correctly
6. Plots render with data points

Check the debug log at `%LOCALAPPDATA%/DataViewer-Enterprise/dve.log` for normalization messages:
- "Detected legacy Format A/B/C/D" messages
- "Sheet normalized to Format E layout" messages
- No error messages

- [ ] **Step 3: Verify GTI ODM file (Format E) still works unchanged**

Load `GTI ODM testing 3-16-2026.xlsx` and confirm it works identically to before — normalization should NOT trigger for this file. Check log for "Detected Format E — no normalization needed" messages.

- [ ] **Step 4: Fix any issues found, then commit**

```bash
git add -A
git commit -m "fix: address issues found during legacy format testing"
```

---

## Task 9: Final commit and cleanup

- [ ] **Step 1: Final build verification**

Run a clean build and confirm no warnings or errors.

- [ ] **Step 2: Create summary commit**

If no issues were found in Task 8 and all intermediate commits are clean, no action needed. Otherwise, create a final cleanup commit.

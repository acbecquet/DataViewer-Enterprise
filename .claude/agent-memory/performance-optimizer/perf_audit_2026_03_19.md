---
name: Performance audit 2026-03-19
description: Complete line-level performance audit of DataViewer Enterprise codebase covering all key source files. 29 findings from CRITICAL to LOW.
type: project
---

Full codebase performance audit completed 2026-03-19.

**Why:** User requested thorough analysis to identify optimization opportunities without sacrificing image quality or removing features.

**How to apply:** Use this as the prioritized backlog for implementing performance fixes. Start with CRITICAL (Finding 27: Python debouncing), then HIGH items.

## Key Findings Summary

**CRITICAL (1):**
- F27: MainWindow writeCellToExcel spawns Python per edit -- needs debouncing with QTimer (500ms)

**HIGH (7):**
- F2: ZipWriter::toByteArray no reserve -- multi-MB reallocations
- F4: DatabaseManager missing PRAGMA synchronous=NORMAL (safe with WAL)
- F5: DatabaseManager::loadFile N+1 queries for data_rows and images
- F11: ReportGenerator::loadAndCropImage double-decodes JPEG for fast-path
- F15: MainWindow::displayCurrentSample 600+ heap allocs per navigation
- F18: MainWindow::buildCleanedFile deep-copies entire FileResult every report
- F24: PptxWriter duplicate media blobs per slide (bg/logo repeated N times)

**MEDIUM (13):**
- F26: findPython probes 3 processes each call -- cache result
- F6: deduplicateFiles deletes without transaction
- F7/F8: JSON parsing in SQL queries -- use json_extract()
- F9: buildTable contains() chain per cell
- F12: compressImageBlob same double-decode pattern
- F13/F14: QString concatenation in loops (PptxWriter XML)
- F16: updateProperties 60 allocs per navigation
- F17: full sheet recalc for single-cell property edit
- F19/F20: unnecessary copies and double-builds for cleanup

**LOW (8):**
- F1: hand-rolled CRC32 (use zlib), F3: byte-at-a-time append
- F10: medianOf O(n log n) -> O(n), F21-23: RadarChart minor
- F25: O(n) row mapping, F28: isNotes contains per cell, F29: mkpath per sample

## Files Analyzed
- src/database/DatabaseManager.cpp (949 lines)
- src/reporting/PptxWriter.cpp (~1150 lines)
- src/reporting/ReportGenerator.cpp (624 lines)
- src/utils/ZipWriter.cpp (261 lines)
- src/MainWindow.cpp (2247 lines)
- src/ui/SensoryDialog.cpp (~1400 lines)
- src/ui/RadarChartWidget.cpp (267 lines)
- src/widgets/RibbonWidget.cpp (372 lines -- no issues found)

## Notes
- No pipeline/ directory exists separately
- ZipWriter is at src/utils/ZipWriter.cpp (not src/reporting/)
- RibbonWidget.cpp is clean -- static UI construction, no hot paths
- SensoryDialog chart refresh already well-debounced (150ms timer)

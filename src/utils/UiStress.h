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

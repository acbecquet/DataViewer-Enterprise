#pragma once

#include <QDialog>

#include "utils/OutputPaths.h"   // ReportMode

namespace DVE {

// "Pick a default mode:" picker (v2.7.0). Shown once on the first launch after
// install (no saved value yet) and reopened any time via Settings > Set Default
// Mode. Clicking a mode persists it as the system-wide default the app opens
// in; Esc/close selects nothing and writes nothing (callers decide the
// first-run fallback).
class DefaultModeDialog : public QDialog {
    Q_OBJECT
public:
    explicit DefaultModeDialog(QWidget* parent = nullptr);

    // The mode picked via one of the three buttons; meaningful only when the
    // dialog was accepted. Tpm until a button is clicked.
    ReportMode selectedMode() const { return m_selected; }

    // Persisted default ("SDR"/"DataViewerEnterprise", key ui/defaultMode).
    // savedDefaultMode() falls back to Tpm when unset or unparseable.
    static bool       hasSavedDefaultMode();
    static ReportMode savedDefaultMode();
    static void       setSavedDefaultMode(ReportMode mode);

private:
    ReportMode m_selected = ReportMode::Tpm;
};

} // namespace DVE

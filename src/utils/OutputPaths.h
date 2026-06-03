#pragma once

#include <QString>
#include <QStringList>

namespace DVE {

enum class ReportMode { Tpm, Sensory, DetailedSensory };

// Single source of truth for report/save output directories + report filenames.
// Pure resolvers (firstExistingDir / resolve* / reportFileName / sanitize) are
// side-effect-free; only configuredDir/setConfiguredDir touch QSettings.
class OutputPaths {
public:
    // Persisted per-mode output folders. QSettings("SDR","DataViewerEnterprise"),
    // keys output/tpmDir, output/sensoryDir, output/detailedDir. "" when unset.
    static QString configuredDir(ReportMode mode);
    static void    setConfiguredDir(ReportMode mode, const QString& dir);

    // QStandardPaths Documents, falling back to ~/Documents.
    static QString documentsDir();

    // Report dialogs: configuredDir(mode) -> lastUsedDir -> Documents (first existing).
    static QString resolveReportDir(ReportMode mode, const QString& lastUsedDir);
    // Non-report dialogs: lastUsedDir -> Documents.
    static QString resolveSaveDir(const QString& lastUsedDir);

    // "Foo" -> "Foo_report.pptx"; ("Foo","Sheet 1") -> "Foo_Sheet_1_report.pptx".
    // A trailing ".pptx" on base is stripped; empty base -> "untitled".
    static QString reportFileName(const QString& base, const QString& sheet = QString());

    // Strip Windows-illegal chars (\ / : * ? " < > |), whitespace -> \'_\',
    // collapse runs of \'_\', trim leading/trailing \'_\'.
    static QString sanitize(const QString& raw);

    // First entry that is non-empty AND exists on disk; else fallback. Pure.
    static QString firstExistingDir(const QStringList& candidates, const QString& fallback);

private:
    static QString settingsKey(ReportMode mode);
};

} // namespace DVE

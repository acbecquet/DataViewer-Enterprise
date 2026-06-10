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

    // Every mode dialog (report / save / load / new): configuredDir(mode) ->
    // lastUsedDir -> Documents (first existing). The configured per-mode folder
    // always wins when set.
    static QString resolveDir(ReportMode mode, const QString& lastUsedDir);

    // "Foo" -> "Foo_report.pptx"; ("Foo","Sheet 1") -> "Foo_Sheet_1_report.pptx".
    // A trailing ".pptx" on base is stripped; empty base -> "untitled".
    static QString reportFileName(const QString& base, const QString& sheet = QString());

    // Auto-derived save path: resolveDir(mode,lastUsedDir) + "/" + sanitize(label) + ext.
    // Single source of truth for the silent (no-dialog) sensory/detailed save filenames.
    // `ext` may be given with or without a leading dot. An empty/whitespace label falls
    // back to "untitled" so the result is always a valid path.
    static QString autoSavePath(ReportMode mode, const QString& sessionLabel,
                                const QString& lastUsedDir, const QString& ext);

    // Strip Windows-illegal chars (\ / : * ? " < > |), whitespace -> \'_\',
    // collapse runs of \'_\', trim leading/trailing \'_\'.
    static QString sanitize(const QString& raw);

    // v2.5.0 RC4 — next iterator for a duplicate/renamed session name. If the
    // input ends in a trailing "_<digits>" run it is parsed as an integer and
    // incremented ("T_1"->"T_2", "T_9"->"T_10", "T_03"->"T_4" — leading zeros
    // are NOT preserved); otherwise "_1" is appended ("T"->"T_1"). ONLY a
    // trailing _<digits> run counts as an iterator: an underscore-word like
    // "Vape_Test" is left intact and gets a fresh "_1" ("Vape_Test_1"). Pure.
    static QString nextSuffixedName(const QString& name);

    // First entry that is non-empty AND exists on disk; else fallback. Pure.
    static QString firstExistingDir(const QStringList& candidates, const QString& fallback);

private:
    static QString settingsKey(ReportMode mode);
};

} // namespace DVE

QT += core gui widgets sql concurrent network

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17

TARGET   = DataViewer
TEMPLATE = app

# ─── Single-source version ────────────────────────────────────────────────────
# Bump this VERSION to release. main.cpp picks it up via the DVE_APP_VERSION
# preprocessor define below. build_installer.bat parses the same line out of
# this file and passes it to ISCC as /DAppVersion=<value>.
VERSION = 2.7.1
DEFINES += DVE_APP_VERSION=\\\"$$VERSION\\\"

# ─── QXlsx (Excel read/write) ─────────────────────────────────────────────────
include(external/QXlsx/QXlsx/QXlsx.pri)
INCLUDEPATH += $$PWD/external/QXlsx/header

# ─── zlib for PPTX ZIP writer (ships with MinGW) ──────────────────────────────
LIBS += -lz

# ─── Include search paths ─────────────────────────────────────────────────────
INCLUDEPATH += src \
               src/pipeline \
               src/reporting \
               src/plotting \
               src/widgets \
               src/database \
               src/utils \
               src/ui

# ─── Sources ──────────────────────────────────────────────────────────────────
SOURCES += \
    src/main.cpp \
    src/MainWindow.cpp \
    src/RemoteCellHelpers.cpp \
    src/ExcelReader.cpp \
    src/pipeline/TpmCalculator.cpp \
    src/pipeline/SheetProcessors.cpp \
    src/pipeline/DataCleanup.cpp \
    src/pipeline/RegimeUtils.cpp \
    src/pipeline/DataProcessor.cpp \
    src/pipeline/SensoryData.cpp \
    src/pipeline/DetailedSensoryData.cpp \
    src/pipeline/NotesStory.cpp \
    src/pipeline/ReportDataJson.cpp \
    src/reporting/LayoutCommand.cpp \
    src/reporting/PptxWriter.cpp \
    src/reporting/ReportGenerator.cpp \
    src/reporting/ReportLayout.cpp \
    src/reporting/SensoryReportSource.cpp \
    src/plotting/PlotEngine.cpp \
    src/plotting/PlotWidget.cpp \
    src/widgets/RibbonWidget.cpp \
    src/widgets/PresenceDotsDelegate.cpp \
    src/widgets/PresenceAvatarBar.cpp \
    src/widgets/RowDeletedBanner.cpp \
    src/widgets/OfflineBanner.cpp \
    src/widgets/IncompleteDataBanner.cpp \
    src/widgets/NotesStoryPanel.cpp \
    src/widgets/FlowLayout.cpp \
    src/widgets/ScrollHost.cpp \
    src/utils/AppTheme.cpp \
    src/utils/ResponsiveLayout.cpp \
    src/utils/ZipWriter.cpp \
    src/utils/XmlBuilder.cpp \
    src/utils/ImageUtils.cpp \
    src/database/DatabaseManager.cpp \
    src/database/DatabaseOps.cpp \
    src/database/PersistWorker.cpp \
    src/database/SnapshotRegenWorker.cpp \
    src/database/CompatClassifier.cpp \
    src/database/IdentityManager.cpp \
    src/database/IdentityPromptDialog.cpp \
    src/database/ConfigLoader.cpp \
    src/database/PostgresConnection.cpp \
    src/database/NotificationListener.cpp \
    src/database/LiveSync.cpp \
    src/database/LiveSyncWorker.cpp \
    src/database/VersionLookup.cpp \
    src/database/PresenceManager.cpp \
    src/database/MigrationReport.cpp \
    src/database/MigrationTool.cpp \
    src/database/UniqueViolationDialog.cpp \
    src/database/OfflineSnapshot.cpp \
    src/database/ConnectionMonitor.cpp \
    src/database/DbRepair.cpp \
    src/database/RawGridJson.cpp \
    src/database/WriteOutcome.cpp \
    src/ui/NewFileDialog.cpp \
    src/ui/HeaderEditDialog.cpp \
    src/ui/ImageViewDialog.cpp \
    src/ui/DatabaseBrowserDialog.cpp \
    src/ui/DataCleanupDialog.cpp \
    src/ui/DefaultModeDialog.cpp \
    src/ui/ImageInboxDialog.cpp \
    src/ui/SensoryPanel.cpp \
    src/ui/DetailedSensoryPanel.cpp \
    src/ui/PropertiesPanel.cpp \
    src/ui/RadarChartWidget.cpp \
    src/ui/ReportPreviewDialog.cpp \
    src/ui/RecoverDialog.cpp \
    src/ui/SamplesCheckboxPanel.cpp \
    src/ui/SlideCanvasItems.cpp \
    src/ui/SopDialog.cpp \
    src/utils/SingleInstance.cpp \
    src/utils/UpdateChecker.cpp \
    src/utils/SelfTest.cpp \
    src/utils/UiStress.cpp \
    src/utils/SopLoader.cpp \
    src/utils/OutputPaths.cpp \
    src/utils/RecoveryManager.cpp \
    src/utils/MipFallback.cpp \
    src/utils/ExcelWritePayload.cpp

# ─── Headers ──────────────────────────────────────────────────────────────────
HEADERS += \
    src/MainWindow.h \
    src/RemoteCellHelpers.h \
    src/ExcelReader.h \
    src/pipeline/ReportData.h \
    src/pipeline/NotesStory.h \
    src/pipeline/ReportDataJson.h \
    src/pipeline/TpmCalculator.h \
    src/pipeline/SheetProcessors.h \
    src/pipeline/DataCleanup.h \
    src/pipeline/RegimeUtils.h \
    src/pipeline/DataProcessor.h \
    src/reporting/IReportSource.h \
    src/reporting/LayoutCommand.h \
    src/reporting/PptxWriter.h \
    src/reporting/ReportGenerator.h \
    src/reporting/ReportLayout.h \
    src/reporting/SensoryReportSource.h \
    src/plotting/PlotEngine.h \
    src/plotting/PlotWidget.h \
    src/widgets/RibbonWidget.h \
    src/widgets/PresenceDotsDelegate.h \
    src/widgets/PresenceAvatarBar.h \
    src/widgets/RowDeletedBanner.h \
    src/widgets/OfflineBanner.h \
    src/widgets/IncompleteDataBanner.h \
    src/widgets/NotesStoryPanel.h \
    src/widgets/FlowLayout.h \
    src/widgets/ScrollHost.h \
    src/utils/AppTheme.h \
    src/utils/ResponsiveLayout.h \
    src/utils/ZipWriter.h \
    src/utils/XmlBuilder.h \
    src/utils/ImageUtils.h \
    src/database/DatabaseManager.h \
    src/database/DatabaseOps.h \
    src/database/PersistWorker.h \
    src/database/SnapshotRegenWorker.h \
    src/database/CompatClassifier.h \
    src/database/IdentityManager.h \
    src/database/IdentityPromptDialog.h \
    src/database/ConfigLoader.h \
    src/database/PostgresConnection.h \
    src/database/NotificationListener.h \
    src/database/LiveSync.h \
    src/database/LiveSyncWorker.h \
    src/database/VersionLookup.h \
    src/database/PresenceManager.h \
    src/database/MigrationReport.h \
    src/database/MigrationTool.h \
    src/database/UniqueViolationDialog.h \
    src/database/OfflineSnapshot.h \
    src/database/ConnectionMonitor.h \
    src/database/DbRepair.h \
    src/database/RawGridJson.h \
    src/database/WriteOutcome.h \
    src/ui/NewFileDialog.h \
    src/ui/HeaderEditDialog.h \
    src/ui/ImageViewDialog.h \
    src/ui/DatabaseBrowserDialog.h \
    src/ui/DataCleanupDialog.h \
    src/ui/DefaultModeDialog.h \
    src/ui/ImageInboxDialog.h \
    src/ui/SensoryPanel.h \
    src/ui/DetailedSensoryPanel.h \
    src/ui/PropertiesPanel.h \
    src/ui/RadarChartWidget.h \
    src/ui/ReportPreviewDialog.h \
    src/ui/RecoverDialog.h \
    src/ui/SamplesCheckboxPanel.h \
    src/ui/SlideCanvasItems.h \
    src/ui/SopDialog.h \
    src/pipeline/SensoryData.h \
    src/pipeline/DetailedSensoryData.h \
    src/utils/SingleInstance.h \
    src/utils/UpdateChecker.h \
    src/utils/SelfTest.h \
    src/utils/UiStress.h \
    src/utils/SopLoader.h \
    src/utils/OutputPaths.h \
    src/utils/RecoveryManager.h \
    src/utils/MipFallback.h \
    src/utils/ExcelWritePayload.h

DEFINES += QT_DEPRECATED_WARNINGS
QMAKE_CXXFLAGS += -Wno-deprecated-declarations

# ─── Release optimizations ───────────────────────────────────────────────────
CONFIG(release, debug|release) {
    QMAKE_CXXFLAGS += -O2 -Wpedantic
}

# Treat warnings as errors in all builds
QMAKE_CXXFLAGS += -Werror

# Relax specific warnings triggered inside external/QXlsx
QMAKE_CXXFLAGS += -Wno-error=unused-but-set-variable

# ─── Resources (icons, branding) ─────────────────────────────────────────────
# NOTE: Qt 6.10.1 rcc generates binary TSD format instead of C++ source,
# so we load icons from filesystem via resourcePath() instead.
# RESOURCES += resources/resources.qrc
RC_ICONS   = resources/images/ccell_icon.ico

qnx: target.path = /tmp/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

QT += core gui widgets sql concurrent

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17

TARGET   = DataViewer
TEMPLATE = app

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
    src/ExcelReader.cpp \
    src/pipeline/TpmCalculator.cpp \
    src/pipeline/SheetProcessors.cpp \
    src/pipeline/DataProcessor.cpp \
    src/reporting/PptxWriter.cpp \
    src/reporting/ReportGenerator.cpp \
    src/plotting/PlotEngine.cpp \
    src/plotting/PlotWidget.cpp \
    src/widgets/RibbonWidget.cpp \
    src/utils/AppTheme.cpp \
    src/utils/ZipWriter.cpp \
    src/utils/XmlBuilder.cpp \
    src/database/DatabaseManager.cpp \
    src/ui/NewFileDialog.cpp \
    src/ui/HeaderEditDialog.cpp \
    src/ui/ImageViewDialog.cpp \
    src/ui/DatabaseBrowserDialog.cpp

# ─── Headers ──────────────────────────────────────────────────────────────────
HEADERS += \
    src/MainWindow.h \
    src/ExcelReader.h \
    src/pipeline/ReportData.h \
    src/pipeline/TpmCalculator.h \
    src/pipeline/SheetProcessors.h \
    src/pipeline/DataProcessor.h \
    src/reporting/PptxWriter.h \
    src/reporting/ReportGenerator.h \
    src/plotting/PlotEngine.h \
    src/plotting/PlotWidget.h \
    src/widgets/RibbonWidget.h \
    src/utils/AppTheme.h \
    src/utils/ZipWriter.h \
    src/utils/XmlBuilder.h \
    src/database/DatabaseManager.h \
    src/ui/NewFileDialog.h \
    src/ui/HeaderEditDialog.h \
    src/ui/ImageViewDialog.h \
    src/ui/DatabaseBrowserDialog.h

DEFINES += QT_DEPRECATED_WARNINGS

qnx: target.path = /tmp/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

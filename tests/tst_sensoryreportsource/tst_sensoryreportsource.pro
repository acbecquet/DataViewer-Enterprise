QT += core gui sql network widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/reporting ../../src/pipeline ../../src/database \
               ../../src/utils ../../src/ui

SOURCES += tst_sensoryreportsource.cpp \
           ../../src/reporting/SensoryReportSource.cpp \
           ../../src/reporting/ReportLayout.cpp \
           ../../src/reporting/PptxWriter.cpp \
           ../../src/database/DatabaseManager.cpp \
           ../../src/model/MetricRegistry.cpp \
           ../../src/model/SchemaDrivenReader.cpp \
           ../../src/database/DatabaseOps.cpp \
           ../../src/database/RawGridJson.cpp \
           ../../src/database/OfflineSnapshot.cpp \
           ../../src/utils/MipFallback.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/ConfigLoader.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/pipeline/SensoryData.cpp \
           ../../src/pipeline/DetailedSensoryData.cpp \
           ../../src/utils/ZipWriter.cpp \
           ../../src/utils/XmlBuilder.cpp \
           ../../src/utils/ImageUtils.cpp \
           ../../src/ui/RadarChartWidget.cpp \
           ../../src/plotting/SensoryChartNotes.cpp \
           ../../src/utils/SampleColorMap.cpp \
           ../../src/utils/AppTheme.cpp \
           ../../src/utils/OutputPaths.cpp

HEADERS += ../../src/reporting/SensoryReportSource.h \
           ../../src/reporting/IReportSource.h \
           ../../src/reporting/ReportLayout.h \
           ../../src/reporting/PptxWriter.h \
           ../../src/database/DatabaseManager.h \
           ../../src/database/DatabaseOps.h \
           ../../src/database/OfflineSnapshot.h \
           ../../src/utils/MipFallback.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/ConfigLoader.h \
           ../../src/database/IdentityManager.h \
           ../../src/pipeline/SensoryData.h \
           ../../src/utils/ZipWriter.h \
           ../../src/utils/XmlBuilder.h \
           ../../src/utils/ImageUtils.h \
           ../../src/ui/RadarChartWidget.h \
           ../../src/utils/OutputPaths.h

LIBS += -lz

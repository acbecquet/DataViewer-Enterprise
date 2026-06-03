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
           ../../src/database/OfflineSnapshot.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/ConfigLoader.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/pipeline/SensoryData.cpp \
           ../../src/utils/ZipWriter.cpp \
           ../../src/utils/XmlBuilder.cpp \
           ../../src/utils/ImageUtils.cpp \
           ../../src/ui/RadarChartWidget.cpp \
           ../../src/utils/AppTheme.cpp \
           ../../src/utils/OutputPaths.cpp

HEADERS += ../../src/reporting/SensoryReportSource.h \
           ../../src/reporting/IReportSource.h \
           ../../src/reporting/ReportLayout.h \
           ../../src/reporting/PptxWriter.h \
           ../../src/database/DatabaseManager.h \
           ../../src/database/OfflineSnapshot.h \
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

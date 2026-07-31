QT += core sql testlib network
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting                ../../src/reporting ../../src/database ../common

SOURCES += tst_planc_checkpoint.cpp            ../../src/database/DatabaseManager.cpp            ../../src/model/MetricRegistry.cpp            ../../src/model/SchemaDrivenReader.cpp            ../../src/database/DatabaseOps.cpp            ../../src/database/MetricDefCache.cpp            ../../src/database/RawGridJson.cpp            ../../src/database/OfflineSnapshot.cpp            ../../src/utils/MipFallback.cpp            ../../src/database/ConnectionMonitor.cpp            ../../src/database/PostgresConnection.cpp            ../../src/database/ConfigLoader.cpp            ../../src/database/IdentityManager.cpp            ../../src/utils/OutputPaths.cpp            ../../src/pipeline/SensoryData.cpp            ../../src/pipeline/DetailedSensoryData.cpp

HEADERS += ../../src/database/DatabaseManager.h            ../../src/database/DatabaseOps.h            ../../src/database/MetricDefCache.h            ../../src/database/OfflineSnapshot.h            ../../src/utils/MipFallback.h            ../../src/database/ConnectionMonitor.h            ../../src/database/PostgresConnection.h            ../../src/database/ConfigLoader.h            ../../src/database/IdentityManager.h            ../../src/utils/OutputPaths.h            ../../src/pipeline/ReportData.h            ../../src/pipeline/SensoryData.h            ../../src/pipeline/DetailedSensoryData.h

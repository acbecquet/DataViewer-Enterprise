QT += core sql testlib network
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting                ../../src/reporting ../../src/database ../common

SOURCES += tst_offlinesnapshot.cpp            ../../src/database/OfflineSnapshot.cpp            ../../src/database/MetricDefCache.cpp            ../../src/database/PostgresConnection.cpp            ../../src/database/ConfigLoader.cpp            ../../src/pipeline/SensoryData.cpp            ../../src/database/RawGridJson.cpp            ../../src/utils/MipFallback.cpp            ../../src/database/SnapshotRegenWorker.cpp

HEADERS += ../../src/database/OfflineSnapshot.h            ../../src/database/MetricDefCache.h            ../../src/database/PostgresConnection.h            ../../src/database/ConfigLoader.h            ../../src/pipeline/ReportData.h            ../../src/pipeline/SensoryData.h            ../../src/pipeline/DetailedSensoryData.h            ../../src/database/RawGridJson.h            ../../src/utils/MipFallback.h            ../../src/database/SnapshotRegenWorker.h

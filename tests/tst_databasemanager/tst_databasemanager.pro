QT += core sql testlib network
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting \
               ../../src/reporting ../../src/database ../common

SOURCES += tst_databasemanager.cpp \
           ../../src/database/DatabaseManager.cpp \
           ../../src/database/RawGridJson.cpp \
           ../../src/database/OfflineSnapshot.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/ConfigLoader.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/pipeline/SensoryData.cpp

HEADERS += ../../src/database/DatabaseManager.h \
           ../../src/database/OfflineSnapshot.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/ConfigLoader.h \
           ../../src/database/IdentityManager.h \
           ../../src/pipeline/ReportData.h \
           ../../src/pipeline/SensoryData.h \
           ../../src/pipeline/DetailedSensoryData.h \
           ../common/TestHelpers.h

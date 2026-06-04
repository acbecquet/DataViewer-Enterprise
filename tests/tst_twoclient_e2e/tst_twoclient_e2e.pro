QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database ../../src/pipeline

SOURCES += tst_twoclient_e2e.cpp \
           ../../src/database/DatabaseManager.cpp \
           ../../src/database/RawGridJson.cpp \
           ../../src/database/OfflineSnapshot.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/NotificationListener.cpp \
           ../../src/database/PresenceManager.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/database/ConfigLoader.cpp \
           ../../src/database/LiveSync.cpp \
           ../../src/database/LiveSyncWorker.cpp \
           ../../src/pipeline/SensoryData.cpp

HEADERS += ../../src/database/DatabaseManager.h \
           ../../src/database/OfflineSnapshot.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/NotificationListener.h \
           ../../src/database/PresenceManager.h \
           ../../src/database/IdentityManager.h \
           ../../src/database/ConfigLoader.h \
           ../../src/database/LiveSync.h \
           ../../src/database/LiveSyncWorker.h \
           ../../src/pipeline/ReportData.h \
           ../../src/pipeline/SensoryData.h \
           ../../src/pipeline/DetailedSensoryData.h

QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database ../../src/pipeline

SOURCES += tst_livesync.cpp \
           ../../src/database/LiveSync.cpp \
           ../../src/database/LiveSyncWorker.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/NotificationListener.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/database/ConfigLoader.cpp \
           ../../src/database/OfflineSnapshot.cpp \
           ../../src/utils/MipFallback.cpp \
           ../../src/database/RawGridJson.cpp \
           ../../src/pipeline/SensoryData.cpp

HEADERS += ../../src/database/LiveSync.h \
           ../../src/database/LiveSyncWorker.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/NotificationListener.h \
           ../../src/database/IdentityManager.h \
           ../../src/database/ConfigLoader.h \
           ../../src/database/OfflineSnapshot.h \
           ../../src/utils/MipFallback.h \
           ../../src/database/RawGridJson.h \
           ../../src/pipeline/ReportData.h \
           ../../src/pipeline/SensoryData.h \
           ../../src/pipeline/DetailedSensoryData.h

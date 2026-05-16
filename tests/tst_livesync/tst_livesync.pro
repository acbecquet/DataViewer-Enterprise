QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_livesync.cpp \
           ../../src/database/LiveSync.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/NotificationListener.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/database/ConfigLoader.cpp

HEADERS += ../../src/database/LiveSync.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/NotificationListener.h \
           ../../src/database/IdentityManager.h \
           ../../src/database/ConfigLoader.h

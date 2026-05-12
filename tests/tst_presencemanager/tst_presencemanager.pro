QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_presencemanager.cpp \
           ../../src/database/PresenceManager.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/database/ConfigLoader.cpp

HEADERS += ../../src/database/PresenceManager.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/IdentityManager.h \
           ../../src/database/ConfigLoader.h

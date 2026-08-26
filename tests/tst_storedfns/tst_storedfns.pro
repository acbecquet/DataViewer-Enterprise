QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database ../common

SOURCES += tst_storedfns.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/ConfigLoader.cpp \
           ../../src/database/IdentityManager.cpp

HEADERS += ../../src/database/PostgresConnection.h \
           ../../src/database/ConfigLoader.h \
           ../../src/database/IdentityManager.h

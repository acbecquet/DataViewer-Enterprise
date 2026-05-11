QT       += core testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_identitymanager.cpp            ../../src/database/IdentityManager.cpp

HEADERS += ../../src/database/IdentityManager.h

QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_postgresconnection.cpp            ../../src/database/PostgresConnection.cpp            ../../src/database/ConfigLoader.cpp

HEADERS += ../../src/database/PostgresConnection.h            ../../src/database/ConfigLoader.h

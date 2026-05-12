QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_notificationlistener.cpp \
           ../../src/database/NotificationListener.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/ConfigLoader.cpp

HEADERS += ../../src/database/NotificationListener.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/ConfigLoader.h

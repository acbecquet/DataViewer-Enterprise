QT       += core gui widgets testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/widgets

SOURCES += tst_presencedotsdelegate.cpp            ../../src/widgets/PresenceDotsDelegate.cpp

HEADERS += ../../src/widgets/PresenceDotsDelegate.h

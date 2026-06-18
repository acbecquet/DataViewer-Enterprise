QT       += core testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_compatclassifier.cpp \
           ../../src/database/CompatClassifier.cpp

HEADERS += ../../src/database/CompatClassifier.h

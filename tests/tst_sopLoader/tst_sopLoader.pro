QT += core testlib
CONFIG += qt warn_on testcase c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_sopLoader

INCLUDEPATH += $$PWD \
               $$PWD/../../src \
               $$PWD/../../src/utils \
               $$PWD/../common

include($$PWD/../../external/QXlsx/QXlsx/QXlsx.pri)

SOURCES += tst_sopLoader.cpp \
           $$PWD/../../src/utils/SopLoader.cpp

HEADERS += $$PWD/../../src/utils/SopLoader.h

LIBS += -lz

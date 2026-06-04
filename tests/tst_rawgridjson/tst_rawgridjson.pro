QT += core testlib
QT -= gui
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_rawgridjson

INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_rawgridjson.cpp \
           ../../src/database/RawGridJson.cpp
HEADERS += ../../src/database/RawGridJson.h

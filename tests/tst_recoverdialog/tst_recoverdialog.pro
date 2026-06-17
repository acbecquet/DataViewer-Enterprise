QT += core gui widgets testlib concurrent
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_recoverdialog

INCLUDEPATH += ../../src ../../src/ui ../../src/utils ../../src/pipeline

HEADERS += \
    ../../src/ui/RecoverDialog.h \
    ../../src/utils/RecoveryManager.h \
    ../../src/utils/MipFallback.h

SOURCES += \
    tst_recoverdialog.cpp \
    ../../src/ui/RecoverDialog.cpp \
    ../../src/utils/RecoveryManager.cpp \
    ../../src/utils/MipFallback.cpp

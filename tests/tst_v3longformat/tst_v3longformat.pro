QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

# The 3b half of this suite links NO application source: the thing under test is
# server-side SQL (dve_migrate_to_long_format() plus the data_rows_v / samples_v
# compat views). Every value comparison is executed INSIDE Postgres and only
# counts and booleans cross the wire, so the harness cannot be fooled by Qt's
# double<->text formatting on the way in or out.
#
# v3 Phase 3c adds ONE C++ unit: MetricDefCache, the (kind, key) ->
# metric_defs.id resolver the save path uses. It depends on nothing but Qt SQL,
# so it comes in alone - no DatabaseManager, no pipeline. Keep it that way; the
# whole point of this suite is that it is cheap and SQL-first.
INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_v3longformat.cpp \
           ../../src/database/MetricDefCache.cpp

HEADERS += ../../src/database/MetricDefCache.h

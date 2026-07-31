QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

# Deliberately links NO application source. This suite is a rehearsal harness
# for deploy/postgres/migrations/2026-07-31-v3-long-format.sql - the thing under
# test is server-side SQL (dve_migrate_to_long_format() plus the data_rows_v /
# samples_v compat views), not C++. Every value comparison is executed INSIDE
# Postgres and only counts and booleans cross the wire, so the harness cannot be
# fooled by Qt's double<->text formatting on the way in or out.
SOURCES += tst_v3longformat.cpp

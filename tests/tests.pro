TEMPLATE = subdirs

# TODO(Plan B Phase 3): re-enable tst_databasemanager and tst_sensoryreportsource
# after 3b/3c/3d rewrite — both depend on the old DatabaseManager::open(":memory:")
# API and won't compile against the new Postgres-backed signature.

SUBDIRS += \
    tst_tpmcalculator \
    tst_xmlbuilder \
    tst_zipwriter \
    tst_excelreader \
    tst_sheetprocessors \
    tst_dataprocessor \
    tst_plotengine \
    tst_pptxwriter \
    tst_imageutils \
    tst_reportgenerator \
    tst_sopLoader \
    tst_reportlayout \
    tst_slidecanvasitems \
    tst_layoutcommand \
    tst_identitymanager \
    tst_configloader \
    tst_postgresconnection \
    tst_notificationlistener \
    tst_presencemanager \
    tst_migrationtool

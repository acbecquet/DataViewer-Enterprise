#include <QtTest>
#include "WriteOutcome.h"

using namespace DVE;

class tst_WriteOutcome : public QObject
{
    Q_OBJECT
private slots:
    void success_isSaved();
    void offline_isRetryOffline();
    void versionMismatch_isRetryConflict();
    void rowDeleted_isRetryConflict();
    void uniqueViolation_isRetryError();
    void otherError_isRetryError();
};

void tst_WriteOutcome::success_isSaved()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::Success)),         int(LoadSavePolicy::Saved)); }

void tst_WriteOutcome::offline_isRetryOffline()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::OfflineReadOnly)), int(LoadSavePolicy::RetryOffline)); }

void tst_WriteOutcome::versionMismatch_isRetryConflict()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::VersionMismatch)), int(LoadSavePolicy::RetryConflict)); }

void tst_WriteOutcome::rowDeleted_isRetryConflict()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::RowDeleted)),      int(LoadSavePolicy::RetryConflict)); }

void tst_WriteOutcome::uniqueViolation_isRetryError()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::UniqueViolation)), int(LoadSavePolicy::RetryError)); }

void tst_WriteOutcome::otherError_isRetryError()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::OtherError)),      int(LoadSavePolicy::RetryError)); }

QTEST_APPLESS_MAIN(tst_WriteOutcome)
#include "tst_writeoutcome.moc"

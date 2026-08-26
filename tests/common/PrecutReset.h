#pragma once
// v3 Phase 3d (index D-3d-8): reset the PRIVATE rehearsal database
// (dve_test_precut) back to its pre-cutover shape. The rehearsal harness's
// cutover slots legitimately leave it cut over for the rest of the container's
// lifetime, so ANY test that needs the pre-cutover shape (tst_v3longformat's
// own initTestCase; tst_databasemanager's connect-probe slot) must un-cut
// first rather than assume suite order.
//
// This un-cut exists ONLY for the private rehearsal database - production
// rollback is the D7 backup, never a rename-back (post-cutover writes live
// only in the long tables and a rename-back would strand them). Rename-back
// is sound here because views bind base tables by OID: the renamed-back
// tables keep their OIDs, so data_rows_v / samples_v keep reading them.

#include <QtSql>
#include <QString>

namespace DVE {
namespace TestSeed {

// Returns false only when the database IS cut over and the un-cut failed
// (leaving it in an indeterminate state a caller must not ignore).
inline bool uncutPrecutDatabase(QSqlDatabase& db, QString* outError = nullptr)
{
    QSqlQuery q(db);
    if (!q.exec(QStringLiteral(
            "SELECT (c.relkind = 'v') FROM pg_class c "
            "WHERE c.oid = to_regclass('data_rows')"))) {
        if (outError) *outError = q.lastError().text();
        return false;
    }
    if (!q.next() || !q.value(0).toBool())
        return true;   // already pre-cutover (or no relation at all)

    const char* const steps[] = {
        "DROP VIEW data_rows",
        "DROP VIEW samples",
        "ALTER TABLE data_rows_pre_v3 RENAME TO data_rows",
        "ALTER TABLE samples_core RENAME TO samples",
        "DELETE FROM schema_meta WHERE key = 'v3_long_format_cutover'",
    };
    for (const char* sql : steps) {
        if (!q.exec(QLatin1String(sql))) {
            if (outError)
                *outError = QLatin1String(sql) + QStringLiteral(" -- ")
                          + q.lastError().text();
            return false;
        }
    }
    return true;
}

} // namespace TestSeed
} // namespace DVE

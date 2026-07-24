#include "DatabaseManager.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "ConfigLoader.h"
#include "OfflineSnapshot.h"
#include "RawGridJson.h"
#include "../utils/OutputPaths.h"   // v2.5.0 RC4: nextSuffixedName for auto-suffix wrappers
#include "DatabaseOps.h"

#include <QDebug>
#include <QSqlQuery>
#include <QSqlError>
#include <QSqlRecord>
#include <QSqlDatabase>
#include <QVariant>
#include <QUuid>
#include <QDateTime>
#include <QFile>
#include <QFileInfo>
#include <QDir>
#include <QStandardPaths>
#include <QRectF>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>
#include <QRegularExpression>
#include <QCryptographicHash>

namespace DVE {

// ── settings keys ───────────────────────────────────────────────────────────
// Cumulative-layout JSON for the multi-session sensory radar lives in the
// settings table under this key. Defined once so save and load can't drift.
namespace {
constexpr const char* kCumulativeLayoutKey = "sensory.cumulative_layout";

// Postgres SQLSTATE for unique_violation. Mapped to WriteResult::UniqueViolation
// by every tryWrite* method's INSERT branch.
constexpr const char* kSqlStateUniqueViolation = "23505";

// v2.0.2 H8: write a BLOB to the session ImageCache directory using a
// content-hash filename. When the same image (e.g. shared device-photo)
// is loaded into multiple samples or sessions, the second materialise
// short-circuits since the on-disk file already exists with the same
// hash-prefixed name. Skipping the write avoids one syscall per duplicate
// blob during loadFile / loadSensorySession.
QString materialiseImageBlob(const QString& cacheDir,
                             const QByteArray& blob,
                             const QString& fileName)
{
    const QString md5 = QString::fromLatin1(
        QCryptographicHash::hash(blob, QCryptographicHash::Md5).toHex());
    const QString path = cacheDir + "/" + md5 + "_" + fileName;
    QFile existing(path);
    if (existing.exists() && existing.size() == blob.size()) {
        // Content-hash + size match — reuse without rewriting.
        return path;
    }
    QFile out(path);
    if (out.open(QIODevice::WriteOnly)) {
        out.write(blob);
        out.close();
    }
    return path;
}
}

// --- ctor / dtor / open / close ---------------------------------------------
DatabaseManager::DatabaseManager(QObject* parent)
    : QObject(parent), m_pg(new PostgresConnection(this)) {}

DatabaseManager::~DatabaseManager() { close(); }

bool DatabaseManager::open(const DbConfig& cfg, IdentityManager* identity) {
    m_identity = identity;
    m_cfg      = cfg;
    m_haveCfg  = true;
    m_online = false;
    if (!m_pg->open(cfg)) {
        m_lastError = QStringLiteral("open(connect): ") + m_pg->lastError();
        m_open = false;
        return false;
    }
    m_lastError.clear();
    m_open = true;
    // Default state on successful open: online + no snapshot. ConnectionMonitor
    // (C3) flips m_online to false when ping detects the server is unreachable.
    m_online = true;
    // Reconcile additive columns a pre-migration live DB may be missing (e.g.
    // data_rows.puffing_regime / tests.raw_grid). Idempotent + best-effort:
    // a failure here is logged but must not block an otherwise-good connection.
    ensureSchema();
    return true;
}

bool DatabaseManager::reopen() {
    if (!m_haveCfg) {
        m_lastError = QStringLiteral("reopen: open() was never called successfully");
        return false;
    }
    if (m_pg) m_pg->close();
    m_open   = false;
    m_online = false;
    if (!m_pg->open(m_cfg)) {
        m_lastError = QStringLiteral("reopen(connect): ") + m_pg->lastError();
        return false;
    }
    m_lastError.clear();
    m_open   = true;
    m_online = true;
    ensureSchema();   // self-heal additive columns on reconnect too (see open())
    return true;
}

void DatabaseManager::close() {
    m_online = false;
    if (m_pg) m_pg->close();
    m_open = false;
}

bool DatabaseManager::isOpen() const {
    return m_open && m_pg && m_pg->isOpen();
}

// Additive, idempotent schema reconciliation — see the header for rationale.
// Self-heals a live DB created from an older init.sql by adding any missing
// post-baseline additive column. Never drops/renames, so it cannot lose data.
void DatabaseManager::ensureSchema() {
    if (!isOpen()) return;

    // Every column added by a migration AFTER the original v2.0 baseline must
    // be listed here so a DB created before that migration self-heals on
    // connect. These mirror the inline columns in deploy/postgres/init.sql and
    // the ADD COLUMN migrations under deploy/postgres/migrations/:
    //   - data_rows.puffing_regime  (2026-05-29-v2.2.1-per-row-regime.sql)
    //   - tests.raw_grid            (2026-06-04-v2.x-tests-raw-grid.sql)
    //   - files.added_at            (2026-06-10-files-added-at-identity.sql, F6)
    //   - files.app_version,
    //     sensory_sessions.app_version,
    //     detailed_sensory_sessions.app_version
    //                               (2026-06-11-app-version-stamping.sql, A1 —
    //                                additive nullable; filled by the stamp
    //                                trigger healed below)
    // All are nullable or NOT NULL-with-DEFAULT/additive — adding one never
    // rewrites existing rows destructively (files.added_at backfills existing
    // rows with the migration time via its DEFAULT now()).
    struct AdditiveColumn { const char* table; const char* column; const char* ddl; };
    static const AdditiveColumn kAdditiveColumns[] = {
        { "data_rows", "puffing_regime",
          "ALTER TABLE data_rows ADD COLUMN IF NOT EXISTS puffing_regime TEXT" },
        { "tests", "raw_grid",
          "ALTER TABLE tests ADD COLUMN IF NOT EXISTS raw_grid JSONB" },
        { "tests", "from_inferred_schema",
          "ALTER TABLE tests ADD COLUMN IF NOT EXISTS from_inferred_schema "
          "INTEGER NOT NULL DEFAULT 0" },
        { "files", "added_at",
          "ALTER TABLE files ADD COLUMN IF NOT EXISTS added_at "
          "TIMESTAMPTZ NOT NULL DEFAULT now()" },
        { "files", "app_version",
          "ALTER TABLE files ADD COLUMN IF NOT EXISTS app_version TEXT" },
        { "sensory_sessions", "app_version",
          "ALTER TABLE sensory_sessions ADD COLUMN IF NOT EXISTS app_version TEXT" },
        { "detailed_sensory_sessions", "app_version",
          "ALTER TABLE detailed_sensory_sessions ADD COLUMN IF NOT EXISTS app_version TEXT" },
    };

    QSqlDatabase& db = m_pg->queryDb();
    for (const auto& col : kAdditiveColumns) {
        const QString table  = QString::fromLatin1(col.table);
        const QString column = QString::fromLatin1(col.column);

        // Cheap catalog lookup first: only ALTER when the column is genuinely
        // missing, so the common (already-present) path never takes the brief
        // ACCESS EXCLUSIVE lock ALTER TABLE acquires — which would otherwise
        // stall other users' writes on the shared live DB on every connect.
        // CAST(? AS regclass) resolves the table name through search_path
        // exactly as the app's own unqualified queries do.
        QSqlQuery check(db);
        check.prepare("SELECT 1 FROM pg_attribute "
                      "WHERE attrelid = CAST(? AS regclass) "
                      "AND attname = ? AND NOT attisdropped");
        check.addBindValue(table);
        check.addBindValue(column);
        if (check.exec()) {
            if (check.next())
                continue;   // already present — nothing to do
        } else {
            // Catalog read failed (e.g. table genuinely absent). An ALTER would
            // fail too; log and move on rather than abort an otherwise-good
            // connection.
            logDebug(QStringLiteral("ensureSchema: cannot inspect %1.%2: %3")
                         .arg(table, column, check.lastError().text()));
            continue;
        }

        QSqlQuery alter(db);
        if (alter.exec(QString::fromLatin1(col.ddl))) {
            logDebug(QStringLiteral("ensureSchema: added missing column %1.%2")
                         .arg(table, column));
        } else {
            logDebug(QStringLiteral("ensureSchema: could not add %1.%2: %3")
                         .arg(table, column, alter.lastError().text()));
        }
    }

    // ── app_version stamping (v2.4.2 A1) ─────────────────────────────────────
    // Server-side era stamp: a BEFORE INSERT OR UPDATE trigger fills app_version
    // from the connection's application_name (set to "DataViewer/<ver>" in the
    // connect string by pgSharedConnectOptions()). FAIL-SAFE by construction:
    // nullable column, NO CHECK, NEVER RAISE, reads current_setting(...,true)
    // (missing_ok), and only FILLS a NULL via COALESCE — never blanks a good
    // stamp, so a reconnected NULL-name session can't erase an era. Old clients
    // send no application_name -> rows stay NULL -> shown as "pre-v2.4.2".
    // The function body is an idempotent in-place swap (CREATE OR REPLACE, no
    // data lock); the trigger creation is catalog-guarded on pg_trigger so the
    // already-healed path takes no DDL lock. Best-effort, never thrown.
    {
        QSqlQuery fn(db);
        if (!fn.exec(QStringLiteral(
                "CREATE OR REPLACE FUNCTION dve_stamp_app_version() "
                "RETURNS TRIGGER AS $$ BEGIN "
                "  NEW.app_version := COALESCE("
                "      NEW.app_version, "
                "      NULLIF(current_setting('application_name', true), '')); "
                "  RETURN NEW; "
                "END; $$ LANGUAGE plpgsql;"))) {
            logDebug(QStringLiteral("ensureSchema: could not create "
                         "dve_stamp_app_version: %1").arg(fn.lastError().text()));
        }
        static const char* kStampTables[] = {
            "files", "sensory_sessions", "detailed_sensory_sessions" };
        for (const char* t : kStampTables) {
            const QString table = QString::fromLatin1(t);
            const QString trg =
                QStringLiteral("trg_%1_stamp_app_version").arg(table);
            QSqlQuery chk(db);
            chk.prepare(QStringLiteral(
                "SELECT 1 FROM pg_trigger "
                "WHERE tgrelid = CAST(? AS regclass) AND tgname = ?"));
            chk.addBindValue(table);
            chk.addBindValue(trg);
            bool known = false, present = false;
            if (chk.exec()) { known = true; present = chk.next(); }
            else logDebug(QStringLiteral("ensureSchema: cannot inspect %1 stamp "
                         "trigger: %2").arg(table, chk.lastError().text()));
            if (known && !present) {
                QSqlQuery mk(db);
                if (mk.exec(QStringLiteral(
                        "CREATE TRIGGER trg_%1_stamp_app_version "
                        "BEFORE INSERT OR UPDATE ON %1 "
                        "FOR EACH ROW EXECUTE FUNCTION dve_stamp_app_version()")
                        .arg(table))) {
                    logDebug(QStringLiteral("ensureSchema: app_version stamp "
                                 "trigger added on %1").arg(table));
                } else {
                    logDebug(QStringLiteral("ensureSchema: could not add %1 "
                                 "stamp trigger: %2")
                                 .arg(table, mk.lastError().text()));
                }
            }
        }
    }

    // ── files: heal to the F6 (file_path, added_at) versioned identity ───────
    // USER REQUIREMENT (v2.5.0): re-adding the same .xlsx later must mint a NEW
    // historical row, not overwrite the existing one. The files.added_at column
    // is healed by the additive loop above; here we swap uniqueness from the
    // legacy single-column UNIQUE(file_path) (idx_files_path) to the composite
    // UNIQUE(file_path, added_at) so two versions of one path coexist. Guard on
    // the presence of the new composite index so the common (already-healed)
    // path takes no DDL lock. Best-effort: every step logs on failure and is
    // never thrown — mirrors the sensory_header_presets and additive-column
    // healing. Requires files.added_at to already exist (the additive loop ran
    // first); if a legacy DB couldn't add it the index build below fails and is
    // logged, leaving the legacy unique index in place (still correct, just
    // un-versioned) until the next connect retries.
    {
        bool idxKnown = false, hasCompositeIdx = false;
        {
            QSqlQuery idxChk(db);
            if (idxChk.exec(QStringLiteral(
                    "SELECT 1 FROM pg_indexes "
                    "WHERE tablename = 'files' "
                    "AND indexname = 'uq_files_path_added'"))) {
                idxKnown = true;
                hasCompositeIdx = idxChk.next();
            } else {
                logDebug(QStringLiteral("ensureSchema: cannot inspect files "
                             "indexes: %1").arg(idxChk.lastError().text()));
            }
        }
        if (idxKnown && !hasCompositeIdx) {
            // Best-effort, each step independent: build the composite unique
            // index FIRST, then drop the legacy single-column unique index +
            // any old constraint form. Ordering the create before the drop
            // means a path with one row per version is never momentarily
            // un-protected. IF EXISTS / IF NOT EXISTS keep it idempotent.
            struct Step { const char* what; const char* ddl; };
            static const Step kSteps[] = {
                { "create composite unique index",
                  "CREATE UNIQUE INDEX IF NOT EXISTS uq_files_path_added "
                  "ON files(file_path, added_at)" },
                { "drop legacy single-column unique index",
                  "DROP INDEX IF EXISTS idx_files_path" },
                { "drop legacy single-column unique constraint",
                  "ALTER TABLE files DROP CONSTRAINT IF EXISTS files_file_path_key" },
            };
            bool createdComposite = false;
            for (const auto& step : kSteps) {
                QSqlQuery s(db);
                if (!s.exec(QString::fromLatin1(step.ddl))) {
                    logDebug(QStringLiteral("ensureSchema: files %1 failed: %2")
                                 .arg(QString::fromLatin1(step.what),
                                      s.lastError().text()));
                } else if (qstrcmp(step.what, "create composite unique index") == 0) {
                    createdComposite = true;
                }
            }
            if (createdComposite)
                logDebug(QStringLiteral("ensureSchema: files uniqueness swapped "
                             "to (file_path, added_at) — versioned re-adds (F6)"));
        }
    }

    // ── dve_commit_cell_json: heal to numeric JSONB storage (DATAVIEWER-4) ───
    // ROOT of the original "scores reset to 5" revert. The LiveSync per-cell
    // commit function historically stored every value via to_jsonb($2::text),
    // so a numeric score streamed live landed in the JSONB as a STRING ("7.0"),
    // and the C++ reader's QJsonValue::toDouble(default) returned the DEFAULT
    // for a string — every live-streamed score reverted on the next fresh load.
    // The new body stores numeric-looking values as JSON NUMBERS (text fallback
    // for names / free text); the tolerant jsonToDouble reader repairs the rows
    // already corrupted in the live DB.
    //
    // We REPLACE the CANONICAL 7-argument OCC form (p_expected_version, from
    // 2026-05-17-v2.0.2-fixes.sql) — the one LiveSyncWorker actually dispatches
    // (it binds 7 positional args). The legacy 6-arg form in init.sql is an
    // unused overload and is intentionally left alone (touching the scalar
    // dve_commit_cell is out of scope; see tst_storedfns's arg-count probe).
    // Guarded on prosrc so the common (already-healed) path takes no DDL: only
    // CREATE OR REPLACE when the 7-arg body still lacks the numeric coercion.
    // Best-effort, catalog-guarded, idempotent, never thrown — same contract as
    // the healing blocks above. CREATE OR REPLACE with an unchanged signature is
    // a cheap in-place body swap.
    {
        bool needsHeal = false, probed = false;
        {
            QSqlQuery chk(db);
            // The 7-arg OCC overload specifically (pronargs = 7). If a legacy DB
            // somehow lacks it, the 6-arg form is unused by the live worker, so
            // there is nothing to heal — leave it.
            if (chk.exec(QStringLiteral(
                    "SELECT prosrc FROM pg_proc "
                    "WHERE proname = 'dve_commit_cell_json' AND pronargs = 7"))) {
                probed = true;
                if (chk.next()) {
                    const QString src = chk.value(0).toString();
                    // Marker: the healed body contains 'to_jsonb($2::numeric'.
                    // Its absence means the old text-only body is still live.
                    needsHeal = !src.contains(QStringLiteral("to_jsonb($2::numeric"));
                }
            } else {
                logDebug(QStringLiteral("ensureSchema: cannot inspect "
                             "dve_commit_cell_json: %1").arg(chk.lastError().text()));
            }
        }
        if (probed && needsHeal) {
            QSqlQuery repl(db);
            // Same name + same 7-arg signature + same OCC behavior as
            // 2026-05-17-v2.0.2-fixes.sql; ONLY the to_jsonb expression changes
            // (numeric-looking values -> JSON number). Mirror of
            // deploy/postgres/migrations/2026-06-10-commit-cell-json-numeric.sql.
            // NOTE: the dynamic-SQL string passed to format() MUST be ONE
            // literal each. PostgreSQL does NOT concatenate two quoted literals
            // separated only by spaces on the same line (it needs a newline
            // between them), so the format() argument is kept as a single long
            // string rather than the multi-line broken-up form the .sql files
            // can use. '' is the escaped single-quote inside the SQL literal.
            const QString ddl = QStringLiteral(
                "CREATE OR REPLACE FUNCTION dve_commit_cell_json("
                "    p_table TEXT, p_row_id BIGINT, p_path_text TEXT,"
                "    p_path_arr TEXT[], p_value TEXT, p_uuid TEXT,"
                "    p_expected_version INT DEFAULT NULL"
                ") RETURNS BOOLEAN AS $fn$ "
                "DECLARE affected INT; "
                "BEGIN "
                "    PERFORM set_config('dve.live_column', 'json_path:' || p_path_text, true); "
                "    PERFORM set_config('dve.live_value',  p_value, true); "
                "    IF p_expected_version IS NULL THEN "
                "        EXECUTE format('UPDATE %I SET json_data = jsonb_set(json_data, $1, "
                "(CASE WHEN $2 ~ ''^-?[0-9]+(\\.[0-9]+)?$'' THEN to_jsonb($2::numeric) "
                "ELSE to_jsonb($2::text) END)::jsonb, true), updated_by = $3 WHERE id = $4', "
                "p_table) USING p_path_arr, p_value, p_uuid, p_row_id; "
                "    ELSE "
                "        EXECUTE format('UPDATE %I SET json_data = jsonb_set(json_data, $1, "
                "(CASE WHEN $2 ~ ''^-?[0-9]+(\\.[0-9]+)?$'' THEN to_jsonb($2::numeric) "
                "ELSE to_jsonb($2::text) END)::jsonb, true), updated_by = $3 "
                "WHERE id = $4 AND version = $5', "
                "p_table) USING p_path_arr, p_value, p_uuid, p_row_id, p_expected_version; "
                "    END IF; "
                "    GET DIAGNOSTICS affected = ROW_COUNT; "
                "    RETURN affected > 0; "
                "END; $fn$ LANGUAGE plpgsql;");
            if (repl.exec(ddl)) {
                logDebug(QStringLiteral("ensureSchema: dve_commit_cell_json healed "
                             "to numeric JSONB storage (DATAVIEWER-4)"));
            } else {
                logDebug(QStringLiteral("ensureSchema: could not heal "
                             "dve_commit_cell_json: %1").arg(repl.lastError().text()));
            }
        }
    }

    // ── dve_commit_cell_json 6-arg overload: heal to numeric JSONB (v2.4.2) ──
    // v2.4.1 healed only the 7-arg OCC form the live worker dispatches. v2.4.2
    // also heals the 6-arg legacy overload (init.sql, no p_expected_version) so
    // any old/alt caller — or future code — that hits the 6-arg signature can't
    // re-introduce string-typed scores. Same catalog-guard / best-effort
    // contract as the 7-arg block above (prosrc marker, pronargs=6). The healed
    // body relies on the BEFORE UPDATE bump_version trigger for version/updated_at
    // (mirrors the 7-arg healed body), changing ONLY the to_jsonb coercion.
    {
        bool needsHeal = false, probed = false;
        {
            QSqlQuery chk(db);
            if (chk.exec(QStringLiteral(
                    "SELECT prosrc FROM pg_proc "
                    "WHERE proname = 'dve_commit_cell_json' AND pronargs = 6"))) {
                probed = true;
                if (chk.next())
                    needsHeal = !chk.value(0).toString()
                                     .contains(QStringLiteral("to_jsonb($2::numeric"));
            } else {
                logDebug(QStringLiteral("ensureSchema: cannot inspect 6-arg "
                             "dve_commit_cell_json: %1").arg(chk.lastError().text()));
            }
        }
        if (probed && needsHeal) {
            QSqlQuery repl(db);
            const QString ddl = QStringLiteral(
                "CREATE OR REPLACE FUNCTION dve_commit_cell_json("
                "    p_table TEXT, p_row_id BIGINT, p_path_text TEXT,"
                "    p_path_arr TEXT[], p_value TEXT, p_uuid TEXT"
                ") RETURNS BOOLEAN AS $fn$ "
                "DECLARE affected INT; "
                "BEGIN "
                "    PERFORM set_config('dve.live_column', 'json_path:' || p_path_text, true); "
                "    PERFORM set_config('dve.live_value',  p_value, true); "
                "    EXECUTE format('UPDATE %I SET json_data = jsonb_set(json_data, $1, "
                "(CASE WHEN $2 ~ ''^-?[0-9]+(\\.[0-9]+)?$'' THEN to_jsonb($2::numeric) "
                "ELSE to_jsonb($2::text) END)::jsonb, true), updated_by = $3 WHERE id = $4', "
                "p_table) USING p_path_arr, p_value, p_uuid, p_row_id; "
                "    GET DIAGNOSTICS affected = ROW_COUNT; "
                "    RETURN affected > 0; "
                "END; $fn$ LANGUAGE plpgsql;");
            if (repl.exec(ddl)) {
                logDebug(QStringLiteral("ensureSchema: 6-arg dve_commit_cell_json "
                             "healed to numeric JSONB storage"));
            } else {
                logDebug(QStringLiteral("ensureSchema: could not heal 6-arg "
                             "dve_commit_cell_json: %1").arg(repl.lastError().text()));
            }
        }
    }

    // ── notify_row_change: maintenance-suppression early-return (v2.4.2 R4) ──
    // Heal live DBs so a bulk normalize (dve.maintenance='1') doesn't NOTIFY-
    // storm clients. Guard on prosrc so the already-healed path takes no work.
    {
        bool needsHeal = false, probed = false;
        QSqlQuery chk(db);
        if (chk.exec(QStringLiteral(
                "SELECT prosrc FROM pg_proc WHERE proname = 'notify_row_change'"))) {
            probed = true;
            if (chk.next())
                needsHeal = !chk.value(0).toString()
                                 .contains(QStringLiteral("dve.maintenance"));
        } else {
            logDebug(QStringLiteral("ensureSchema: cannot inspect notify_row_change: %1")
                         .arg(chk.lastError().text()));
        }
        if (probed && needsHeal) {
            QSqlQuery repl(db);
            const QString ddl = QStringLiteral(
                "CREATE OR REPLACE FUNCTION notify_row_change() RETURNS TRIGGER AS $fn$ "
                "DECLARE payload JSONB; BEGIN "
                "  IF current_setting('dve.maintenance', true) = '1' THEN RETURN NULL; END IF; "
                "  payload := jsonb_build_object('table', TG_TABLE_NAME, 'op', TG_OP, "
                "    'id', COALESCE(NEW.id, OLD.id), "
                "    'updated_by', COALESCE(NEW.updated_by, OLD.updated_by)); "
                "  IF current_setting('dve.live_column', true) <> '' THEN "
                "    payload := payload || jsonb_build_object('column', "
                "      current_setting('dve.live_column', true), 'new_value', "
                "      current_setting('dve.live_value', true)); END IF; "
                "  PERFORM pg_notify('dataviewer_changes', payload::text); "
                "  RETURN NULL; END; $fn$ LANGUAGE plpgsql;");
            if (repl.exec(ddl))
                logDebug(QStringLiteral("ensureSchema: notify_row_change healed "
                             "(maintenance suppression)"));
            else
                logDebug(QStringLiteral("ensureSchema: could not heal "
                             "notify_row_change: %1").arg(repl.lastError().text()));
        }
    }

    // ── dve_normalize_legacy_json: create-or-replace heal (v2.4.2 R4) ────────
    // The legacy-score normalizer rewrites numeric-looking STRING scores to JSON
    // numbers across both session tables. Unconditionally CREATE OR REPLACE it
    // on every connect (a metadata-only op that touches no data) so live DBs and
    // the test container always carry the current body. Best-effort: log + carry
    // on if it fails (the one-time run below — Task 3 — will simply find nothing
    // to call). Body is byte-identical to init.sql's, incl. ORDER BY elem.ord so
    // sample order is preserved through the rewrite.
    {
        QSqlQuery repl(db);
        const QString ddl = QStringLiteral(
            "CREATE OR REPLACE FUNCTION dve_normalize_legacy_json() RETURNS INTEGER AS $fn$ "
            "DECLARE "
            "    keys TEXT[] := ARRAY["
            "        'Burnt Taste','Vapor Volume','Overall Flavor','Smoothness','Overall Liking',"
            "        'Burn Taste','Flavor Intensity','Throat Irritation','Nasal Irritation',"
            "        'Vapor Quality Overall','Cough','Volume Consistency',"
            "        'Performance Consistency','Vapor Temperature','Vapor vs Oil']; "
            "    total INTEGER := 0; n INTEGER; tbl TEXT; "
            "BEGIN "
            "    SET LOCAL statement_timeout = 0; "
            // NOTE: SQL '--' line-comments are intentionally omitted in this
            // single-line DDL literal — without a newline they'd comment out the
            // rest of the function body. The init.sql/migration copies (multi-line)
            // do carry the '-- tx-local; suppresses NOTIFY' annotation on this line.
            "    PERFORM set_config('dve.maintenance', '1', true); "
            "    FOREACH tbl IN ARRAY ARRAY['sensory_sessions','detailed_sensory_sessions'] LOOP "
            "        EXECUTE format($q$ "
            "            UPDATE %I s SET json_data = jsonb_set( "
            "                s.json_data, '{samples}', "
            "                (SELECT jsonb_agg( "
            "                    COALESCE( "
            "                      (SELECT jsonb_object_agg(kv.key, "
            "                          CASE WHEN kv.key = ANY($1) "
            "                                AND jsonb_typeof(kv.value) = 'string' "
            "                                AND (kv.value #>> '{}') ~ '^-?[0-9]+(\\.[0-9]+)?$' "
            "                               THEN to_jsonb((kv.value #>> '{}')::numeric) "
            "                               ELSE kv.value END) "
            "                       FROM jsonb_each(elem.value) kv), "
            "                      elem.value) "
            "                    ORDER BY elem.ord) "
            "                 FROM jsonb_array_elements(s.json_data->'samples') "
            "                      WITH ORDINALITY AS elem(value, ord))) "
            "            WHERE jsonb_typeof(s.json_data->'samples') = 'array' "
            "              AND EXISTS ( "
            "                SELECT 1 FROM jsonb_array_elements(s.json_data->'samples') e, "
            "                              jsonb_each(e.value) kv "
            "                WHERE kv.key = ANY($1) "
            "                  AND jsonb_typeof(kv.value) = 'string' "
            "                  AND (kv.value #>> '{}') ~ '^-?[0-9]+(\\.[0-9]+)?$') "
            "        $q$, tbl) USING keys; "
            "        GET DIAGNOSTICS n = ROW_COUNT; "
            "        total := total + n; "
            "    END LOOP; "
            "    RETURN total; "
            "END; $fn$ LANGUAGE plpgsql;");
        if (repl.exec(ddl))
            logDebug(QStringLiteral("ensureSchema: dve_normalize_legacy_json "
                         "create-or-replace OK"));
        else
            logDebug(QStringLiteral("ensureSchema: could not create "
                         "dve_normalize_legacy_json: %1").arg(repl.lastError().text()));
    }

    // ── v2.4.2 R4: one-time legacy-score normalize (advisory-locked, gated) ──
    // Serializes concurrent first-connects (advisory xact lock) so N clients
    // don't each run the full-table normalize. schema_meta-gated so it runs
    // ONCE; the nightly pg_cron job (init.sql) handles ongoing convergence.
    // Wrapped in an explicit transaction because ensureSchema is otherwise
    // autocommit and pg_advisory_xact_lock would release immediately. Strictly
    // best-effort: every step logs + carries on, nothing throws.
    if (db.transaction()) {
        bool ok = true;
        { QSqlQuery lk(db);
          ok = lk.exec(QStringLiteral("SELECT pg_advisory_xact_lock(4242002)"));
          if (!ok)
              logDebug(QStringLiteral("ensureSchema: normalize advisory lock failed: %1")
                           .arg(lk.lastError().text())); }
        bool gateKnown = false, alreadyRun = false;
        if (ok) {
            QSqlQuery g(db);
            gateKnown = g.exec(QStringLiteral(
                "SELECT 1 FROM schema_meta WHERE key='v242_legacy_score_normalize'"));
            if (gateKnown) alreadyRun = g.next();
            else
                logDebug(QStringLiteral("ensureSchema: normalize gate probe failed: %1")
                             .arg(g.lastError().text()));
        }
        if (ok && gateKnown && !alreadyRun) {
            QSqlQuery n(db);
            if (n.exec(QStringLiteral("SELECT dve_normalize_legacy_json()"))) {
                QSqlQuery m(db);
                if (!m.exec(QStringLiteral(
                        "INSERT INTO schema_meta(key,value) "
                        "VALUES ('v242_legacy_score_normalize', now()::text) "
                        "ON CONFLICT (key) DO NOTHING")))
                    logDebug(QStringLiteral("ensureSchema: could not record legacy-score "
                                 "normalize marker: %1").arg(m.lastError().text()));
                logDebug(QStringLiteral("ensureSchema: legacy-score normalize ran (one-time)"));
            } else {
                logDebug(QStringLiteral("ensureSchema: legacy-score normalize failed: %1")
                             .arg(n.lastError().text()));
            }
        }
        if (!db.commit()) {
            db.rollback();
            logDebug(QStringLiteral("ensureSchema: normalize tx commit failed (rolled back)"));
        }
    } else {
        logDebug(QStringLiteral("ensureSchema: could not begin normalize tx (skipped)"));
    }

    // ── sensory_header_presets: heal to the DATAVIEWER-2 test-scoped shape ───
    // The sample-name dropdown becomes scoped to the current test, so the
    // shared preset pool grows a `test_name` column and its uniqueness moves
    // from (kind, value) to a UNIQUE expression index on
    // (kind, value, COALESCE(test_name, '')) — letting the same value live
    // once per test (and once globally, where test_name IS NULL). A live DB
    // created from an init.sql that predates this table (the test container is
    // one such case) must self-heal on connect, exactly like the additive
    // columns above. All three reconciliations are catalog-guarded so the
    // common (already-healed) path takes no DDL lock, idempotent, and
    // best-effort: any failure is logged and skipped, never thrown.
    {
        // (1) Create the table fresh in the new shape if it's absent. The
        //     CAST(... AS regclass) form would *throw* on a missing relation,
        //     so use to_regclass(), which returns NULL instead.
        bool tableExists = false;
        bool tableKnown  = false;
        {
            QSqlQuery chk(db);
            if (chk.exec(QStringLiteral(
                    "SELECT to_regclass('public.sensory_header_presets') "
                    "IS NOT NULL"))) {
                tableKnown = true;
                if (chk.next()) tableExists = chk.value(0).toBool();
            } else {
                logDebug(QStringLiteral("ensureSchema: cannot probe "
                             "sensory_header_presets: %1")
                             .arg(chk.lastError().text()));
            }
        }

        if (tableKnown && !tableExists) {
            QSqlQuery create(db);
            if (create.exec(QStringLiteral(
                    "CREATE TABLE IF NOT EXISTS sensory_header_presets ("
                    "    id          BIGSERIAL PRIMARY KEY,"
                    "    kind        TEXT NOT NULL CHECK (kind IN "
                    "                    ('test_name', 'media', 'sample_name')),"
                    "    value       TEXT NOT NULL CHECK (length(trim(value)) > 0),"
                    "    test_name   TEXT,"
                    "    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),"
                    "    created_by  TEXT,"
                    "    version     INTEGER NOT NULL DEFAULT 1,"
                    "    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),"
                    "    updated_by  TEXT)"))) {
                logDebug(QStringLiteral("ensureSchema: created "
                             "sensory_header_presets (test-scoped shape)"));
            } else {
                logDebug(QStringLiteral("ensureSchema: could not create "
                             "sensory_header_presets: %1")
                             .arg(create.lastError().text()));
            }

            // Register the table with the shared version-bump + NOTIFY trigger
            // pools, uniform with the rest of the schema (see the 2026-05-20
            // migration). The trigger functions (bump_version, notify_row_change)
            // ship in init.sql; on a DB so old they're absent this CREATE
            // TRIGGER fails — that's a consistency nicety, NOT load-bearing for
            // DV-2 (saveSensoryHeaderPresets only INSERTs ON CONFLICT DO
            // NOTHING, never UPDATEs), so we log and carry on. Idempotent via
            // DROP TRIGGER IF EXISTS.
            QSqlQuery trg(db);
            if (!trg.exec(QStringLiteral(
                    "DO $$ BEGIN"
                    "  EXECUTE 'DROP TRIGGER IF EXISTS "
                    "      trg_sensory_header_presets_bump_version "
                    "      ON sensory_header_presets;"
                    "      CREATE TRIGGER trg_sensory_header_presets_bump_version "
                    "      BEFORE UPDATE ON sensory_header_presets "
                    "      FOR EACH ROW EXECUTE FUNCTION bump_version();';"
                    "  EXECUTE 'DROP TRIGGER IF EXISTS "
                    "      trg_sensory_header_presets_notify "
                    "      ON sensory_header_presets;"
                    "      CREATE TRIGGER trg_sensory_header_presets_notify "
                    "      AFTER INSERT OR UPDATE OR DELETE ON sensory_header_presets "
                    "      FOR EACH ROW EXECUTE FUNCTION notify_row_change();';"
                    "END $$;"))) {
                logDebug(QStringLiteral("ensureSchema: sensory_header_presets "
                             "trigger registration skipped: %1")
                             .arg(trg.lastError().text()));
            }
        }

        // (2) Legacy table already present but missing test_name → add it.
        //     Catalog-guard so the healed path takes no ALTER lock.
        if (tableExists) {
            QSqlQuery colChk(db);
            colChk.prepare(QStringLiteral(
                "SELECT 1 FROM pg_attribute "
                "WHERE attrelid = CAST('sensory_header_presets' AS regclass) "
                "AND attname = 'test_name' AND NOT attisdropped"));
            bool hasCol = false, colKnown = false;
            if (colChk.exec()) { colKnown = true; hasCol = colChk.next(); }
            else {
                logDebug(QStringLiteral("ensureSchema: cannot inspect "
                             "sensory_header_presets.test_name: %1")
                             .arg(colChk.lastError().text()));
            }
            if (colKnown && !hasCol) {
                QSqlQuery addCol(db);
                if (addCol.exec(QStringLiteral(
                        "ALTER TABLE sensory_header_presets "
                        "ADD COLUMN IF NOT EXISTS test_name TEXT"))) {
                    logDebug(QStringLiteral("ensureSchema: added "
                                 "sensory_header_presets.test_name"));
                } else {
                    logDebug(QStringLiteral("ensureSchema: could not add "
                                 "sensory_header_presets.test_name: %1")
                                 .arg(addCol.lastError().text()));
                }
            }
        }

        // (3) Swap uniqueness to the test-scoped expression index. Guard on the
        //     presence of an index whose definition mentions COALESCE(test_name
        //     so the common (already-healed) path neither drops the old
        //     constraint nor rebuilds the index. A table just created above in
        //     (1) has no such index yet, so it falls straight into this block
        //     within the same connect and gets the unique + lookup indexes —
        //     no second pass needed. A legacy table healed in (2) does too.
        bool idxKnown = false, hasCoalesceIdx = false;
        {
            QSqlQuery idxChk(db);
            if (idxChk.exec(QStringLiteral(
                    "SELECT 1 FROM pg_indexes "
                    "WHERE tablename = 'sensory_header_presets' "
                    "AND indexdef LIKE '%COALESCE(test_name%'"))) {
                idxKnown = true;
                hasCoalesceIdx = idxChk.next();
            } else {
                logDebug(QStringLiteral("ensureSchema: cannot inspect "
                             "sensory_header_presets indexes: %1")
                             .arg(idxChk.lastError().text()));
            }
        }
        if (idxKnown && !hasCoalesceIdx) {
            // Best-effort, each step independent: drop the legacy UNIQUE
            // (kind, value) constraint and its companion lookup index, then
            // build the test-scoped unique + lookup indexes. IF EXISTS / IF
            // NOT EXISTS everywhere keeps it idempotent and re-runnable.
            struct Step { const char* what; const char* ddl; };
            static const Step kSteps[] = {
                { "drop legacy unique constraint",
                  "ALTER TABLE sensory_header_presets "
                  "DROP CONSTRAINT IF EXISTS sensory_header_presets_kind_value_key" },
                { "drop legacy lookup index",
                  "DROP INDEX IF EXISTS idx_sensory_header_presets_kind" },
                { "create test-scoped unique index",
                  "CREATE UNIQUE INDEX IF NOT EXISTS uq_shp_kind_value_test "
                  "ON sensory_header_presets (kind, value, COALESCE(test_name, ''))" },
                { "create test-scoped lookup index",
                  "CREATE INDEX IF NOT EXISTS idx_shp_kind_test "
                  "ON sensory_header_presets (kind, test_name)" },
            };
            for (const auto& step : kSteps) {
                QSqlQuery s(db);
                if (!s.exec(QString::fromLatin1(step.ddl))) {
                    logDebug(QStringLiteral("ensureSchema: sensory_header_presets "
                                 "%1 failed: %2")
                                 .arg(QString::fromLatin1(step.what),
                                      s.lastError().text()));
                }
            }
            logDebug(QStringLiteral("ensureSchema: sensory_header_presets "
                         "uniqueness swapped to (kind, value, COALESCE(test_name,''))"));
        }
    }

    // ── DATAVIEWER-2: one-time backfill of (test -> sample-name) presets ─────
    // So the test-scoped sample-name dropdown is useful from day one, seed it
    // once from existing session history: every distinct sample name that has
    // ever been saved under a test title becomes a `sample_name` preset row
    // scoped to that title. The scope MUST be the *test title*, not the session
    // name — that is exactly what the live save/load path keys on
    // (SensoryPanel/DetailedSensoryPanel pass m_testTitleEdit to
    // saveSensoryHeaderPresets; loadSampleNamesForTest matches the test_name
    // column on that title). Sessions whose json carries a blank/absent
    // test_title are skipped (the live path sends those names to the global
    // NULL pool, which a scoped lookup never reads — backfilling them would add
    // unreachable rows).
    //
    // Gated by a schema_meta marker so it runs ONCE, not on every connect.
    // Idempotent via ON CONFLICT DO NOTHING (matches the test-scoped unique
    // index). Best-effort: every step logs on failure and is never thrown, so a
    // backfill hiccup can't abort an otherwise-good connection. A row whose
    // json lacks 'samples' yields no LATERAL rows, so empty/absent arrays are
    // harmless.
    {
        QSqlDatabase& db = m_pg->queryDb();
        // Defensive: schema_meta ships in init.sql, but a DB old enough to
        // predate it would otherwise make the marker probe error out and
        // re-run the backfill on every connect. Catalog-guard the CREATE so the
        // common (already-present) path emits no `NOTICE: relation already
        // exists` on every connect — matching the catalog-guarded presets
        // healing above. Best-effort: a failed probe just falls through to the
        // CREATE IF NOT EXISTS, which is still a no-op when the table is there.
        {
            QSqlQuery sm(db);
            sm.exec(QStringLiteral("SELECT to_regclass('public.schema_meta')"));
            const bool schemaMetaMissing = !(sm.next() && !sm.value(0).isNull());
            if (schemaMetaMissing) {
                QSqlQuery c(db);
                c.exec(QStringLiteral(
                    "CREATE TABLE IF NOT EXISTS schema_meta (key TEXT PRIMARY KEY, value TEXT)"));
            }
        }

        QSqlQuery g(db);
        const bool gateKnown =
            g.exec(QStringLiteral("SELECT 1 FROM schema_meta "
                                  "WHERE key = 'dv2_sample_name_backfill'"));
        if (!gateKnown) {
            logDebug(QStringLiteral("ensureSchema: sample-name backfill gate probe "
                                    "failed (skipping): %1").arg(g.lastError().text()));
        } else if (!g.next()) {
            // Shared INSERT…SELECT skeleton; only the source table differs.
            const auto backfillFrom = [&](const char* table, const char* label) {
                QSqlQuery b(db);
                const QString sql = QStringLiteral(
                    "INSERT INTO sensory_header_presets "
                    "    (kind, value, test_name, created_by, updated_by) "
                    "SELECT DISTINCT 'sample_name', "
                    "       trim(smp->>'name'), "
                    "       trim(ss.json_data->>'test_title'), "
                    "       'backfill', 'backfill' "
                    "FROM %1 ss "
                    "CROSS JOIN LATERAL "
                    "    jsonb_array_elements(ss.json_data->'samples') AS smp "
                    "WHERE ss.json_data->>'test_title' IS NOT NULL "
                    "  AND length(trim(ss.json_data->>'test_title')) > 0 "
                    "  AND smp->>'name' IS NOT NULL "
                    "  AND length(trim(smp->>'name')) > 0 "
                    "ON CONFLICT (kind, value, COALESCE(test_name, '')) DO NOTHING")
                    .arg(QString::fromLatin1(table));
                if (!b.exec(sql)) {
                    logDebug(QStringLiteral("ensureSchema: %1 sample-name backfill "
                                            "failed: %2")
                                 .arg(QString::fromLatin1(label),
                                      b.lastError().text()));
                }
            };
            backfillFrom("sensory_sessions",          "sensory");
            backfillFrom("detailed_sensory_sessions", "detailed");

            QSqlQuery m(db);
            if (!m.exec(QStringLiteral(
                    "INSERT INTO schema_meta (key, value) "
                    "VALUES ('dv2_sample_name_backfill', now()::text) "
                    "ON CONFLICT (key) DO NOTHING"))) {
                // If the marker can't be written the backfill simply re-runs on
                // the next connect — still safe (ON CONFLICT DO NOTHING), just
                // not free. Log so it's diagnosable.
                logDebug(QStringLiteral("ensureSchema: could not record sample-name "
                                        "backfill marker: %1").arg(m.lastError().text()));
            } else {
                logDebug(QStringLiteral("ensureSchema: sample-name preset backfill "
                                        "completed (one-time)"));
            }
        }
    }
}

// ── Offline mode (Plan C) ───────────────────────────────────────────────────
// Lifetime: m_snapshot is owned by MainWindow; DatabaseManager just holds a
// raw pointer. MainWindow guarantees it outlives every save/read path.
void DatabaseManager::setOfflineSnapshot(OfflineSnapshot* snap) {
    m_snapshot = snap;
}

// Soft state — set by ConnectionMonitor in response to a ping failure /
// reconnect, not by close(). close() also clears it for hygiene.
void DatabaseManager::setOnline(bool b) {
    m_online = b;
}

QString DatabaseManager::currentPath() const { return QString(); }

void DatabaseManager::logDebug(const QString& msg) const {
    qDebug().noquote() << "[DatabaseManager]" << msg;
}


// ============================================================================
//  Hierarchical file storage
// ============================================================================
WriteResult DatabaseManager::tryWriteFile(const FileResult& result) {
    // Delegate to the mutable-ref overload via a local copy. The writeback
    // (post-save id + version) is discarded — callers who need it must
    // pass a mutable reference. Offline guard lives in the mutable overload
    // so this wrapper auto-inherits it.
    FileResult copy = result;
    return tryWriteFile(copy);
}

WriteResult DatabaseManager::tryWriteFile(FileResult& result) {
    // SP4.5 Stage 2a: the RowDeleted/VersionMismatch retry logic moved
    // verbatim into DVE::persistFileOnConnection (DatabaseOps.cpp) so the UI
    // thread and the background PersistWorker share one body. Same semantics.
    return DVE::persistFileOnConnection(m_pg->queryDb(), DVE::writerUuid(m_identity),
                                        m_online, isOpen(), result, &m_lastError);
}

bool DatabaseManager::saveFile(const FileResult& result) {
    // Const-ref delegates to the const-ref tryWriteFile, which itself
    // delegates to the mutable-ref variant via a local copy and discards
    // the writeback — matching the legacy fire-and-forget semantics of
    // this bool shim.
    return tryWriteFile(result) == WriteResult::Success;
}

// --- hasFile ----------------------------------------------------------------
bool DatabaseManager::hasFile(const QString& filePath) const {
    m_lastError.clear();
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return !m_snapshot->loadFileByPath(filePath).filePath.isEmpty();
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("hasFile: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT 1 FROM files WHERE file_path = ? LIMIT 1");
    q.addBindValue(filePath);
    if (!q.exec()) {
        m_lastError = QStringLiteral("hasFile(select): ") + q.lastError().text();
        return false;
    }
    return q.next();
}

// --- loadFile ---------------------------------------------------------------
// Pure read path - no transaction. Walks files -> tests -> samples ->
// data_rows + images. SELECT now pulls id+version so subsequent saves can
// participate in optimistic concurrency.
FileResult DatabaseManager::loadFile(int id) const {
    m_lastError.clear();
    FileResult result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->loadFile(id);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadFile: database not open");
        return result;
    }

    QSqlDatabase& db = m_pg->queryDb();

    // Step 1: file metadata + id + version
    {
        QSqlQuery q(db);
        q.prepare("SELECT id, version, file_path, file_name, template_version "
                  "FROM files WHERE id = ?");
        q.addBindValue(id);
        if (!q.exec()) {
            m_lastError = QStringLiteral("loadFile(SELECT files): ")
                          + q.lastError().text();
            return result;
        }
        if (!q.next()) return result;  // not found, leave result empty
        result.id              = q.value(0).toInt();
        result.version         = q.value(1).toInt();
        result.filePath        = q.value(2).toString();
        result.fileName        = q.value(3).toString();
        result.templateVersion = q.value(4).toString();
    }

    // Step 2: tests. SELECT id+version so the C3 id-aware upsert can UPDATE
    // existing rows in place instead of the legacy DELETE-cascade-rebuild.
    // raw_grid appended last (index 7) so no existing index shifts.
    struct TestInfo {
        qint64 id; int version;
        QString sheetName; QString templateVersion;
        double avgTPM; double stddevTPM; bool isRaw;
        QString rawGrid;   // JSONB column; empty string when NULL (non-raw sheets)
        bool fromInferred; // from_inferred_schema: appended last so no index shifts
    };
    QVector<TestInfo> tests;
    {
        QSqlQuery q(db);
        q.prepare("SELECT id, version, sheet_name, template_version, overall_avg_tpm, "
                  "overall_stddev_tpm, is_raw_table, raw_grid, from_inferred_schema FROM tests "
                  "WHERE file_id = ? ORDER BY sort_order");
        q.addBindValue(id);
        if (!q.exec()) {
            m_lastError = QStringLiteral("loadFile(SELECT tests): ")
                          + q.lastError().text();
            return result;
        }
        while (q.next()) {
            tests.append({q.value(0).toLongLong(), q.value(1).toInt(),
                          q.value(2).toString(), q.value(3).toString(),
                          q.value(4).toDouble(), q.value(5).toDouble(),
                          q.value(6).toInt() != 0,
                          q.value(7).toString(),
                          q.value(8).toInt() != 0});
        }
    }

    // v2.0.2 M1 — load all descendants of this file in bulk and bucket by
    // parent id, rather than issuing one SELECT per test (samples) and one
    // per sample (data_rows / images). The pre-refactor query count was
    // 1 + 1 + N_tests + 2 × N_samples — easily 200+ round trips on a
    // typical file. The new path issues 4 bulk SELECTs filtered by file_id
    // for a total of 5 queries regardless of file shape.

    // Bulk SELECT 1/3 — all samples whose test belongs to this file.
    QHash<qint64, QVector<SampleResult>> samplesByTest;
    QHash<qint64, qint64>                testForSample;
    {
        QSqlQuery q(db);
        q.prepare("SELECT s.id, s.test_id, s.version, s.sample_name, s.sample_id, "
                  "s.date, s.tester, s.media, s.viscosity, s.resistance, "
                  "s.voltage, s.power, s.heating_technology, s.puffing_regime, "
                  "s.initial_oil_mass, s.average_tpm, s.stddev_tpm, "
                  "s.avg_power_density, s.efficiency_percent, "
                  "s.total_oil_consumed, s.total_puffs, s.normalized_tpm, "
                  "s.burn_status, s.clog_status, s.leak_status "
                  "FROM samples s JOIN tests t ON s.test_id = t.id "
                  "WHERE t.file_id = ? ORDER BY s.test_id, s.sort_order");
        q.addBindValue(id);
        if (q.exec()) {
            while (q.next()) {
                SampleResult sr;
                sr.id                  = q.value(0).toLongLong();
                const qint64 testId    = q.value(1).toLongLong();
                sr.version             = q.value(2).toInt();
                sr.sampleName          = q.value(3).toString();
                sr.sampleID            = q.value(4).toString();
                sr.date                = q.value(5).toString();
                sr.tester              = q.value(6).toString();
                sr.media               = q.value(7).toString();
                sr.viscosity           = q.value(8).toDouble();
                sr.resistance          = q.value(9).toDouble();
                sr.voltage             = q.value(10).toDouble();
                sr.power               = q.value(11).toDouble();
                sr.heatingTechnology   = q.value(12).toString();
                sr.puffingRegime       = q.value(13).toString();
                sr.initialOilMass      = q.value(14).toDouble();
                sr.averageTPM          = q.value(15).toDouble();
                sr.stdDevTPM           = q.value(16).toDouble();
                sr.averagePowerDensity = q.value(17).toDouble();
                sr.efficiencyPercent   = q.value(18).toDouble();
                sr.totalOilConsumed    = q.value(19).toDouble();
                sr.totalPuffs          = q.value(20).toInt();
                sr.normalizedTPM       = q.value(21).toDouble();
                sr.burnStatus          = q.value(22).toString();
                sr.clogStatus          = q.value(23).toString();
                sr.leakStatus          = q.value(24).toString();
                testForSample.insert(sr.id, testId);
                samplesByTest[testId].append(sr);
            }
        } else {
            m_lastError = QStringLiteral("loadFile(bulk SELECT samples): ")
                          + q.lastError().text();
        }
    }

    // Bulk SELECT 2/3 — all data_rows whose sample belongs to this file.
    QHash<qint64, QVector<DataRow>> rowsBySample;
    QSet<qint64> samplesWithRegime;
    {
        QSqlQuery q(db);
        q.prepare("SELECT dr.id, dr.sample_id, dr.version, dr.puffs, "
                  "dr.before_weight, dr.after_weight, dr.draw_pressure, "
                  "dr.resistance, dr.smell, dr.clog, dr.notes, dr.tpm, "
                  "dr.tpm_power_density, dr.variation_tpm, dr.oil_consumed, "
                  "dr.puffing_regime "
                  "FROM data_rows dr "
                  "JOIN samples s ON dr.sample_id = s.id "
                  "JOIN tests   t ON s.test_id    = t.id "
                  "WHERE t.file_id = ? ORDER BY dr.sample_id, dr.sort_order");
        q.addBindValue(id);
        if (q.exec()) {
            while (q.next()) {
                DataRow dr;
                dr.id              = q.value(0).toLongLong();
                const qint64 sId   = q.value(1).toLongLong();
                dr.version         = q.value(2).toInt();
                dr.puffs           = q.value(3).toDouble();
                dr.beforeWeight    = q.value(4).toDouble();
                dr.afterWeight     = q.value(5).toDouble();
                dr.drawPressure    = q.value(6).toDouble();
                dr.resistance      = q.value(7).toDouble();
                dr.smell           = q.value(8).toString();
                dr.clog            = q.value(9).toString();
                dr.notes           = q.value(10).toString();
                dr.tpm             = q.value(11).toDouble();
                dr.tpmPowerDensity = q.value(12).toDouble();
                dr.variationTPM    = q.value(13).toDouble();
                dr.oilConsumed     = q.value(14).toDouble();
                const QVariant pr  = q.value(15);
                if (!pr.isNull()) { dr.puffingRegime = pr.toString(); samplesWithRegime.insert(sId); }
                rowsBySample[sId].append(dr);
            }
        } else {
            m_lastError = QStringLiteral("loadFile(bulk SELECT data_rows): ")
                          + q.lastError().text();
        }
    }

    // Bulk SELECT 3/3 — all images whose sample belongs to this file.
    // Image BLOBs are materialised to %LOCALAPPDATA%/.../ImageCache so
    // imagePaths can point at on-disk files.
    struct ImageInfo {
        qint64 sampleId; qint64 id; int version;
        QString fileName; QByteArray blob;
        QRectF layout; QRectF crop;
    };
    QHash<qint64, QVector<ImageInfo>> imagesBySample;
    {
        QSqlQuery q(db);
        q.prepare("SELECT im.id, im.sample_id, im.version, im.file_name, "
                  "im.image_data, im.layout_x, im.layout_y, im.layout_w, "
                  "im.layout_h, im.crop_x, im.crop_y, im.crop_w, im.crop_h "
                  "FROM images im "
                  "JOIN samples s ON im.sample_id = s.id "
                  "JOIN tests   t ON s.test_id    = t.id "
                  "WHERE t.file_id = ? ORDER BY im.sample_id, im.sort_order");
        q.addBindValue(id);
        if (q.exec()) {
            while (q.next()) {
                ImageInfo info;
                info.id       = q.value(0).toLongLong();
                info.sampleId = q.value(1).toLongLong();
                info.version  = q.value(2).toInt();
                info.fileName = q.value(3).toString();
                info.blob     = q.value(4).toByteArray();
                info.layout   = QRectF(q.value(5).toDouble(), q.value(6).toDouble(),
                                       q.value(7).toDouble(), q.value(8).toDouble());
                info.crop     = QRectF(q.value(9).toDouble(),  q.value(10).toDouble(),
                                       q.value(11).toDouble(), q.value(12).toDouble());
                imagesBySample[info.sampleId].append(info);
            }
        } else {
            m_lastError = QStringLiteral("loadFile(bulk SELECT images): ")
                          + q.lastError().text();
        }
    }

    const QString tempDir =
        QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation)
        + "/ImageCache";
    if (!imagesBySample.isEmpty()) QDir().mkpath(tempDir);

    // Assemble in test order. Buckets are pre-keyed so the assembly is
    // pure pointer chasing in memory.
    for (const TestInfo& ti : tests) {
        SheetResult sheet;
        sheet.id               = ti.id;
        sheet.version          = ti.version;
        sheet.sheetName        = ti.sheetName;
        sheet.templateVersion  = ti.templateVersion;
        sheet.overallAvgTPM    = ti.avgTPM;
        sheet.overallStdDevTPM = ti.stddevTPM;
        sheet.isRawTable       = ti.isRaw;
        sheet.fromInferredSchema = ti.fromInferred;
        result.sheetNames.append(ti.sheetName);
        // Reconstruct raw grid from JSONB (no-op when ti.rawGrid is empty).
        if (ti.isRaw)
            rawGridFromJson(ti.rawGrid, sheet.rawHeaders, sheet.rawRows);

        QVector<SampleResult> samples = samplesByTest.value(ti.id);
        for (SampleResult& sr : samples) {
            sr.rows = rowsBySample.value(sr.id);
            if (samplesWithRegime.contains(sr.id)) sheet.hasPerRowRegime = true;
            const QVector<ImageInfo>& imgs = imagesBySample.value(sr.id);
            for (const ImageInfo& info : imgs) {
                const QString tempPath = materialiseImageBlob(
                    tempDir, info.blob, info.fileName);
                sr.imagePaths.append(tempPath);
                sr.imageLayouts.append(info.layout);
                sr.imageCrops.append(info.crop);
                sr.imageIds.append(info.id);
                sr.imageVersions.append(info.version);
            }
            sheet.samples.append(sr);
        }

        // Rebuild plot series from row data (not stored separately).
        for (const SampleResult& sr : sheet.samples) {
            for (const DataRow& dr : sr.rows) {
                if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
                sheet.tpmTrend.append(dr.tpm);
                sheet.puffCounts.append(dr.puffs);
            }
        }

        // Detect incomplete DB data: a raw sheet with no grid, or a normal
        // sheet with aggregate-TPM samples that have no per-row data (legacy /
        // partially-migrated records). Consumed by the Task 8 banner.
        {
            bool incomplete = false;
            if (sheet.isRawTable) {
                incomplete = sheet.rawHeaders.isEmpty();
            } else {
                for (const SampleResult& s : sheet.samples) {
                    if (s.averageTPM > 0.0 && s.rows.isEmpty()) {
                        incomplete = true;
                        break;
                    }
                }
            }
            sheet.dbDataIncomplete = incomplete;
        }

        result.sheets.append(sheet);
    }

    logDebug(QString("Loaded file id=%1 '%2' (%3 sheets, version=%4)")
                 .arg(id).arg(result.fileName).arg(result.sheets.size())
                 .arg(result.version));
    return result;
}

FileResult DatabaseManager::loadFileByPath(const QString& filePath) const {
    m_lastError.clear();
    FileResult result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->loadFileByPath(filePath);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadFileByPath: database not open");
        return result;
    }
    QSqlQuery q(m_pg->queryDb());
    // F6: a path may now have several versioned rows. Return the MOST RECENTLY
    // added one — this is the row an OCC-recovery re-save (persistLoadedFile)
    // should continue, and the most useful target for any path-based load.
    q.prepare("SELECT id FROM files WHERE file_path = ? "
              "ORDER BY added_at DESC, id DESC LIMIT 1");
    q.addBindValue(filePath);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadFileByPath(SELECT files): ")
                      + q.lastError().text();
        return result;
    }
    if (q.next())
        return loadFile(q.value(0).toInt());
    return result;
}

// --- listFiles --------------------------------------------------------------
QVector<FileRecord> DatabaseManager::listFiles() const {
    m_lastError.clear();
    QVector<FileRecord> records;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->listFiles();
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return records;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("listFiles: database not open");
        return records;
    }

    QSqlQuery q(m_pg->queryDb());
    // F6: pull added_at and order by it so versions of one path appear
    // newest-first; the DB browser surfaces every version distinctly.
    q.prepare("SELECT id, file_path, file_name, loaded_at, template_version, "
              "sheet_count, sample_count, added_at, app_version "
              "FROM files ORDER BY added_at DESC, id DESC");
    if (!q.exec()) {
        m_lastError = QStringLiteral("listFiles(SELECT files): ")
                      + q.lastError().text();
        return records;
    }
    while (q.next()) {
        FileRecord r;
        r.id              = q.value(0).toInt();
        r.filePath        = q.value(1).toString();
        r.fileName        = q.value(2).toString();
        r.loadedAt        = q.value(3).toString();
        r.templateVersion = q.value(4).toString();
        r.sheetCount      = q.value(5).toInt();
        r.sampleCount     = q.value(6).toInt();
        r.addedAt         = q.value(7).toDateTime().toString(Qt::ISODate);
        r.appVersion      = q.value(8).toString();
        records.append(r);
    }
    return records;
}

// --- removeFile -------------------------------------------------------------
bool DatabaseManager::removeFile(int id) {
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("removeFile: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM files WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("removeFile(DELETE files): ")
                      + q.lastError().text();
        return false;
    }
    return true;
}

// --- normalizeLegacyScores --------------------------------------------------
// SP4 A5: "Repair Legacy Scores" — invoke the server-side normalizer NOW. This
// is the same idempotent, lossless string->numeric coercion the nightly pg_cron
// job runs; it rewrites every damaged row across BOTH session tables (not a
// selection). Returns the count of rows rewritten, or -1 on error.
int DatabaseManager::normalizeLegacyScores() {
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("normalizeLegacyScores: database is offline (read-only mode)");
        return -1;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("normalizeLegacyScores: database not open");
        return -1;
    }
    QSqlQuery q(m_pg->queryDb());
    if (!q.exec(QStringLiteral("SELECT dve_normalize_legacy_json()"))) {
        m_lastError = QStringLiteral("normalizeLegacyScores: ")
                      + q.lastError().text();
        return -1;
    }
    return q.next() ? q.value(0).toInt() : 0;
}

// --- deduplicateFiles -------------------------------------------------------
// 1. Delete every row whose template_version is the literal string "unknown"
//    (these come from earlier broken loads).
// 2. For each distinct file_name, keep the N most recently loaded rows and
//    delete the rest. CASCADE handles all the children.
int DatabaseManager::deduplicateFiles(int keepPerName) {
    m_lastError.clear();
    if (!m_online) {
        // Returns 0 (consistent with existing "early-out" semantics — the
        // method also returns 0 when nothing is removed).
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return 0;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("deduplicateFiles: database not open");
        return 0;
    }

    QSqlDatabase& db = m_pg->queryDb();
    int deleted = 0;

    // 1. unknown-template wipeout
    {
        QSqlQuery q(db);
        q.prepare("SELECT id FROM files WHERE template_version = 'unknown'");
        if (q.exec()) {
            QVector<int> ids;
            while (q.next()) ids.append(q.value(0).toInt());
            for (int id : ids) {
                if (removeFile(id)) ++deleted;
            }
        } else {
            m_lastError = QStringLiteral("deduplicateFiles(unknown-scan): ")
                          + q.lastError().text();
        }
    }

    // 2. per-file_name retention
    QStringList names;
    {
        QSqlQuery q(db);
        if (q.exec("SELECT DISTINCT file_name FROM files")) {
            while (q.next()) names.append(q.value(0).toString());
        } else {
            m_lastError = QStringLiteral("deduplicateFiles(name-scan): ")
                          + q.lastError().text();
        }
    }
    for (const QString& name : names) {
        QSqlQuery q(db);
        q.prepare("SELECT id FROM files WHERE file_name = ? ORDER BY loaded_at DESC");
        q.addBindValue(name);
        if (!q.exec()) continue;

        QVector<int> ids;
        while (q.next()) ids.append(q.value(0).toInt());

        for (int i = keepPerName; i < ids.size(); ++i) {
            if (removeFile(ids[i])) ++deleted;
        }
    }

    logDebug(QString("Deduplicated files: removed %1 entries").arg(deleted));
    return deleted;
}

// --- recentFilePaths --------------------------------------------------------
QStringList DatabaseManager::recentFilePaths() const {
    m_lastError.clear();
    QStringList paths;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            // OfflineSnapshot has no dedicated recentFilePaths; derive from
            // listFiles() (already sorted by loaded_at DESC) and cap at 20 to
            // match the online SELECT.
            const auto recs = m_snapshot->listFiles();
            for (int i = 0; i < recs.size() && paths.size() < 20; ++i)
                paths << recs[i].filePath;
            return paths;
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return paths;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("recentFilePaths: database not open");
        return paths;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT file_path FROM files ORDER BY loaded_at DESC LIMIT 20");
    if (!q.exec()) {
        m_lastError = QStringLiteral("recentFilePaths(SELECT files): ")
                      + q.lastError().text();
        return paths;
    }
    while (q.next()) paths << q.value(0).toString();
    return paths;
}

// ============================================================================
//  Sensory sessions
// ============================================================================
//
// JSON-serialization contract: all SensorySession fields are packed into a
// single JSONB blob (json_data). A subset (session_name, tester_name, date,
// assessor_name, media, puff_length, timestamp) also goes into dedicated
// columns to support the natural-key UNIQUE index and SELECT-without-parse on
// the listing path. layout_json lives in its own column so the report-preview
// preserves it independently of saveSensorySession (saveSensoryLayout below
// UPDATEs only layout_json).
//
// Optimistic concurrency: when s.id != -1 && s.version > 0 the row is
// UPDATEd with WHERE id = ? AND version = ?; rowcount == 0 triggers a follow-
// up SELECT classified into VersionMismatch / RowDeleted. Fresh sessions
// (s.id == -1) INSERT and map SQLSTATE 23505 (duplicate natural-key) to
// UniqueViolation.

namespace {

// Thin wrappers around the canonical pipeline-layer JSON helpers
// (see src/pipeline/SensoryData.cpp). The compact byte serialization +
// the QByteArray-in / bool-out signature are preserved so callers
// don't need to change. All three persistence paths (this file's
// JSONB column, OfflineSnapshot's local SQLite copy, and SensoryPanel's
// user-facing .json export/import) route through the same encoder /
// decoder, so any future field add lands in one place.
QString serializeSensoryJson(const SensorySession& s)
{
    return QString::fromUtf8(
        QJsonDocument(sensorySessionToJson(s)).toJson(QJsonDocument::Compact));
}

bool deserializeSensoryJson(const QByteArray& bytes, SensorySession& sess)
{
    const QJsonDocument doc = QJsonDocument::fromJson(bytes);
    if (doc.isNull() || !doc.isObject()) return false;
    sess = sensorySessionFromJson(doc.object());
    return true;
}

// DetailedSensorySession JSON now lives in the pipeline layer
// (src/pipeline/DetailedSensoryData.cpp), shared with the Plan C recovery
// snapshot. The thin wrappers below preserve the old compact-string /
// QByteArray-in / bool-out call shape so the call sites in this file stay
// unchanged.
QString serializeDetailedSensoryJson(const DetailedSensorySession& s)
{
    return QString::fromUtf8(
        QJsonDocument(detailedSensorySessionToJson(s)).toJson(QJsonDocument::Compact));
}

bool deserializeDetailedSensoryJson(const QByteArray& bytes, DetailedSensorySession& sess)
{
    const QJsonDocument doc = QJsonDocument::fromJson(bytes);
    if (doc.isNull() || !doc.isObject()) return false;
    sess = detailedSensorySessionFromJson(doc.object());
    return true;
}

// Read all rows from one of the *_images tables into the supplied path/layout/
// crop vectors. BYTEA blobs are materialised under AppLocalDataLocation/
// ImageCache/ so callers can treat them as on-disk files (mirrors the
// file-hierarchy loadFile pattern from 3a/3b).
//
// C3: outIds/outVersions are optional. When supplied, populated parallel to
// outPaths so upsertImagesFor can target rows by id on the next save instead
// of the legacy DELETE-cascade-rebuild that wiped concurrent users' images.
void loadImagesFor(QSqlDatabase& db,
                   const QString& tableName,
                   const QString& cachePrefix,
                   qint64 sessionId,
                   QStringList* outPaths,
                   QVector<QRectF>* outLayouts,
                   QVector<QRectF>* outCrops,
                   QVector<qint64>* outIds = nullptr,
                   QVector<int>* outVersions = nullptr)
{
    QSqlQuery qi(db);
    if (!qi.prepare(QString("SELECT id, version, file_name, image_data, "
                            "layout_x, layout_y, layout_w, layout_h, "
                            "crop_x, crop_y, crop_w, crop_h "
                            "FROM %1 WHERE session_id = ? ORDER BY sort_order").arg(tableName))) {
        return;
    }
    qi.addBindValue(static_cast<qlonglong>(sessionId));
    if (!qi.exec()) return;

    const QString tempDir =
        QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation)
        + "/ImageCache";
    QDir().mkpath(tempDir);

    while (qi.next()) {
        const qint64 imgId   = qi.value(0).toLongLong();
        const int    imgVer  = qi.value(1).toInt();
        const QString fileName = qi.value(2).toString();
        const QByteArray blob   = qi.value(3).toByteArray();
        const QRectF layout(qi.value(4).toDouble(), qi.value(5).toDouble(),
                            qi.value(6).toDouble(), qi.value(7).toDouble());
        const QRectF crop(qi.value(8).toDouble(), qi.value(9).toDouble(),
                          qi.value(10).toDouble(), qi.value(11).toDouble());

        // v2.0.2 H8: content-hash filename for cross-session dedup. The
        // cachePrefix is no longer needed for uniqueness — identical blobs
        // share one on-disk file regardless of which session loaded them.
        Q_UNUSED(cachePrefix);
        const QString tempPath = materialiseImageBlob(tempDir, blob, fileName);
        Q_UNUSED(sessionId);
        outPaths->append(tempPath);
        outLayouts->append(layout);
        outCrops->append(crop);
        if (outIds)      outIds->append(imgId);
        if (outVersions) outVersions->append(imgVer);
    }
}

// C3: id-aware upsert for *_images. Replaces the legacy DELETE+insertImagesFor
// destructive rebuild used by tryWriteSensorySession (const&/by-ref) and
// tryWriteDetailedSensorySession (const&/by-ref). Mirrors the three-phase
// algorithm from tryWriteFile (commit d9092c8): pre-image SELECT, per-row
// UPDATE-or-INSERT, post-prune. inOutIds and inOutVersions are sized to
// match imagePaths.size() on entry; on return they contain the post-save
// server-assigned id/version for every position. Aborts (returns false +
// outError) on any server error or OCC version-mismatch.
bool upsertImagesFor(QSqlDatabase& db,
                     const QString& tableName,
                     qint64 sessionId,
                     const QStringList& imagePaths,
                     const QVector<QRectF>& imageLayouts,
                     const QVector<QRectF>& imageCrops,
                     QVector<qint64>& inOutIds,
                     QVector<int>& inOutVersions,
                     const QString& who,
                     QString* outError)
{
    auto setError = [outError](const QString& msg) {
        if (outError) *outError = msg;
    };

    // Phase A: pre-image
    QSet<qint64> preIds;
    {
        QSqlQuery q(db);
        if (!q.prepare(QString("SELECT id FROM %1 WHERE session_id = ?").arg(tableName))) {
            setError(q.lastError().text());
            return false;
        }
        q.addBindValue(static_cast<qlonglong>(sessionId));
        if (!q.exec()) {
            setError(q.lastError().text());
            return false;
        }
        while (q.next()) preIds.insert(q.value(0).toLongLong());
    }

    // Normalise the parallel id/version vectors to match imagePaths length.
    const int n = imagePaths.size();
    while (inOutIds.size()      < n) inOutIds.append(-1);
    while (inOutVersions.size() < n) inOutVersions.append(0);

    // Prepare UPDATE and INSERT statements
    QSqlQuery updateImg(db), insertImg(db);
    if (!updateImg.prepare(QString(
            "UPDATE %1 SET session_id = ?, sort_order = ?, file_name = ?, "
            "image_data = ?, layout_x = ?, layout_y = ?, layout_w = ?, "
            "layout_h = ?, crop_x = ?, crop_y = ?, crop_w = ?, crop_h = ?, "
            "updated_by = ? "
            "WHERE id = ? AND version = ? RETURNING version").arg(tableName))) {
        setError(updateImg.lastError().text());
        return false;
    }
    if (!insertImg.prepare(QString(
            "INSERT INTO %1 (session_id, sort_order, file_name, image_data, "
            "layout_x, layout_y, layout_w, layout_h, crop_x, crop_y, crop_w, "
            "crop_h, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "RETURNING id, version").arg(tableName))) {
        setError(insertImg.lastError().text());
        return false;
    }

    QSet<qint64> postIds;
    for (int i = 0; i < n; ++i) {
        QByteArray imgData;
        QFile imgFile(imagePaths[i]);
        if (imgFile.open(QIODevice::ReadOnly)) {
            constexpr qint64 kMaxImageSize = 100 * 1024 * 1024;
            if (imgFile.size() <= kMaxImageSize)
                imgData = imgFile.readAll();
            else
                qWarning() << "Skipping oversized image:" << imgFile.fileName();
        }
        const QRectF layout = (i < imageLayouts.size()) ? imageLayouts[i] : QRectF();
        const QRectF crop   = (i < imageCrops.size())   ? imageCrops[i]   : QRectF(0,0,1,1);
        const QString fname = QFileInfo(imagePaths[i]).fileName();

        const qint64 id = inOutIds[i];
        const int    ver = inOutVersions[i];

        if (id != -1 && ver > 0) {
            updateImg.bindValue(0,  static_cast<qlonglong>(sessionId));
            updateImg.bindValue(1,  i);
            updateImg.bindValue(2,  fname);
            updateImg.bindValue(3,  imgData);
            updateImg.bindValue(4,  layout.x());
            updateImg.bindValue(5,  layout.y());
            updateImg.bindValue(6,  layout.width());
            updateImg.bindValue(7,  layout.height());
            updateImg.bindValue(8,  crop.x());
            updateImg.bindValue(9,  crop.y());
            updateImg.bindValue(10, crop.width());
            updateImg.bindValue(11, crop.height());
            updateImg.bindValue(12, who);
            updateImg.bindValue(13, static_cast<qlonglong>(id));
            updateImg.bindValue(14, ver);
            if (!updateImg.exec()) { setError(updateImg.lastError().text()); return false; }
            if (!updateImg.next()) {
                // OCC miss: classify and surface
                QString detail;
                const WriteResult cls = classifyMissingUpdate(
                    db, tableName, id, &detail);
                setError(QStringLiteral("UPDATE %1 id=%2: %3").arg(tableName)
                             .arg(id).arg(cls == WriteResult::VersionMismatch
                                 ? "version mismatch"
                                 : (cls == WriteResult::RowDeleted ? "row deleted" : detail)));
                return false;
            }
            inOutVersions[i] = updateImg.value(0).toInt();
            postIds.insert(id);
        } else {
            insertImg.bindValue(0,  static_cast<qlonglong>(sessionId));
            insertImg.bindValue(1,  i);
            insertImg.bindValue(2,  fname);
            insertImg.bindValue(3,  imgData);
            insertImg.bindValue(4,  layout.x());
            insertImg.bindValue(5,  layout.y());
            insertImg.bindValue(6,  layout.width());
            insertImg.bindValue(7,  layout.height());
            insertImg.bindValue(8,  crop.x());
            insertImg.bindValue(9,  crop.y());
            insertImg.bindValue(10, crop.width());
            insertImg.bindValue(11, crop.height());
            insertImg.bindValue(12, who);
            if (!insertImg.exec() || !insertImg.next()) {
                setError(insertImg.lastError().text());
                return false;
            }
            inOutIds[i]      = insertImg.value(0).toLongLong();
            inOutVersions[i] = insertImg.value(1).toInt();
            postIds.insert(inOutIds[i]);
        }
    }

    // Phase C: prune orphans
    QStringList orphanCsv;
    orphanCsv.reserve(preIds.size());
    for (qint64 id : preIds) {
        if (!postIds.contains(id)) orphanCsv.append(QString::number(id));
    }
    if (!orphanCsv.isEmpty()) {
        QSqlQuery q(db);
        // Safe to interpolate: orphanCsv elements are all qint64-from-DB ids.
        if (!q.exec(QString("DELETE FROM %1 WHERE id IN (%2)")
                        .arg(tableName, orphanCsv.join(",")))) {
            setError(q.lastError().text());
            return false;
        }
    }
    return true;
}

// -- Sensory save core --------------------------------------------------------
// Both saveSensorySession overloads (const and by-ref) share this body. The
// caller passes optional pointers for the post-write id and version, which
// the by-ref overload then propagates back into its struct.
//
// Three branches mirror tryWriteFile:
//   (a) s.id != -1 && s.version > 0 → UPDATE WHERE id=? AND version=?
//   (b) s.id == -1                    → INSERT
//   (c) UPSERT by natural key — used as a fallback when the caller doesn't
//       have id+version. NOT used here: we explicitly forbid silent UPSERTs
//       once optimistic concurrency is in play, because they mask
//       VersionMismatch into a successful overwrite of someone else's
//       changes. The const overload still has to support the "save a fresh
//       struct with possibly-conflicting natural key" case — that goes
//       through INSERT and surfaces UniqueViolation. Callers must then
//       load-then-merge.
WriteResult tryWriteSensoryCore(QSqlDatabase& db,
                                const SensorySession& s,
                                const QString& who,
                                const QString& jsonStr,
                                qint64* outId,
                                int* outVersion,
                                QString* outError)
{
    auto setError = [outError](const QString& msg) {
        if (outError) *outError = msg;
    };

    if (s.id != -1 && s.version > 0) {
        // DATAVIEWER-4: read-merge-write. Pull the current blob (same txn) and
        // keep LiveSync-owned per-cell scores so this wholesale write can't
        // reset them to the serializer's 5.0 default. Falls back to the raw
        // in-memory blob if the row is gone -- the guarded UPDATE below then
        // returns RowDeleted exactly as before.
        //
        // The same SELECT also fetches the row's CURRENT committed version,
        // which the UPDATE binds instead of the routinely-stale s.version
        // (LiveSync per-cell commits and this client's own prior saves both
        // bump the DB version). Binding s.version made routine whole-session
        // saves fail with VersionMismatch — which callers treated as "already
        // synced" and silently dropped.
        //
        // DESIGN (v2.5.0 decision, see
        // docs/superpowers/specs/2026-06-10-v24-save-sync-regressions-evidence.md):
        // a whole-session save deliberately adopts the current version, so two
        // clients each saving the whole session resolve as ROW-LEVEL
        // last-writer-wins BY DESIGN. Cross-client cell-level protection lives
        // elsewhere: the LiveSync per-cell stream plus the DB-score-preserving
        // merge above (dirty-aware in plan Task 3). The SELECT takes FOR
        // UPDATE, so a concurrent whole-session saver serializes behind this
        // transaction instead of interleaving with the merge read — the
        // in-transaction race is closed. The `AND version = ?` clause is now a
        // defensive invariant whose mismatch outcome is unreachable. s.version
        // remains the no-row fallback, preserving RowDeleted classification.
        QString jsonToWrite = jsonStr;
        int expectedVersion = s.version;
        {
            QSqlQuery sel(db);
            sel.prepare("SELECT json_data, version FROM sensory_sessions WHERE id = ? FOR UPDATE");
            sel.addBindValue(s.id);
            if (sel.exec() && sel.next()) {
                const QJsonObject dbRoot =
                    QJsonDocument::fromJson(sel.value(0).toString().toUtf8()).object();
                const QJsonObject memRoot =
                    QJsonDocument::fromJson(jsonStr.toUtf8()).object();
                // v2.5.0 Task 3 (RC2): pass the panel-supplied dirty-cell set so
                // scores the user edited this run stay in-memory-authoritative
                // even when LiveSync never streamed them (id<=0 at edit time or
                // a broken sync connection). Untouched scores stay DB-authoritative.
                jsonToWrite = QString::fromUtf8(QJsonDocument(
                    mergeSensoryPreservingDbScores(memRoot, dbRoot, s.dirtyCells,
                                                   s.removedSampleUids))
                        .toJson(QJsonDocument::Compact));
                expectedVersion = sel.value(1).toInt();
            }
        }

        QSqlQuery q(db);
        q.prepare(R"(
            UPDATE sensory_sessions SET
                session_name  = ?,
                tester_name   = ?,
                assessor_name = ?,
                media         = ?,
                puff_length   = ?,
                date          = ?,
                timestamp     = ?,
                json_data     = CAST(? AS JSONB),
                updated_by    = ?
            WHERE id = ? AND version = ?
            RETURNING id, version
        )");
        q.addBindValue(s.sessionName);
        q.addBindValue(s.testerName);
        q.addBindValue(s.assessorName);
        q.addBindValue(s.media);
        q.addBindValue(s.puffLength);
        q.addBindValue(s.date);
        q.addBindValue(s.timestamp);
        q.addBindValue(jsonToWrite);
        q.addBindValue(who);
        q.addBindValue(s.id);
        q.addBindValue(expectedVersion);
        if (!q.exec()) {
            const QString code = q.lastError().nativeErrorCode();
            setError(QStringLiteral("UPDATE sensory_sessions: ") + q.lastError().text());
            if (code == QString::fromLatin1(kSqlStateUniqueViolation))
                return WriteResult::UniqueViolation;
            return WriteResult::OtherError;
        }
        if (!q.next()) {
            // No row matched id+version. classifyMissingUpdate handles its
            // own error-text on internal SQL failure.
            QString detail;
            const WriteResult cls = classifyMissingUpdate(
                db, QStringLiteral("sensory_sessions"), s.id, &detail);
            if (cls == WriteResult::VersionMismatch) {
                setError(QStringLiteral("UPDATE sensory_sessions: version mismatch "
                                        "(id=%1, expected version=%2)")
                             .arg(s.id).arg(expectedVersion));
            } else if (cls == WriteResult::RowDeleted) {
                setError(QStringLiteral("UPDATE sensory_sessions: row deleted "
                                        "(id=%1)").arg(s.id));
            } else {
                setError(QStringLiteral("UPDATE sensory_sessions classify: ") + detail);
            }
            return cls;
        }
        if (outId)      *outId = q.value(0).toLongLong();
        if (outVersion) *outVersion = q.value(1).toInt();
        return WriteResult::Success;
    }

    // INSERT branch — fresh struct. layout_json is NULL on insert; the
    // separate saveSensoryLayout path UPDATEs it later.
    QSqlQuery q(db);
    q.prepare(R"(
        INSERT INTO sensory_sessions
            (session_name, tester_name, assessor_name, media, puff_length,
             date, timestamp, json_data, layout_json, updated_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, CAST(? AS JSONB), NULL, ?)
        RETURNING id, version
    )");
    q.addBindValue(s.sessionName);
    q.addBindValue(s.testerName);
    q.addBindValue(s.assessorName);
    q.addBindValue(s.media);
    q.addBindValue(s.puffLength);
    q.addBindValue(s.date);
    q.addBindValue(s.timestamp);
    q.addBindValue(jsonStr);
    q.addBindValue(who);
    if (!q.exec() || !q.next()) {
        const QString code = q.lastError().nativeErrorCode();
        setError(QStringLiteral("INSERT sensory_sessions: ") + q.lastError().text());
        if (code == QString::fromLatin1(kSqlStateUniqueViolation))
            return WriteResult::UniqueViolation;
        return WriteResult::OtherError;
    }
    if (outId)      *outId = q.value(0).toLongLong();
    if (outVersion) *outVersion = q.value(1).toInt();
    return WriteResult::Success;
}

} // namespace

WriteResult DatabaseManager::tryWriteSensorySession(const SensorySession& s)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteSensorySession: database not open");
        return WriteResult::OtherError;
    }

    const QString jsonStr = serializeSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteSensorySession(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    logDebug(QString("tryWriteSensorySession: name='%1' tester='%2' date='%3' samples=%4 id=%5 v=%6")
                 .arg(s.sessionName, s.testerName, s.date)
                 .arg(s.samples.size()).arg(s.id).arg(s.version));

    qint64 sessionId = -1;
    int    newVer   = 0;
    QString coreErr;
    const WriteResult coreResult = tryWriteSensoryCore(db, s, who, jsonStr,
                                                       &sessionId, &newVer, &coreErr);
    if (coreResult != WriteResult::Success) {
        m_lastError = QStringLiteral("tryWriteSensorySession(") + coreErr + QStringLiteral(")");
        db.rollback();
        logDebug(m_lastError);
        return coreResult;
    }

    // C3: id-aware upsert (replaces the legacy DELETE-then-rebuild that wiped
    // concurrent users' images on every save). The const overload takes local
    // mutable copies of the id/version vectors — callers wanting the post-save
    // identities back should use the by-ref overload below.
    QVector<qint64> localIds = s.imageIds;
    QVector<int>    localVers = s.imageVersions;
    QString imgErr;
    if (!upsertImagesFor(db, "sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops,
                         localIds, localVers, who, &imgErr)) {
        m_lastError = QStringLiteral("tryWriteSensorySession(upsert sensory_images): ") + imgErr;
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteSensorySession(commit): ")
                      + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }
    return WriteResult::Success;
}

WriteResult DatabaseManager::tryWriteSensorySession(SensorySession& s)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteSensorySession: database not open");
        return WriteResult::OtherError;
    }

    const QString jsonStr = serializeSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    // RC1 wrapper resilience (v2.5.0 decision, see
    // docs/superpowers/specs/2026-06-10-v24-save-sync-regressions-evidence.md):
    // the core's fresh-version OCC (commits 0c21100/377f827) makes a genuine
    // VersionMismatch near-unreachable, and RowDeleted means another client
    // removed this row out-of-band — neither must ever silently drop the user's
    // in-memory edits. So:
    //   * RowDeleted   -> reset id/version on a local copy and re-run the core
    //                     through its INSERT branch, re-creating the user's data
    //                     as a fresh row (the SELECT-found-no-row left the txn
    //                     intact, so no rollback is needed first).
    //   * VersionMismatch (now near-unreachable) -> retry the core ONCE; its
    //                     fresh-version read happens again on the retry.
    // The struct is passed by const-ref to the core, so the re-INSERT runs on a
    // mutable local copy; the byRef back-fill below still propagates whatever
    // id/version the surviving write produced.
    SensorySession local = s;
    qint64 sessionId = -1;
    int    newVer    = 0;
    QString coreErr;
    WriteResult coreResult = tryWriteSensoryCore(db, local, who, jsonStr,
                                                 &sessionId, &newVer, &coreErr);
    if (coreResult == WriteResult::RowDeleted) {
        logDebug(QStringLiteral("tryWriteSensorySession(byRef): row deleted out-of-band "
                                "(id=%1) — re-INSERTing in-memory data as a fresh row")
                     .arg(local.id));
        // The session row was deleted out-of-band; its sensory_images rows were
        // CASCADE-removed with it. Reset the session anchor AND every image
        // anchor so the re-INSERT recreates the children under the new session
        // id — mirroring resetFileIdsForReinsert for the TPM path. The image
        // anchors are reset on `s` (not just `local`) because upsertImagesFor
        // below is called with s.imageIds/s.imageVersions: if those still held
        // the deleted ids it would take its UPDATE branch (WHERE id=<old>),
        // match nothing, and the whole save would fail forever (dirty flag
        // retries indefinitely).
        local.id = -1; local.version = 0;
        for (qint64& imgId : s.imageIds)   imgId = -1;
        for (int& imgVer : s.imageVersions) imgVer = 0;
        coreResult = tryWriteSensoryCore(db, local, who, jsonStr,
                                         &sessionId, &newVer, &coreErr);
    } else if (coreResult == WriteResult::VersionMismatch) {
        logDebug(QStringLiteral("tryWriteSensorySession(byRef): version mismatch "
                                "(id=%1) — retrying core once").arg(local.id));
        coreResult = tryWriteSensoryCore(db, local, who, jsonStr,
                                         &sessionId, &newVer, &coreErr);
    }
    if (coreResult != WriteResult::Success) {
        m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(") + coreErr + QStringLiteral(")");
        db.rollback();
        logDebug(m_lastError);
        return coreResult;
    }

    // C3: id-aware upsert. By-ref overload back-fills s.imageIds/imageVersions
    // with the server-assigned identities so the next save can find the same
    // rows and UPDATE in place (no DELETE-rebuild). On a RowDeleted re-INSERT
    // recovery the RowDeleted branch above already zeroed s.imageIds/imageVersions,
    // so every image takes the INSERT branch and hangs off the new sessionId.
    QString imgErr;
    if (!upsertImagesFor(db, "sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops,
                         s.imageIds, s.imageVersions, who, &imgErr)) {
        m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(upsert sensory_images): ")
                      + imgErr;
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteSensorySession(byRef)(commit): ")
                      + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    s.id      = static_cast<int>(sessionId);
    s.version = newVer;
    return WriteResult::Success;
}

bool DatabaseManager::saveSensorySession(const SensorySession& s) {
    return tryWriteSensorySession(s) == WriteResult::Success;
}

bool DatabaseManager::saveSensorySession(SensorySession& s) {
    return tryWriteSensorySession(s) == WriteResult::Success;
}

WriteResult DatabaseManager::tryWriteSensorySessionAutoSuffix(SensorySession& s,
                                                              int maxAttempts)
{
    // v2.5.0 RC4 — duplicate/renamed sessions self-resolve with _1/_2/_3 instead
    // of the old modal "name already in use" block that fed the June-10 endless
    // re-INSERT loop (rename detected -> id=-1 -> INSERT -> 23505 -> dialog ->
    // skip -> stale baseline -> repeat). On each UniqueViolation we bump BOTH
    // sessionName and testTitle (buildSession regenerates sessionName FROM
    // testTitle, so suffixing only sessionName would regress on the next save)
    // and retry. The byRef back-fill writes id/version on Success; we also stamp
    // originalSessionName so the panel's rename detector treats the resolved name
    // as the new baseline and never re-collides.
    WriteResult r = tryWriteSensorySession(s);
    int attempts = 0;
    while (r == WriteResult::UniqueViolation && attempts < maxAttempts) {
        s.sessionName = OutputPaths::nextSuffixedName(s.sessionName);
        s.testTitle   = OutputPaths::nextSuffixedName(s.testTitle);
        ++attempts;
        logDebug(QStringLiteral("tryWriteSensorySessionAutoSuffix: name taken — "
                                "retrying as \"%1\" (attempt %2)")
                     .arg(s.sessionName).arg(attempts));
        r = tryWriteSensorySession(s);
    }
    if (r == WriteResult::Success)
        s.originalSessionName = s.sessionName;   // kill the rename loop at the source
    return r;
}

QVector<SensorySession> DatabaseManager::loadSensorySessions() const
{
    m_lastError.clear();
    QVector<SensorySession> result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            // OfflineSnapshot exposes listSensoryRecords + per-id loadSensorySession
            // but no bulk loader. Derive the full list one at a time. OK for
            // typical session counts (< thousands).
            const auto recs = m_snapshot->listSensoryRecords();
            result.reserve(recs.size());
            for (const auto& r : recs) {
                SensorySession s = m_snapshot->loadSensorySession(r.id);
                if (s.id > 0) result.append(s);
            }
            return result;
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadSensorySessions: database not open");
        return result;
    }

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT id, version, json_data FROM sensory_sessions ORDER BY id DESC");
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadSensorySessions(SELECT): ")
                      + q.lastError().text();
        return result;
    }

    // Step 1: read every row's (id, version, json) into memory before we issue
    // the per-session image queries — re-entering the cursor on the same
    // QSqlQuery while another QSqlQuery is in flight can confuse the QPSQL
    // driver.
    struct Row { qint64 id; int version; QByteArray json; };
    QVector<Row> rows;
    while (q.next()) {
        Row r;
        r.id      = q.value(0).toLongLong();
        r.version = q.value(1).toInt();
        r.json    = q.value(2).toString().toUtf8();
        rows.append(r);
    }

    for (const Row& r : rows) {
        SensorySession sess;
        if (!deserializeSensoryJson(r.json, sess)) continue;
        sess.id      = static_cast<int>(r.id);
        sess.version = r.version;

        loadImagesFor(db, "sensory_images", "dve_sensimg", r.id,
                      &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops,
                      &sess.imageIds, &sess.imageVersions);
        result.append(sess);
    }
    return result;
}

SensorySession DatabaseManager::loadSensorySession(int id) const
{
    m_lastError.clear();
    SensorySession sess;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->loadSensorySession(id);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return sess;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadSensorySession: database not open");
        return sess;
    }

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT version, json_data FROM sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadSensorySession(SELECT): ")
                      + q.lastError().text();
        return sess;
    }
    if (!q.next()) return sess;  // not found

    const int rowVersion = q.value(0).toInt();
    if (!deserializeSensoryJson(q.value(1).toString().toUtf8(), sess)) return sess;
    sess.id      = id;
    sess.version = rowVersion;

    loadImagesFor(db, "sensory_images", "dve_sensimg", id,
                  &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops,
                  &sess.imageIds, &sess.imageVersions);
    return sess;
}

// SP4 A4: the score keys (shared with dve_normalize_legacy_json in init.sql)
// whose string-typed values mark a legacy pre-numeric write. A row is flagged
// "Legacy string scores" if any score key still holds a string; the heal
// converts these to numbers, so post-heal this reads zero (doubles as a
// verification signal). CASE-guarded so a malformed row (non-array samples /
// non-object element) cannot error the listing query. Used by both the sensory
// and detailed-sensory listings below.
static const char* const kLegacyScoreScanExists =
    "EXISTS (SELECT 1 FROM jsonb_array_elements("
    "    CASE WHEN jsonb_typeof(json_data->'samples')='array' "
    "         THEN json_data->'samples' ELSE '[]'::jsonb END) s, "
    "  jsonb_each(CASE WHEN jsonb_typeof(s.value)='object' "
    "         THEN s.value ELSE '{}'::jsonb END) kv "
    "  WHERE jsonb_typeof(kv.value)='string' AND kv.key = ANY(ARRAY["
    "    'Burnt Taste','Vapor Volume','Overall Flavor','Smoothness','Overall Liking',"
    "    'Burn Taste','Flavor Intensity','Throat Irritation','Nasal Irritation',"
    "    'Vapor Quality Overall','Cough','Volume Consistency','Performance Consistency',"
    "    'Vapor Temperature','Vapor vs Oil'])) AS has_legacy_string_scores";

QVector<SensoryRecord> DatabaseManager::listSensoryRecords() const
{
    m_lastError.clear();
    QVector<SensoryRecord> result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->listSensoryRecords();
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("listSensoryRecords: database not open");
        return result;
    }

    QSqlQuery q(m_pg->queryDb());
    // Pull from columns first; tap json_data only for the extras the listing
    // needs (test_title, tester_name, sample count). jsonb_array_length is
    // O(1) on a JSONB so it's cheap to compute server-side per row.
    const QString sensSql = QStringLiteral(
              "SELECT id, session_name, assessor_name, media, date, "
              "       json_data->>'test_title'   AS test_title, "
              "       json_data->>'tester_name'  AS tester_name, "
              "       COALESCE(jsonb_array_length(json_data->'samples'), 0) AS sample_count, "
              "       app_version, ")
          + QLatin1String(kLegacyScoreScanExists)
          + QStringLiteral(" FROM sensory_sessions ORDER BY id DESC");
    q.prepare(sensSql);
    if (!q.exec()) {
        m_lastError = QStringLiteral("listSensoryRecords(SELECT): ")
                      + q.lastError().text();
        return result;
    }

    while (q.next()) {
        SensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();
        rec.testTitle    = q.value(5).toString();
        rec.testerName   = q.value(6).toString();
        rec.sampleCount  = q.value(7).toInt();
        rec.appVersion           = q.value(8).toString();
        rec.hasLegacyStringScores = q.value(9).toBool();
        result.append(rec);
    }
    return result;
}

bool DatabaseManager::removeSensorySession(int id)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("removeSensorySession: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("removeSensorySession(DELETE): ")
                      + q.lastError().text();
        return false;
    }
    return true;
}

QString DatabaseManager::nextDefaultTestName() const
{
    m_lastError.clear();
    if (!m_online) {
        // Naming a brand-new test is implicitly a write-side operation; even
        // if we returned a value derived from the snapshot, the user can't
        // actually create the session offline. Return the sentinel default
        // and surface the offline state via lastError().
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return QStringLiteral("test_0001");
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("nextDefaultTestName: database not open");
        return QStringLiteral("test_0001");
    }

    // Scan existing test_NNNN titles; return one past the max. Gaps in the
    // numbering are preserved on purpose — sequential is what users expect.
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT json_data->>'test_title' FROM sensory_sessions "
              "WHERE json_data->>'test_title' ~ '^test_[0-9]+$'");
    int maxNum = 0;
    const QRegularExpression rx(QStringLiteral("^test_(\\d+)$"));
    if (q.exec()) {
        while (q.next()) {
            const QString title = q.value(0).toString();
            const auto match = rx.match(title);
            if (match.hasMatch()) {
                const int num = match.captured(1).toInt();
                if (num > maxNum) maxNum = num;
            }
        }
    } else {
        m_lastError = QStringLiteral("nextDefaultTestName(SELECT): ")
                      + q.lastError().text();
    }
    return QString("test_%1").arg(maxNum + 1, 4, 10, QLatin1Char('0'));
}

// ============================================================================
//  Detailed Sensory Sessions
// ============================================================================
//
// Mirrors the sensory path. Notably, detailed_sensory_sessions has NO
// layout_json column — the radar chart layout for detailed sensory mode is
// not persisted separately. Only the json_data blob round-trips.

namespace {

// Shared core for both tryWriteDetailedSensorySession overloads. Lives in the
// transaction the caller has already started; on Success populates outId and
// outVersion (server-assigned) so the mutable-ref overload can write them back.
//
// Branches mirror tryWriteSensoryCore:
//   (a) s.id != -1 && s.version > 0 → UPDATE WHERE id=? AND version=?
//                                     RETURNING id, version
//   (b) s.id == -1                    → INSERT RETURNING id, version
WriteResult tryWriteDetailedSensoryCore(QSqlDatabase& db,
                                        const DetailedSensorySession& s,
                                        const QString& who,
                                        const QString& jsonStr,
                                        qint64* outId,
                                        int* outVersion,
                                        QString* outError)
{
    auto setError = [outError](const QString& msg) {
        if (outError) *outError = msg;
    };

    if (s.id != -1 && s.version > 0) {
        // DATAVIEWER-4: read-merge-write. Pull the current blob (same txn) and
        // keep LiveSync-owned per-cell scores so this wholesale write can't
        // reset them to the serializer's 0.0 default. Falls back to the raw
        // in-memory blob if the row is gone -- the guarded UPDATE below then
        // returns RowDeleted exactly as before.
        //
        // The same SELECT also fetches the row's CURRENT committed version,
        // bound below instead of the routinely-stale s.version. By design a
        // whole-session save adopts that version → ROW-LEVEL last-writer-wins
        // across clients (v2.5.0 decision); cross-client cell protection is the
        // LiveSync per-cell stream + the DB-score-preserving merge above
        // (dirty-aware in plan Task 3). With FOR UPDATE the in-transaction
        // SELECT→UPDATE race is closed, so the `AND version = ?` clause is a
        // defensive invariant whose mismatch outcome is unreachable; s.version
        // is the no-row fallback so RowDeleted classification is unchanged. See
        // tryWriteSensoryCore for the full rationale.
        QString jsonToWrite = jsonStr;
        int expectedVersion = s.version;
        {
            QSqlQuery sel(db);
            sel.prepare("SELECT json_data, version FROM detailed_sensory_sessions WHERE id = ? FOR UPDATE");
            sel.addBindValue(s.id);
            if (sel.exec() && sel.next()) {
                const QJsonObject dbRoot =
                    QJsonDocument::fromJson(sel.value(0).toString().toUtf8()).object();
                const QJsonObject memRoot =
                    QJsonDocument::fromJson(jsonStr.toUtf8()).object();
                // v2.5.0 Task 3 (RC2): dirty-cell set keeps this run's local
                // score edits authoritative; untouched scores stay DB-authoritative.
                jsonToWrite = QString::fromUtf8(QJsonDocument(
                    mergeDetailedSensoryPreservingDbScores(memRoot, dbRoot, s.dirtyCells))
                        .toJson(QJsonDocument::Compact));
                expectedVersion = sel.value(1).toInt();
            }
        }

        QSqlQuery q(db);
        q.prepare(R"(
            UPDATE detailed_sensory_sessions SET
                session_name  = ?,
                tester_name   = ?,
                assessor_name = ?,
                media         = ?,
                date          = ?,
                timestamp     = ?,
                json_data     = CAST(? AS JSONB),
                updated_by    = ?
            WHERE id = ? AND version = ?
            RETURNING id, version
        )");
        q.addBindValue(s.sessionName);
        q.addBindValue(s.testerName);
        q.addBindValue(s.assessorName);
        q.addBindValue(s.media);
        q.addBindValue(s.date);
        q.addBindValue(s.timestamp);
        q.addBindValue(jsonToWrite);
        q.addBindValue(who);
        q.addBindValue(s.id);
        q.addBindValue(expectedVersion);
        if (!q.exec()) {
            const QString code = q.lastError().nativeErrorCode();
            setError(QStringLiteral("UPDATE detailed_sensory_sessions: ") + q.lastError().text());
            if (code == QString::fromLatin1(kSqlStateUniqueViolation))
                return WriteResult::UniqueViolation;
            return WriteResult::OtherError;
        }
        if (!q.next()) {
            QString detail;
            const WriteResult cls = classifyMissingUpdate(
                db, QStringLiteral("detailed_sensory_sessions"), s.id, &detail);
            if (cls == WriteResult::VersionMismatch) {
                setError(QStringLiteral("UPDATE detailed_sensory_sessions: version mismatch "
                                        "(id=%1, expected version=%2)")
                             .arg(s.id).arg(expectedVersion));
            } else if (cls == WriteResult::RowDeleted) {
                setError(QStringLiteral("UPDATE detailed_sensory_sessions: row deleted "
                                        "(id=%1)").arg(s.id));
            } else {
                setError(QStringLiteral("UPDATE detailed_sensory_sessions classify: ") + detail);
            }
            return cls;
        }
        if (outId)      *outId      = q.value(0).toLongLong();
        if (outVersion) *outVersion = q.value(1).toInt();
        return WriteResult::Success;
    }

    // INSERT branch — fresh struct.
    QSqlQuery q(db);
    q.prepare(R"(
        INSERT INTO detailed_sensory_sessions
            (session_name, tester_name, assessor_name, media, date, timestamp,
             json_data, updated_by)
        VALUES (?, ?, ?, ?, ?, ?, CAST(? AS JSONB), ?)
        RETURNING id, version
    )");
    q.addBindValue(s.sessionName);
    q.addBindValue(s.testerName);
    q.addBindValue(s.assessorName);
    q.addBindValue(s.media);
    q.addBindValue(s.date);
    q.addBindValue(s.timestamp);
    q.addBindValue(jsonStr);
    q.addBindValue(who);
    if (!q.exec() || !q.next()) {
        const QString code = q.lastError().nativeErrorCode();
        setError(QStringLiteral("INSERT detailed_sensory_sessions: ") + q.lastError().text());
        if (code == QString::fromLatin1(kSqlStateUniqueViolation))
            return WriteResult::UniqueViolation;
        return WriteResult::OtherError;
    }
    if (outId)      *outId      = q.value(0).toLongLong();
    if (outVersion) *outVersion = q.value(1).toInt();
    return WriteResult::Success;
}

} // namespace

WriteResult DatabaseManager::tryWriteDetailedSensorySession(const DetailedSensorySession& s)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession: database not open");
        return WriteResult::OtherError;
    }

    const QString jsonStr = serializeDetailedSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    qint64 sessionId = -1;
    int    newVer   = 0;
    QString coreErr;
    const WriteResult coreResult = tryWriteDetailedSensoryCore(
        db, s, who, jsonStr, &sessionId, &newVer, &coreErr);
    if (coreResult != WriteResult::Success) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(") + coreErr + QStringLiteral(")");
        db.rollback();
        logDebug(m_lastError);
        return coreResult;
    }

    // C3: id-aware upsert (was DELETE+rebuild). const overload uses local
    // mutable copies of the id/version vectors.
    QVector<qint64> localIds = s.imageIds;
    QVector<int>    localVers = s.imageVersions;
    QString imgErr;
    if (!upsertImagesFor(db, "detailed_sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops,
                         localIds, localVers, who, &imgErr)) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(upsert images): ") + imgErr;
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(commit): ")
                      + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }
    return WriteResult::Success;
}

// Mutable-ref overload — see header. Mirrors the const-ref path but writes the
// post-save id+version back into `s`.
WriteResult DatabaseManager::tryWriteDetailedSensorySession(DetailedSensorySession& s)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession: database not open");
        return WriteResult::OtherError;
    }

    const QString jsonStr = serializeDetailedSensoryJson(s);
    QSqlDatabase& db = m_pg->queryDb();
    const QString who = writerUuid(m_identity);

    if (!db.transaction()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(begin transaction): ")
                      + db.lastError().text();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    // RC1 wrapper resilience — twin of tryWriteSensorySession(byRef). RowDeleted
    // re-INSERTs the in-memory data as a fresh row; VersionMismatch (now
    // near-unreachable) retries the core once. See that function for the full
    // rationale and the v2.5.0 last-writer-wins design decision.
    DetailedSensorySession local = s;
    qint64 sessionId = -1;
    int    newVer    = 0;
    QString coreErr;
    WriteResult coreResult = tryWriteDetailedSensoryCore(
        db, local, who, jsonStr, &sessionId, &newVer, &coreErr);
    if (coreResult == WriteResult::RowDeleted) {
        logDebug(QStringLiteral("tryWriteDetailedSensorySession(byRef): row deleted "
                                "out-of-band (id=%1) — re-INSERTing in-memory data "
                                "as a fresh row").arg(local.id));
        // Session row deleted out-of-band; its detailed_sensory_images rows were
        // CASCADE-removed. Reset the session anchor AND every image anchor (on
        // `s`, since upsertImagesFor below reads s.imageIds/imageVersions) so the
        // re-INSERT recreates the children under the new session id. Without the
        // image reset, upsertImagesFor would take its UPDATE branch (WHERE
        // id=<old>), match nothing, and the save would fail forever.
        local.id = -1; local.version = 0;
        for (qint64& imgId : s.imageIds)   imgId = -1;
        for (int& imgVer : s.imageVersions) imgVer = 0;
        coreResult = tryWriteDetailedSensoryCore(
            db, local, who, jsonStr, &sessionId, &newVer, &coreErr);
    } else if (coreResult == WriteResult::VersionMismatch) {
        logDebug(QStringLiteral("tryWriteDetailedSensorySession(byRef): version mismatch "
                                "(id=%1) — retrying core once").arg(local.id));
        coreResult = tryWriteDetailedSensoryCore(
            db, local, who, jsonStr, &sessionId, &newVer, &coreErr);
    }
    if (coreResult != WriteResult::Success) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(") + coreErr + QStringLiteral(")");
        db.rollback();
        logDebug(m_lastError);
        return coreResult;
    }

    // C3: id-aware upsert. By-ref overload back-fills s.imageIds/imageVersions.
    // On a RowDeleted re-INSERT recovery the branch above already zeroed
    // s.imageIds/imageVersions, so every image takes the INSERT branch and hangs
    // off the new sessionId.
    QString imgErr;
    if (!upsertImagesFor(db, "detailed_sensory_images", sessionId,
                         s.imagePaths, s.imageLayouts, s.imageCrops,
                         s.imageIds, s.imageVersions, who, &imgErr)) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(upsert images): ") + imgErr;
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    if (!db.commit()) {
        m_lastError = QStringLiteral("tryWriteDetailedSensorySession(byRef)(commit): ")
                      + db.lastError().text();
        db.rollback();
        logDebug(m_lastError);
        return WriteResult::OtherError;
    }

    s.id      = static_cast<int>(sessionId);
    s.version = newVer;
    return WriteResult::Success;
}

bool DatabaseManager::saveDetailedSensorySession(const DetailedSensorySession& s) {
    return tryWriteDetailedSensorySession(s) == WriteResult::Success;
}

WriteResult DatabaseManager::tryWriteDetailedSensorySessionAutoSuffix(
    DetailedSensorySession& s, int maxAttempts)
{
    // v2.5.0 RC4 — detailed twin of tryWriteSensorySessionAutoSuffix. Suffixes
    // sessionName + testTitle in lockstep on UniqueViolation and retries. There
    // is no originalSessionName on DetailedSensorySession (no in-place rename
    // branch in onUpdateDatabase), so there is nothing to re-baseline here.
    WriteResult r = tryWriteDetailedSensorySession(s);
    int attempts = 0;
    while (r == WriteResult::UniqueViolation && attempts < maxAttempts) {
        s.sessionName = OutputPaths::nextSuffixedName(s.sessionName);
        s.testTitle   = OutputPaths::nextSuffixedName(s.testTitle);
        ++attempts;
        logDebug(QStringLiteral("tryWriteDetailedSensorySessionAutoSuffix: name taken — "
                                "retrying as \"%1\" (attempt %2)")
                     .arg(s.sessionName).arg(attempts));
        r = tryWriteDetailedSensorySession(s);
    }
    return r;
}

QVector<DetailedSensorySession> DatabaseManager::loadDetailedSensorySessions() const
{
    m_lastError.clear();
    QVector<DetailedSensorySession> result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            const auto recs = m_snapshot->listDetailedSensoryRecords();
            result.reserve(recs.size());
            for (const auto& r : recs) {
                DetailedSensorySession s = m_snapshot->loadDetailedSensorySession(r.id);
                if (s.id > 0) result.append(s);
            }
            return result;
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadDetailedSensorySessions: database not open");
        return result;
    }

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT id, version, json_data FROM detailed_sensory_sessions ORDER BY id DESC");
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadDetailedSensorySessions(SELECT): ")
                      + q.lastError().text();
        return result;
    }

    struct Row { qint64 id; int version; QByteArray json; };
    QVector<Row> rows;
    while (q.next()) {
        Row r;
        r.id      = q.value(0).toLongLong();
        r.version = q.value(1).toInt();
        r.json    = q.value(2).toString().toUtf8();
        rows.append(r);
    }

    for (const Row& r : rows) {
        DetailedSensorySession sess;
        if (!deserializeDetailedSensoryJson(r.json, sess)) continue;
        sess.id      = static_cast<int>(r.id);
        sess.version = r.version;
        loadImagesFor(db, "detailed_sensory_images", "dve_detsensimg", r.id,
                      &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops,
                      &sess.imageIds, &sess.imageVersions);
        result.append(sess);
    }
    return result;
}

DetailedSensorySession DatabaseManager::loadDetailedSensorySession(int id) const
{
    m_lastError.clear();
    DetailedSensorySession sess;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->loadDetailedSensorySession(id);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return sess;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadDetailedSensorySession: database not open");
        return sess;
    }

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT version, json_data FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadDetailedSensorySession(SELECT): ")
                      + q.lastError().text();
        return sess;
    }
    if (!q.next()) return sess;

    const int rowVersion = q.value(0).toInt();
    if (!deserializeDetailedSensoryJson(q.value(1).toString().toUtf8(), sess)) return sess;
    sess.id      = id;
    sess.version = rowVersion;

    loadImagesFor(db, "detailed_sensory_images", "dve_detsensimg", id,
                  &sess.imagePaths, &sess.imageLayouts, &sess.imageCrops,
                  &sess.imageIds, &sess.imageVersions);
    return sess;
}

QVector<DetailedSensoryRecord> DatabaseManager::listDetailedSensoryRecords() const
{
    m_lastError.clear();
    QVector<DetailedSensoryRecord> result;
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->listDetailedSensoryRecords();
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return result;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("listDetailedSensoryRecords: database not open");
        return result;
    }

    QSqlQuery q(m_pg->queryDb());
    const QString detSql = QStringLiteral(
              "SELECT id, session_name, assessor_name, media, date, "
              "       json_data->>'test_title'  AS test_title, "
              "       json_data->>'tester_name' AS tester_name, "
              "       COALESCE(jsonb_array_length(json_data->'samples'), 0) AS sample_count, "
              "       app_version, ")
          + QLatin1String(kLegacyScoreScanExists)
          + QStringLiteral(" FROM detailed_sensory_sessions ORDER BY id DESC");
    q.prepare(detSql);
    if (!q.exec()) {
        m_lastError = QStringLiteral("listDetailedSensoryRecords(SELECT): ")
                      + q.lastError().text();
        return result;
    }

    while (q.next()) {
        DetailedSensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();
        rec.testTitle    = q.value(5).toString();
        rec.testerName   = q.value(6).toString();
        rec.sampleCount  = q.value(7).toInt();
        rec.appVersion           = q.value(8).toString();
        rec.hasLegacyStringScores = q.value(9).toBool();
        result.append(rec);
    }
    return result;
}

bool DatabaseManager::removeDetailedSensorySession(int id)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("removeDetailedSensorySession: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("DELETE FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = QStringLiteral("removeDetailedSensorySession(DELETE): ")
                      + q.lastError().text();
        return false;
    }
    return true;
}

// ============================================================================
//  Layout JSON persistence (sensory report preview)
// ============================================================================

QString DatabaseManager::loadSensoryLayout(int sessionId) const
{
    m_lastError.clear();
    if (!m_online) {
        // SensorySession has no layoutJson member, and OfflineSnapshot does
        // not currently expose layout_json through its loadSensorySession
        // accessor. The sensory report preview that consumes this layout is
        // a write-targeted UI (Save Layout button writes back), so offline
        // we surface an empty layout — the report builder degrades to the
        // default placement. Plan C C4/C5 may add a dedicated snapshot
        // accessor if the offline preview turns out to need real layouts.
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        Q_UNUSED(sessionId);
        return {};
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("loadSensoryLayout: database not open");
        return {};
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT layout_json FROM sensory_sessions WHERE id = ?");
    q.addBindValue(sessionId);
    if (!q.exec()) {
        m_lastError = QStringLiteral("loadSensoryLayout(SELECT): ")
                      + q.lastError().text();
        return {};
    }
    if (!q.next()) return {};
    return q.value(0).toString();
}

bool DatabaseManager::saveSensoryLayout(int sessionId, const QString& layoutJson)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("saveSensoryLayout: database not open");
        return false;
    }
    // v2.0.2 H4: route through dve_commit_session_layout. The stored
    // function sets dve.live_column/dve.live_value session vars so the
    // notify_row_change trigger emits a column-enriched NOTIFY, and
    // it enforces OCC on `version`. For layout saves we read the
    // current version inline and retry once on a stale-version miss —
    // layout saves are coarse (a whole chart positioning) and the
    // semantic intent is last-writer-wins, but routing through the
    // stored function fixes both the missing trigger-payload context
    // and the version-staleness-after-save bug the open-coded UPDATE had.
    QSqlDatabase& db = m_pg->queryDb();
    const QVariant layoutVal =
        layoutJson.isEmpty() ? QVariant() : QVariant(layoutJson);
    const QString writer = writerUuid(m_identity);

    for (int attempt = 0; attempt < 2; ++attempt) {
        // Read the current version. If the row doesn't exist, fail fast.
        int currentVersion = 0;
        {
            QSqlQuery vq(db);
            vq.prepare("SELECT version FROM sensory_sessions WHERE id = ?");
            vq.addBindValue(sessionId);
            if (!vq.exec()) {
                m_lastError = QStringLiteral("saveSensoryLayout(SELECT version): ")
                              + vq.lastError().text();
                return false;
            }
            if (!vq.next()) {
                m_lastError = QStringLiteral(
                    "saveSensoryLayout: row not found (id=%1)").arg(sessionId);
                return false;
            }
            currentVersion = vq.value(0).toInt();
        }

        QSqlQuery q(db);
        q.prepare("SELECT dve_commit_session_layout(?, ?, ?::jsonb, ?, ?)");
        q.addBindValue(QStringLiteral("sensory_sessions"));
        q.addBindValue(sessionId);
        q.addBindValue(layoutVal);
        q.addBindValue(writer);
        q.addBindValue(currentVersion);
        if (!q.exec()) {
            m_lastError = QStringLiteral("saveSensoryLayout(stored fn): ")
                          + q.lastError().text();
            return false;
        }
        if (!q.next()) {
            m_lastError = QStringLiteral(
                "saveSensoryLayout: no row returned from stored function");
            return false;
        }
        const QVariant ret = q.value(0);
        if (!ret.isNull()) {
            // success — function returned the new version
            return true;
        }
        // OCC miss: another writer bumped version between our SELECT and
        // the function call. For layout saves the semantic is last-writer-
        // wins, so retry once with the freshly-read version. On a second
        // miss we surface the conflict to the caller.
    }

    m_lastError = QStringLiteral(
        "saveSensoryLayout: version contention (concurrent writers); retry exhausted");
    return false;
}

QString DatabaseManager::loadCumulativeLayout() const
{
    return getSetting(QString::fromLatin1(kCumulativeLayoutKey));
}

bool DatabaseManager::saveCumulativeLayout(const QString& layoutJson)
{
    return setSetting(QString::fromLatin1(kCumulativeLayoutKey), layoutJson);
}

// ============================================================================
//  Natural-key session lookup (v2.0.6 bulk)
// ============================================================================

// Shared implementation for the two bulk variants — the only
// thing that differs between sensory_sessions and detailed_sensory_sessions
// is the table name. SQL is "SELECT … FROM <table> s JOIN (VALUES …) AS
// k(session_name, tester_name, date) ON s.session_name = k.session_name
// AND s.tester_name = k.tester_name AND s.date = k.date". Chunked at 200
// keys per query so we never approach libpq's ~32k parameter ceiling and
// each round-trip stays bounded.
static QVector<DatabaseManager::SessionKeyMatch>
findSessionsByKeysImpl(QSqlDatabase& db,
                       const QString& table,
                       const QVector<DatabaseManager::NaturalKey>& keys,
                       QString& lastError)
{
    QVector<DatabaseManager::SessionKeyMatch> out;
    if (keys.isEmpty()) return out;

    constexpr int kChunk = 200;
    out.reserve(keys.size());

    for (int start = 0; start < keys.size(); start += kChunk) {
        const int end = qMin(start + kChunk, keys.size());

        QStringList tuples;
        tuples.reserve(end - start);
        for (int i = start; i < end; ++i) tuples << QStringLiteral("(?, ?, ?)");

        const QString sql = QStringLiteral(
            "SELECT k.session_name, k.tester_name, k.date, s.id, s.version "
            "FROM %1 s "
            "JOIN (VALUES %2) AS k(session_name, tester_name, date) "
            "  ON s.session_name = k.session_name "
            " AND s.tester_name = k.tester_name "
            " AND s.date::text  = k.date")
            .arg(table, tuples.join(QStringLiteral(", ")));

        QSqlQuery q(db);
        if (!q.prepare(sql)) {
            lastError = QStringLiteral("findSessionsByKeys(prepare %1): ")
                            .arg(table) + q.lastError().text();
            return out;
        }
        for (int i = start; i < end; ++i) {
            q.addBindValue(keys[i].sessionName);
            q.addBindValue(keys[i].testerName);
            q.addBindValue(keys[i].date);
        }
        if (!q.exec()) {
            lastError = QStringLiteral("findSessionsByKeys(exec %1): ")
                            .arg(table) + q.lastError().text();
            return out;
        }
        while (q.next()) {
            DatabaseManager::SessionKeyMatch m;
            m.sessionName = q.value(0).toString();
            m.testerName  = q.value(1).toString();
            m.date        = q.value(2).toString();
            m.id          = q.value(3).toLongLong();
            m.version     = q.value(4).toInt();
            out.append(m);
        }
    }
    return out;
}

QVector<DatabaseManager::SessionKeyMatch>
DatabaseManager::findSensorySessionsByKeys(const QVector<NaturalKey>& keys) const
{
    m_lastError.clear();
    if (!m_online || !isOpen()) return {};
    return findSessionsByKeysImpl(m_pg->queryDb(),
        QStringLiteral("sensory_sessions"), keys, m_lastError);
}

QVector<DatabaseManager::SessionKeyMatch>
DatabaseManager::findDetailedSensorySessionsByKeys(const QVector<NaturalKey>& keys) const
{
    m_lastError.clear();
    if (!m_online || !isOpen()) return {};
    return findSessionsByKeysImpl(m_pg->queryDb(),
        QStringLiteral("detailed_sensory_sessions"), keys, m_lastError);
}

// ============================================================================
//  Sensory header presets (v2.0.4 QoL)
// ============================================================================

bool DatabaseManager::saveSensoryHeaderPresets(const QString& testName,
                                               const QString& media,
                                               const QStringList& sampleNames)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("saveSensoryHeaderPresets: database not open");
        return false;
    }

    // v2.4.8 FIX: write on an ISOLATED autocommit connection (runIsolatedWrite),
    // NOT the shared connection. Previously these INSERTs ran on m_pg's shared
    // connection; on a flaky network that connection reconnects mid-session, and
    // a preset INSERT that happened to ride an uncommitted transaction was
    // silently dropped on the rollback -- exec() returned success but the row
    // never landed ("Save Test Headers doesn't stick"). A fresh connection is in
    // autocommit, so each INSERT commits immediately and durably, immune to the
    // shared connection's state.
    //
    // ON CONFLICT inference target matches the live test-scoped expression index
    // (kind, value, COALESCE(test_name, '')) — ensureSchema() heals older DBs to
    // this shape. sample_name rows carry the owning test_name so the dropdown can
    // be scoped per test; test_name/media rows stay global with a NULL test_name.
    const QString who = writerUuid(m_identity);
    int newRows = 0;
    const bool ok = runIsolatedWrite([&](QSqlDatabase& db) -> bool {
        QSqlQuery q(db);
        q.prepare("INSERT INTO sensory_header_presets (kind, value, test_name, created_by, updated_by) "
                  "VALUES (?, ?, ?, ?, ?) "
                  "ON CONFLICT (kind, value, COALESCE(test_name, '')) DO NOTHING");
        // Typed NULL so QPSQL sends SQL NULL (not '') for the global kinds.
        const QVariant nullTestName(QMetaType(QMetaType::QString));
        auto insertOne = [&](const QString& kind, const QString& rawValue,
                             const QVariant& testNameOrNull) -> bool {
            const QString value = rawValue.trimmed();
            if (value.isEmpty()) return true;  // silently skip empty
            q.bindValue(0, kind);
            q.bindValue(1, value);
            q.bindValue(2, testNameOrNull);
            q.bindValue(3, who);
            q.bindValue(4, who);
            if (!q.exec()) {
                m_lastError = QStringLiteral("saveSensoryHeaderPresets(%1=%2): ")
                                  .arg(kind, value) + q.lastError().text();
                return false;
            }
            if (q.numRowsAffected() > 0) newRows += q.numRowsAffected();
            return true;
        };
        if (!insertOne(QStringLiteral("test_name"), testName, nullTestName)) return false;
        if (!insertOne(QStringLiteral("media"),     media,    nullTestName)) return false;
        // sample_name rows are scoped to the test they were entered under; a blank
        // test title falls back to the global pool (NULL test_name).
        const QVariant scope = testName.trimmed().isEmpty() ? nullTestName
                                                            : QVariant(testName.trimmed());
        for (const QString& name : sampleNames) {
            if (!insertOne(QStringLiteral("sample_name"), name, scope)) return false;
        }
        return true;
    });
    qInfo().noquote()
        << QStringLiteral("[presets] save test_name='%1' media='%2' samples=%3 -> %4 (%5 new)")
               .arg(testName, media).arg(sampleNames.size())
               .arg(ok ? QStringLiteral("ok") : QStringLiteral("FAILED")).arg(newRows);
    return ok;
}

bool DatabaseManager::deleteSensoryHeaderPreset(const QString& kind,
                                                const QString& value,
                                                const QString& testName)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("deleteSensoryHeaderPreset: database not open");
        return false;
    }
    const QString scoped = testName.trimmed();
    const bool ok = runIsolatedWrite([&](QSqlDatabase& db) -> bool {
        QSqlQuery q(db);
        if (scoped.isEmpty()) {
            // Global kinds (test_name/media) and globally-pooled sample names
            // carry a NULL test_name; match that precisely.
            q.prepare("DELETE FROM sensory_header_presets "
                      "WHERE kind = ? AND value = ? AND test_name IS NULL");
            q.addBindValue(kind);
            q.addBindValue(value);
        } else {
            q.prepare("DELETE FROM sensory_header_presets "
                      "WHERE kind = ? AND value = ? AND test_name = ?");
            q.addBindValue(kind);
            q.addBindValue(value);
            q.addBindValue(scoped);
        }
        if (!q.exec()) {
            m_lastError = QStringLiteral("deleteSensoryHeaderPreset(%1=%2): ")
                              .arg(kind, value) + q.lastError().text();
            return false;
        }
        return true;
    });
    qInfo().noquote() << QStringLiteral("[presets] delete %1='%2' scope=%3 -> %4")
                             .arg(kind, value, scoped.isEmpty() ? QStringLiteral("(global)") : scoped)
                             .arg(ok ? QStringLiteral("ok") : QStringLiteral("FAILED"));
    return ok;
}

bool DatabaseManager::runIsolatedWrite(const std::function<bool(QSqlDatabase&)>& body) const
{
    const QString name = QStringLiteral("dve_preset_write_%1")
                             .arg(QUuid::createUuid().toString(QUuid::WithoutBraces));
    bool ok = false;
    {
        if (QSqlDatabase::contains(name)) QSqlDatabase::removeDatabase(name);
        QSqlDatabase db = QSqlDatabase::addDatabase(QStringLiteral("QPSQL"), name);
        db.setHostName(m_cfg.host);
        db.setPort(m_cfg.port);
        db.setDatabaseName(m_cfg.database);
        db.setUserName(m_cfg.user);            // DbConfig::user (NOT username)
        db.setPassword(m_cfg.password);
        db.setConnectOptions(QStringLiteral("connect_timeout=5;") + pgSharedConnectOptions());
        if (!db.open()) {
            m_lastError = QStringLiteral("runIsolatedWrite(open): ") + db.lastError().text();
        } else {
            applyPgSessionSettings(db);        // fresh connection = autocommit
            ok = body(db);                     // each exec commits immediately
            db.close();
        }
        db = QSqlDatabase();
    }
    QSqlDatabase::removeDatabase(name);
    return ok;
}

bool DatabaseManager::beginSharedTransactionForTesting()
{
    if (!isOpen()) return false;
    return QSqlQuery(m_pg->queryDb()).exec(QStringLiteral("BEGIN"));
}

bool DatabaseManager::rollbackSharedTransactionForTesting()
{
    if (!isOpen()) return false;
    return QSqlQuery(m_pg->queryDb()).exec(QStringLiteral("ROLLBACK"));
}

QStringList DatabaseManager::loadSensoryHeaderPresets(const QString& kind) const
{
    m_lastError.clear();
    if (!m_online || !isOpen()) return {};

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT value FROM sensory_header_presets "
              "WHERE kind = ? ORDER BY lower(value)");
    q.addBindValue(kind);
    if (!q.exec()) {
        // Pre-migration installs: table doesn't exist. Surface in
        // m_lastError for diagnostics but return empty so the UI can
        // fall back to a plain edit.
        m_lastError = QStringLiteral("loadSensoryHeaderPresets: ")
                      + q.lastError().text();
        return {};
    }
    QStringList out;
    while (q.next()) out.append(q.value(0).toString());
    return out;
}

QStringList DatabaseManager::loadSampleNamesForTest(const QString& testName) const
{
    if (!m_online || !isOpen()) return {};

    QSqlDatabase& db = m_pg->queryDb();
    QSqlQuery q(db);
    q.prepare("SELECT value FROM sensory_header_presets "
              "WHERE kind = 'sample_name' AND test_name = ? "
              "ORDER BY lower(value)");
    q.addBindValue(testName);
    if (!q.exec()) return {};  // pre-migration / missing column → empty fallback
    QStringList out;
    while (q.next()) out.append(q.value(0).toString());
    return out;
}

// ============================================================================
//  Settings key/value store
// ============================================================================

bool DatabaseManager::setSetting(const QString& key, const QString& value)
{
    m_lastError.clear();
    if (!m_online) {
        m_lastError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return false;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("setSetting: database not open");
        return false;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("INSERT INTO settings (key, value, updated_by) VALUES (?, ?, ?) "
              "ON CONFLICT (key) DO UPDATE SET "
              "value = EXCLUDED.value, "
              "updated_by = EXCLUDED.updated_by");
    q.addBindValue(key);
    q.addBindValue(value);
    q.addBindValue(writerUuid(m_identity));
    if (!q.exec()) {
        m_lastError = QStringLiteral("setSetting(UPSERT): ")
                      + q.lastError().text();
        return false;
    }
    return true;
}

QString DatabaseManager::getSetting(const QString& key, const QString& defaultVal) const
{
    m_lastError.clear();
    if (!m_online) {
        if (m_snapshot && m_snapshot->isOpen()) {
            return m_snapshot->getSetting(key, defaultVal);
        }
        m_lastError = QStringLiteral("DatabaseManager is offline and no snapshot is set");
        return defaultVal;
    }
    if (!isOpen()) {
        m_lastError = QStringLiteral("getSetting: database not open");
        return defaultVal;
    }
    QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT value FROM settings WHERE key = ?");
    q.addBindValue(key);
    if (!q.exec()) {
        m_lastError = QStringLiteral("getSetting(SELECT): ")
                      + q.lastError().text();
        return defaultVal;
    }
    if (q.next()) return q.value(0).toString();
    return defaultVal;
}

} // namespace DVE

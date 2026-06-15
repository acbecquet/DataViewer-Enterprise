#pragma once

#include <QString>

class QSqlDatabase;

namespace DVE {

struct DbConfig {
    QString host;
    int     port = 5432;
    QString database;
    QString user;
    QString password;
};

class ConfigLoader {
public:
    static bool load(const QString& path, DbConfig& out, QString* err = nullptr);

    static QString encryptPassword(const QString& plaintext);
    static QString decryptPassword(const QString& base64Cipher);

private:
    static QByteArray deriveKey();
};

// ── v2.4.2 Tier-1 transport hardening (shared by every QPSQL connection the
//    app opens: PostgresConnection, LiveSyncWorker, MigrationTool) ───────────
//
// Shared libpq connect-string options. EVERY value is space-free on purpose:
// QPSQL forwards setConnectOptions to libpq by replacing ';' with ' ', so a
// value containing a space (e.g. "DataViewer 2.4.2", or "options=-c
// statement_timeout=...") would be split into a bogus extra keyword. The version
// therefore rides as "DataViewer/<ver>" (slash) and statement_timeout is applied
// separately via applyPgSessionSettings(). Putting application_name + keepalives
// + tcp_user_timeout in the connect STRING (not a post-connect SET) makes them
// reconnect-durable: every (re)open rebuilds the connection from this string.
//   * application_name=DataViewer/<ver> — server-side era stamping reads this.
//   * keepalives + tcp_user_timeout      — the OS tears down a black-holed
//     half-open socket (GFW / sleeping NAS) in ~15s instead of the multi-minute
//     TCP default, surfacing it as a connection-shaped error the reconnect logic
//     already handles instead of hanging a thread.
QString pgSharedConnectOptions();

// Apply session GUCs that can't ride the connect string. Call once right after
// a successful db.open() on every QPSQL connection. Reconnect-durable because
// every (re)open path calls it. Best-effort: returns false on failure but never
// throws — a missing statement_timeout degrades to pre-v2.4.2 behavior, not a
// broken connection.
bool applyPgSessionSettings(QSqlDatabase& db);

} // namespace DVE

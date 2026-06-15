#include "ConfigLoader.h"

#include <QFile>
#include <QSettings>
#include <QCryptographicHash>
#include <QSysInfo>
#include <QByteArray>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>

// Test builds compile this TU without the app's -DDVE_APP_VERSION define.
// Fall back so application_name is still well-formed ("DataViewer/0.0.0-dev").
#ifndef DVE_APP_VERSION
#define DVE_APP_VERSION "0.0.0-dev"
#endif

namespace DVE {

// XOR-with-derived-key cipher. Not strong crypto. Purpose: keep db.conf from
// being casually readable. v1.1 swaps for QAESEncryption.
static QByteArray xorCipher(const QByteArray& in, const QByteArray& key) {
    QByteArray out(in);
    const int kn = key.size();
    if (kn == 0) return out;
    for (int i = 0; i < out.size(); ++i) {
        out[i] = out[i] ^ key[i % kn];
    }
    return out;
}

QByteArray ConfigLoader::deriveKey() {
    // Bind the cipher key to this workstation + Windows user. Copying db.conf
    // to another machine yields a different key — decryption returns garbage.
    QByteArray seed = QSysInfo::machineUniqueId();
    seed += qgetenv("USERNAME");
    seed += "DataViewerDbConfigSalt-v1";
    return QCryptographicHash::hash(seed, QCryptographicHash::Sha256);
}

QString ConfigLoader::encryptPassword(const QString& plaintext) {
    const QByteArray key = deriveKey();
    const QByteArray ct  = xorCipher(plaintext.toUtf8(), key);
    return QString::fromUtf8(ct.toBase64());
}

QString ConfigLoader::decryptPassword(const QString& base64Cipher) {
    const QByteArray key = deriveKey();
    const QByteArray ct  = QByteArray::fromBase64(base64Cipher.toUtf8());
    const QByteArray pt  = xorCipher(ct, key);
    return QString::fromUtf8(pt);
}

bool ConfigLoader::load(const QString& path, DbConfig& out, QString* err) {
    if (!QFile::exists(path)) {
        if (err) *err = QStringLiteral("db.conf not found at ") + path;
        return false;
    }
    QSettings ini(path, QSettings::IniFormat);
    ini.beginGroup("postgres");
    out.host     = ini.value("host").toString();
    out.port     = ini.value("port", 5432).toInt();
    out.database = ini.value("database").toString();
    out.user     = ini.value("user").toString();
    const QString enc = ini.value("password_encrypted").toString();
    ini.endGroup();

    if (out.host.isEmpty()) {
        if (err) *err = QStringLiteral("[postgres] host is missing");
        return false;
    }
    if (out.database.isEmpty()) {
        if (err) *err = QStringLiteral("[postgres] database is missing");
        return false;
    }
    if (out.user.isEmpty()) {
        if (err) *err = QStringLiteral("[postgres] user is missing");
        return false;
    }
    if (enc.isEmpty()) {
        if (err) *err = QStringLiteral("[postgres] password_encrypted is missing");
        return false;
    }
    out.password = decryptPassword(enc);
    return true;
}

QString pgSharedConnectOptions() {
    return QStringLiteral(
        "application_name=DataViewer/" DVE_APP_VERSION ";"
        "keepalives=1;keepalives_idle=10;keepalives_interval=5;keepalives_count=3;"
        "tcp_user_timeout=15000");
}

bool applyPgSessionSettings(QSqlDatabase& db) {
    QSqlQuery q(db);
    if (q.exec(QStringLiteral("SET statement_timeout = 10000"))) return true;
    // Best-effort, but never silent: a lost statement_timeout is the exact
    // protection this adds against a hung-query UI freeze, so surface it.
    qWarning() << "applyPgSessionSettings: SET statement_timeout failed --"
               << q.lastError().text();
    return false;
}

} // namespace DVE

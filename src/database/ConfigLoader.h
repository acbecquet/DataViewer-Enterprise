#pragma once

#include <QString>

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

} // namespace DVE

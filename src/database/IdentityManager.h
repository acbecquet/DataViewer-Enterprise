#pragma once

#include <QObject>
#include <QString>
#include <QStringList>
#include <QUuid>

namespace DVE {

class IdentityManager : public QObject {
    Q_OBJECT
public:
    explicit IdentityManager(QObject* parent = nullptr);

    QUuid   uuid() const;
    QString displayName() const;
    QString color() const;

    void setDisplayName(const QString& name);
    void setColor(const QString& hex);

    bool firstLaunchPending() const;

    static QStringList defaultColorPalette();

private:
    void ensureUuid();
};

} // namespace DVE

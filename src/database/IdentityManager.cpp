#include "IdentityManager.h"
#include <QSettings>
#include <QUuid>

namespace DVE {

namespace {
constexpr auto kKeyUuid  = "identity/uuid";
constexpr auto kKeyName  = "identity/displayName";
constexpr auto kKeyColor = "identity/color";
}

IdentityManager::IdentityManager(QObject* parent) : QObject(parent) {
    ensureUuid();
}

void IdentityManager::ensureUuid() {
    QSettings s;
    const QString existing = s.value(kKeyUuid).toString();
    if (existing.isEmpty() || QUuid(existing).isNull()) {
        const QUuid fresh = QUuid::createUuid();
        s.setValue(kKeyUuid, fresh.toString(QUuid::WithoutBraces));
    }
}

QUuid IdentityManager::uuid() const {
    QSettings s;
    return QUuid(s.value(kKeyUuid).toString());
}

QString IdentityManager::displayName() const {
    QSettings s;
    return s.value(kKeyName).toString();
}

QString IdentityManager::color() const {
    QSettings s;
    return s.value(kKeyColor).toString();
}

void IdentityManager::setDisplayName(const QString& name) {
    QSettings s;
    s.setValue(kKeyName, name);
}

void IdentityManager::setColor(const QString& hex) {
    QSettings s;
    s.setValue(kKeyColor, hex);
}

bool IdentityManager::firstLaunchPending() const {
    return displayName().isEmpty() || color().isEmpty();
}

QStringList IdentityManager::defaultColorPalette() {
    // 12 hand-picked hex codes, balanced for light + dark UI themes.
    return {
        "#ef4444", // red
        "#f97316", // orange
        "#eab308", // amber
        "#84cc16", // lime
        "#10b981", // emerald
        "#06b6d4", // cyan
        "#3b82f6", // blue
        "#6366f1", // indigo
        "#8b5cf6", // violet
        "#d946ef", // fuchsia
        "#ec4899", // pink
        "#64748b"  // slate
    };
}

} // namespace DVE

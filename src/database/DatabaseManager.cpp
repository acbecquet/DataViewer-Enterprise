#include "DatabaseManager.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "ConfigLoader.h"

#include <QDebug>

namespace DVE {

DatabaseManager::DatabaseManager(QObject* parent)
    : QObject(parent), m_pg(new PostgresConnection(this)) {}

DatabaseManager::~DatabaseManager() { close(); }

bool DatabaseManager::open(const DbConfig& cfg, IdentityManager* identity) {
    m_identity = identity;
    if (!m_pg->open(cfg)) {
        m_lastError = m_pg->lastError();
        m_open = false;
        return false;
    }
    m_lastError.clear();
    m_open = true;
    return true;
}

void DatabaseManager::close() {
    if (m_pg) m_pg->close();
    m_open = false;
}

bool DatabaseManager::isOpen() const { return m_open; }

QString DatabaseManager::currentPath() const { return QString(); }

void DatabaseManager::logDebug(const QString& msg) const {
    qDebug().noquote() << "[DatabaseManager]" << msg;
}

// -- Stubs (implemented in sub-batches 3b-3d) --------------------------------

bool DatabaseManager::saveFile(const FileResult&) {
    m_lastError = QStringLiteral("saveFile not yet implemented (3b)");
    return false;
}

bool DatabaseManager::hasFile(const QString&) const { return false; }

FileResult DatabaseManager::loadFile(int) const { return FileResult(); }

FileResult DatabaseManager::loadFileByPath(const QString&) const { return FileResult(); }

QVector<FileRecord> DatabaseManager::listFiles() const { return {}; }

bool DatabaseManager::removeFile(int) { return false; }

int DatabaseManager::deduplicateFiles(int) { return 0; }

QStringList DatabaseManager::recentFilePaths() const { return {}; }

bool DatabaseManager::saveSensorySession(const SensorySession&) { return false; }

bool DatabaseManager::saveSensorySession(SensorySession&) { return false; }

QVector<SensorySession> DatabaseManager::loadSensorySessions() const { return {}; }

SensorySession DatabaseManager::loadSensorySession(int) const { return SensorySession(); }

QVector<SensoryRecord> DatabaseManager::listSensoryRecords() const { return {}; }

bool DatabaseManager::removeSensorySession(int) { return false; }

QString DatabaseManager::nextDefaultTestName() const { return QStringLiteral("Test 1"); }

bool DatabaseManager::saveDetailedSensorySession(const DetailedSensorySession&) { return false; }

QVector<DetailedSensorySession> DatabaseManager::loadDetailedSensorySessions() const { return {}; }

DetailedSensorySession DatabaseManager::loadDetailedSensorySession(int) const { return DetailedSensorySession(); }

QVector<DetailedSensoryRecord> DatabaseManager::listDetailedSensoryRecords() const { return {}; }

bool DatabaseManager::removeDetailedSensorySession(int) { return false; }

QString DatabaseManager::loadSensoryLayout(int) const { return QString(); }

bool DatabaseManager::saveSensoryLayout(int, const QString&) { return false; }

QString DatabaseManager::loadCumulativeLayout() const { return QString(); }

bool DatabaseManager::saveCumulativeLayout(const QString&) { return false; }

bool DatabaseManager::setSetting(const QString&, const QString&) { return false; }

QString DatabaseManager::getSetting(const QString&, const QString& defaultVal) const {
    return defaultVal;
}

} // namespace DVE

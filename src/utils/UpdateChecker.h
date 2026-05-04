#pragma once

#include <QObject>
#include <QTimer>
#include <QVersionNumber>

namespace DVE {

class UpdateChecker : public QObject
{
    Q_OBJECT

public:
    explicit UpdateChecker(QObject* parent = nullptr);

    /// Call once after the main window is shown. Runs an immediate check
    /// then starts the hourly timer.
    void start();

    /// Snapshot of the update folder for diagnostics / self-test.
    struct ProbeResult {
        QString        updateRoot;
        bool           rootExists = false;
        QVersionNumber latest;
        QString        installerPath;
    };
    static ProbeResult probe();

private slots:
    void check();

private:
    static QString        updateRoot();
    static QVersionNumber latestAvailable(QString* installerPathOut = nullptr);
    bool  isSuppressed() const;
    void  showDialog(const QVersionNumber& latest, const QString& installerPath);

    QTimer* m_timer;
    bool    m_suppressedThisSession = false;
};

} // namespace DVE

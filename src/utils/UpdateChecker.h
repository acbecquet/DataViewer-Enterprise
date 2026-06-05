#pragma once

#include <functional>

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

    /// Install a callback invoked synchronously immediately before the
    /// "Update Now" handler hard-exits via std::_Exit(0). That hard exit
    /// bypasses closeEvent and every destructor (identical to a crash), so
    /// the rolling recovery snapshot's last debounced flush may be stale.
    /// The hook (set by MainWindow to flush the recovery snapshot) lets us
    /// persist a complete, current snapshot before the process terminates.
    void setPreExitHook(std::function<void()> hook);

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
    std::function<void()> m_preExitHook;
};

} // namespace DVE

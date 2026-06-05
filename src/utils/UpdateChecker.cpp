#include "UpdateChecker.h"

#include <cstdlib>
#include <QApplication>
#include <QComboBox>
#include <QDate>
#include <QDialog>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QFutureWatcher>
#include <QHBoxLayout>
#include <QLabel>
#include <QMessageBox>
#include <QProcess>
#include <QPushButton>
#include <QSettings>
#include <QVBoxLayout>
#include <QtConcurrent/QtConcurrent>

namespace DVE {

static const int  kCheckIntervalMs = 60 * 60 * 1000; // 1 hour
static const char kSettingsKey[]   = "update/suppressedUntil";

// ── Construction ──────────────────────────────────────────────────────────────

UpdateChecker::UpdateChecker(QObject* parent)
    : QObject(parent)
    , m_timer(new QTimer(this))
{
    m_timer->setInterval(kCheckIntervalMs);
    connect(m_timer, &QTimer::timeout, this, &UpdateChecker::check);
}

void UpdateChecker::start()
{
    check();           // immediate check on startup
    m_timer->start();  // then every hour
}

void UpdateChecker::setPreExitHook(std::function<void()> hook)
{
    m_preExitHook = std::move(hook);
}

// ── Version discovery ─────────────────────────────────────────────────────────

QString UpdateChecker::updateRoot()
{
    const QString rel = "/SynologyDrive/SDR/Device Group/Software Release/Current";
    const QString homePath = QDir::homePath() + rel;
    if (QDir(homePath).exists())
        return homePath;
    // Fallback: root-level Synology Drive install (some machines)
    const QString rootPath = QDir::rootPath() + "SynologyDrive/SDR/Device Group/Software Release/Current";
    return rootPath;
}

QVersionNumber UpdateChecker::latestAvailable(QString* installerPathOut)
{
    const QString rootPath = updateRoot();
    QDir root(rootPath);
    qInfo().noquote() << "scanning update root:" << rootPath;
    if (!root.exists()) {
        qInfo().noquote() << "update root does not exist — bailing";
        return {};
    }

    QVersionNumber best;
    QString        bestInstaller;

    const QStringList entries = root.entryList(QDir::Dirs | QDir::NoDotAndDotDot);
    qInfo().noquote() << "found" << entries.size() << "subfolder(s)";

    for (const QString& rawEntry : entries) {
        // v2.1.0: trim whitespace in case Synology sync left trailing
        // spaces. Strip optional leading 'v' / 'V'. Reject anything that
        // doesn't parse cleanly as a dotted version number — including
        // trailing suffixes like "-rc1".
        QString entry = rawEntry.trimmed();
        QString name  = entry;
        if (name.startsWith(QLatin1Char('v')) || name.startsWith(QLatin1Char('V')))
            name = name.mid(1);

        qsizetype suffixIdx = 0;
        QVersionNumber v = QVersionNumber::fromString(name, &suffixIdx);
        if (v.isNull() || suffixIdx != name.size()) {
            qInfo().noquote()
                << "  skip" << rawEntry << "→ not a version folder";
            continue;
        }

        // v2.1.0: case-insensitive installer lookup. Synology sometimes
        // normalises filenames during sync ("DataViewer-Setup.exe"), and
        // QFile::exists is case-sensitive on the Synology Drive client's
        // virtual filesystem in some configurations.
        QString installer = root.filePath(entry + "/DataViewer-setup.exe");
        if (!QFile::exists(installer)) {
            const QDir subDir(root.filePath(entry));
            // QDir::entryList does case-insensitive name matching by
            // default on Windows, which is what we want here.
            const QStringList exes = subDir.entryList(
                {QStringLiteral("DataViewer*setup*.exe")},
                QDir::Files);
            if (exes.isEmpty()) {
                qInfo().noquote()
                    << "  skip" << rawEntry << "v" << v.toString()
                    << "→ no DataViewer-setup.exe in folder";
                continue;
            }
            installer = subDir.filePath(exes.first());
        }

        qInfo().noquote()
            << "  candidate" << rawEntry << "→ v" << v.toString()
            << "installer=" << QFileInfo(installer).fileName();

        if (v > best) {
            best          = v;
            bestInstaller = installer;
        }
    }

    if (best.isNull())
        qInfo().noquote() << "no valid candidate found";
    else
        qInfo().noquote() << "best candidate: v" << best.toString()
                                    << "at" << bestInstaller;

    if (installerPathOut)
        *installerPathOut = bestInstaller;
    return best;
}

UpdateChecker::ProbeResult UpdateChecker::probe()
{
    ProbeResult r;
    r.updateRoot = updateRoot();
    r.rootExists = QDir(r.updateRoot).exists();
    if (r.rootExists)
        r.latest = latestAvailable(&r.installerPath);
    return r;
}

// ── Suppression ───────────────────────────────────────────────────────────────

bool UpdateChecker::isSuppressed() const
{
    QSettings s("SDR", "DataViewerEnterprise");
    const QString val = s.value(kSettingsKey).toString();
    if (val.isEmpty())
        return false;

    const QDate until = QDate::fromString(val, Qt::ISODate);
    if (!until.isValid())
        return false;

    if (QDate::currentDate() <= until)
        return true;

    // Past date — clean up stale entry
    QSettings("SDR", "DataViewerEnterprise").remove(kSettingsKey);
    return false;
}

// ── Check logic ───────────────────────────────────────────────────────────────

void UpdateChecker::check()
{
    if (m_suppressedThisSession)
        return;

    // Scan the Synology share on a background thread to avoid blocking the UI
    auto* watcher = new QFutureWatcher<QPair<QVersionNumber, QString>>(this);

    connect(watcher, &QFutureWatcher<QPair<QVersionNumber, QString>>::finished,
            this, [this, watcher]() {
        watcher->deleteLater();

        if (m_suppressedThisSession)
            return;

        const QPair<QVersionNumber, QString> result = watcher->result();
        const QVersionNumber& latest    = result.first;
        const QString&        installer = result.second;

        if (latest.isNull() || installer.isEmpty()) {
            qInfo().noquote() << "check: no update available";
            return;
        }

        const QVersionNumber current =
            QVersionNumber::fromString(QApplication::applicationVersion());
        qInfo().noquote() << "check: latest=" << latest.toString()
                                    << "current=" << current.toString();
        if (QVersionNumber::compare(latest, current) <= 0) {
            qInfo().noquote() << "  already at latest — no dialog";
            return;
        }

        if (isSuppressed()) {
            qInfo().noquote() << "  suppressed (Remind Me) — no dialog";
            return;
        }

        qInfo().noquote() << "  showing update dialog";
        showDialog(latest, installer);
    });

    watcher->setFuture(QtConcurrent::run([]() -> QPair<QVersionNumber, QString> {
        QString installerPath;
        const QVersionNumber v = UpdateChecker::latestAvailable(&installerPath);
        return qMakePair(v, installerPath);
    }));
}

// ── Dialog ────────────────────────────────────────────────────────────────────

void UpdateChecker::showDialog(const QVersionNumber& latest,
                               const QString&         installerPath)
{
    QDialog dlg;
    dlg.setWindowTitle("Update Available");
    dlg.setWindowFlags(dlg.windowFlags() & ~Qt::WindowContextHelpButtonHint);
    dlg.setMinimumWidth(420);

    QVBoxLayout* vl = new QVBoxLayout(&dlg);
    vl->setSpacing(8);
    vl->setContentsMargins(16, 16, 16, 12);

    // Message
    QLabel* msg = new QLabel(
        QString("DataViewer Enterprise <b>v%1</b> is available.<br>"
                "You are running v%2.")
            .arg(latest.toString(), QApplication::applicationVersion()));
    msg->setTextFormat(Qt::RichText);
    vl->addWidget(msg);
    vl->addSpacing(4);

    // Remind-in-X-days row
    QHBoxLayout* suppressRow = new QHBoxLayout;
    suppressRow->addWidget(new QLabel("Remind me in:"));
    QComboBox* daysPicker = new QComboBox;
    for (int d : {1, 3, 7, 14, 30})
        daysPicker->addItem(QString("%1 day%2").arg(d).arg(d == 1 ? "" : "s"), d);
    daysPicker->setCurrentIndex(2); // default: 7 days
    suppressRow->addWidget(daysPicker);
    suppressRow->addStretch();
    vl->addLayout(suppressRow);
    vl->addSpacing(8);

    // Buttons
    QHBoxLayout* btnRow = new QHBoxLayout;
    QPushButton* updateBtn = new QPushButton("Update Now");
    QPushButton* remindBtn = new QPushButton("Remind Me");
    QPushButton* laterBtn  = new QPushButton("Later This Session");
    updateBtn->setDefault(true);
    btnRow->addWidget(updateBtn);
    btnRow->addWidget(remindBtn);
    btnRow->addStretch();
    btnRow->addWidget(laterBtn);
    vl->addLayout(btnRow);

    // Update Now: launch installer, then exit hard so the running process
    // doesn't keep the .exe locked / keep the in-memory version stale.
    // QApplication::quit() alone races: if any modal child is open or the
    // main window is mid-shutdown, the install can complete while we're
    // still alive, and the update prompt re-fires next poll because our
    // applicationVersion() in memory is the OLD value.
    connect(updateBtn, &QPushButton::clicked, &dlg, [&]() {
        dlg.accept();
        if (!QProcess::startDetached(installerPath)) {
            QMessageBox::warning(nullptr, "Update Failed",
                "Could not launch the installer.\nPath: " + installerPath);
            return;
        }
        // Give the installer a moment to spawn and grab the AppMutex.
        QApplication::processEvents();
        // std::_Exit below bypasses closeEvent and all destructors (same as
        // a crash), so do ONE synchronous, complete recovery flush first --
        // otherwise the last debounced rolling snapshot could be up to ~2 s
        // stale and the user loses work across the update.
        if (m_preExitHook)
            m_preExitHook();
        // Hard exit. The installer was launched detached; AppMutex on the
        // installer side now blocks until we're gone, then the install
        // proceeds and (per RestartApplications=yes) re-launches us.
        std::_Exit(0);
    });

    // Remind Me: persist suppression date, dismiss
    connect(remindBtn, &QPushButton::clicked, &dlg, [&]() {
        const int days = daysPicker->currentData().toInt();
        QSettings("SDR", "DataViewerEnterprise").setValue(
            kSettingsKey,
            QDate::currentDate().addDays(days).toString(Qt::ISODate));
        dlg.reject();
    });

    // Later: suppress for this session only
    connect(laterBtn, &QPushButton::clicked, &dlg, [&]() {
        m_suppressedThisSession = true;
        dlg.reject();
    });

    dlg.exec();
}

} // namespace DVE

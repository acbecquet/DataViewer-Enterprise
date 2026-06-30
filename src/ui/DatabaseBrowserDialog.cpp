#include "DatabaseBrowserDialog.h"
#include "SensoryPanel.h"
#include "DetailedSensoryPanel.h"
#include "../utils/AppTheme.h"
#include "../utils/OutputPaths.h"
#include "../database/CompatClassifier.h"
#include "../database/DbRepair.h"
#include "../pipeline/DataProcessor.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QHeaderView>
#include <QFrame>
#include <QGraphicsDropShadowEffect>
#include <QMessageBox>
#include <QFileDialog>
#include <QFileInfo>
#include <QMap>
#include <QHash>
#include <QSet>
#include <QDateTime>
#include <QToolButton>
#include <QMenu>
#include <QAction>
#include <QApplication>
#include <QEventLoop>
#include <QProgressDialog>

namespace DVE {

namespace {

// Tolerant date parse for era inference (only consulted for UNSTAMPED rows).
// Files carry an ISO added_at/loaded_at; sensory rows carry a free-form date
// string, so try ISO first then a few common display formats. Returns an
// invalid QDate when nothing parses (classifyEra then buckets to "unstamped").
QDate parseCreationDate(const QString& s)
{
    if (s.trimmed().isEmpty()) return {};
    QDateTime dt = QDateTime::fromString(s, Qt::ISODate);
    if (dt.isValid()) return dt.date();
    QDate d = QDate::fromString(s, Qt::ISODate);
    if (d.isValid()) return d;
    for (const char* fmt : {"yyyy-MM-dd", "yyyy/MM/dd", "MM/dd/yyyy",
                            "M/d/yyyy", "dd-MM-yyyy", "dd/MM/yyyy"}) {
        d = QDate::fromString(s, QLatin1String(fmt));
        if (d.isValid()) return d;
    }
    return {};
}

CompatClass classifyFileRec(const FileRecord& r)
{
    const QString src = r.addedAt.isEmpty() ? r.loadedAt : r.addedAt;
    CompatClass c = classifyEra(r.appVersion, parseCreationDate(src));
    // FileRecord has no cheap missing-regime signal (deferred in SP4-T2), so
    // health here is the no-samples flag only.
    c.health = fileHealth(r.sampleCount, /*missingRegimes=*/false);
    return c;
}

// A DB row is junk-ish when it has no real content: no tester/assessor and a
// blank or default session name. (isPlaceholderSession's id<0 gate can't apply
// here — every browsed row is already persisted — so we mirror only its
// "no header content + default name" half, layered on the sampleCount<=0 trigger
// inside sensoryHealth.)
bool looksLikePlaceholder(const QString& tester, const QString& assessor,
                          const QString& sessionName)
{
    const QString sn = sessionName.trimmed();
    return tester.trimmed().isEmpty()
        && assessor.trimmed().isEmpty()
        && (sn.isEmpty() || sn == QStringLiteral("New Session"));
}

CompatClass classifySensoryRec(const SensoryRecord& r)
{
    CompatClass c = classifyEra(r.appVersion, parseCreationDate(r.date));
    c.health = sensoryHealth(r.hasLegacyStringScores,
                             looksLikePlaceholder(r.testerName, r.assessorName, r.sessionName),
                             r.sampleCount);
    return c;
}

CompatClass classifyDetailedRec(const DetailedSensoryRecord& r)
{
    CompatClass c = classifyEra(r.appVersion, parseCreationDate(r.date));
    c.health = sensoryHealth(r.hasLegacyStringScores,
                             looksLikePlaceholder(r.testerName, r.assessorName, r.sessionName),
                             r.sampleCount);
    return c;
}

// Display string for the Version column ("v2.2.x" plus an "(approx.)" hint when
// the era was inferred from the row's date rather than a stamp).
QString eraDisplay(const CompatClass& c)
{
    return c.approx ? c.eraLabel + QStringLiteral(" (approx.)") : c.eraLabel;
}

// Display string for the Health column.
QString healthDisplay(const CompatClass& c)
{
    return c.isHealthy() ? QStringLiteral("Healthy") : c.health.join(QStringLiteral(", "));
}

// A row passes the version/health filter when no buckets are checked, or it
// matches at least one checked era bucket OR health flag (OR within the dropdown;
// the caller AND-combines this with the free-text search).
bool rowPassesVersionFilter(const CompatClass& c, const QSet<QString>& active)
{
    if (active.isEmpty()) return true;
    if (active.contains(c.eraLabel)) return true;
    for (const QString& h : c.health)
        if (active.contains(h)) return true;
    return false;
}

// SP4 A5: build a capped, human-readable list of the rows a delete will remove,
// for the confirm dialog (so the user sees exactly what goes, name + era, not
// just a count). Expands parent (group) rows to their children; de-dupes when a
// parent and one of its children are both selected. Caps at maxLines.
QString describeRowsForDelete(const QList<QTreeWidgetItem*>& items,
                              int versionCol, int maxLines = 15)
{
    QStringList lines;
    int extra = 0;
    QSet<QTreeWidgetItem*> seen;
    auto addLine = [&](QTreeWidgetItem* it) {
        if (seen.contains(it)) return;
        seen.insert(it);
        QString line = it->text(0).trimmed();
        const QString ver = (versionCol >= 0) ? it->text(versionCol).trimmed() : QString();
        if (!ver.isEmpty()) line += QStringLiteral("  (%1)").arg(ver);
        if (lines.size() < maxLines) lines << QStringLiteral("  • ") + line;
        else ++extra;
    };
    for (QTreeWidgetItem* item : items) {
        if (item->data(0, Qt::UserRole).toInt() > 0) {
            addLine(item);                               // leaf row
        } else {
            for (int c = 0; c < item->childCount(); ++c) // group row -> children
                addLine(item->child(c));
        }
    }
    QString out = lines.join(QLatin1Char('\n'));
    if (extra > 0) out += QStringLiteral("\n  …and %1 more").arg(extra);
    return out;
}

} // namespace

DatabaseBrowserDialog::DatabaseBrowserDialog(DatabaseManager* db, QWidget* parent)
    : QDialog(parent)
    , m_db(db)
{
    setWindowTitle("Database Browser");
    setMinimumSize(560, 380);
    resize(1200, 650);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(8);

    // ── Tab widget ───────────────────────────────────────────────────────────
    m_tabWidget = new QTabWidget(this);
    mainLayout->addWidget(m_tabWidget, 1);

    // ══════════════════════════════════════════════════════════════════════════
    // TPM Data tab
    // ══════════════════════════════════════════════════════════════════════════
    {
        auto* tpmTab = new QWidget;
        auto* tpmLayout = new QVBoxLayout(tpmTab);
        tpmLayout->setSpacing(6);

        // Filter bar
        auto* filterRow = new QHBoxLayout;
        filterRow->addWidget(new QLabel("Search:"));
        m_filterEdit = new QLineEdit(tpmTab);
        m_filterEdit->setPlaceholderText("Filter by file name…");
        m_filterEdit->setClearButtonEnabled(true);
        m_filterEdit->setMinimumHeight(28);
        filterRow->addWidget(m_filterEdit);
        m_tpmFilterBtn = new QToolButton(tpmTab);
        m_tpmFilterBtn->setText("Version");
        m_tpmFilterBtn->setToolButtonStyle(Qt::ToolButtonTextOnly);
        m_tpmFilterBtn->setPopupMode(QToolButton::InstantPopup);
        m_tpmFilterBtn->setMinimumHeight(28);
        m_tpmFilterBtn->setToolTip("Filter by app-version era and data health");
        m_tpmFilterMenu = new QMenu(m_tpmFilterBtn);
        m_tpmFilterBtn->setMenu(m_tpmFilterMenu);
        filterRow->addWidget(m_tpmFilterBtn);
        auto* refreshBtn = new QPushButton("Refresh", tpmTab);
        filterRow->addWidget(refreshBtn);
        tpmLayout->addLayout(filterRow);

        // File tree
        m_tree = new QTreeWidget(tpmTab);
        m_tree->setColumnCount(7);
        m_tree->setHeaderLabels({"File Name", "Loaded At", "Template", "Tests", "Samples", "Version", "Health"});
        m_tree->setSelectionBehavior(QAbstractItemView::SelectRows);
        m_tree->setSelectionMode(QAbstractItemView::ExtendedSelection);
        m_tree->setAlternatingRowColors(true);
        m_tree->setRootIsDecorated(false);
        m_tree->setAnimated(true);
        m_tree->setUniformRowHeights(true);
        m_tree->header()->setDefaultAlignment(Qt::AlignLeft | Qt::AlignVCenter);
        m_tree->header()->setSectionResizeMode(0, QHeaderView::Stretch);
        for (int c = 1; c < m_tree->columnCount(); ++c)
            m_tree->header()->setSectionResizeMode(c, QHeaderView::ResizeToContents);
        m_tree->header()->setStretchLastSection(false);
        m_tree->setSortingEnabled(true);
        m_tree->sortByColumn(1, Qt::DescendingOrder);  // newest first by default
        tpmLayout->addWidget(m_tree);

        // Status label
        m_statusLabel = new QLabel(tpmTab);
        tpmLayout->addWidget(m_statusLabel);

        // Action bar
        auto* tpmActionBar = new QFrame(tpmTab);
        tpmActionBar->setStyleSheet(QString(
            "QFrame { background-color: %1; border-top: 1px solid %2; padding: 8px 12px; }"
        ).arg(AppTheme::surfaceApp().name(), AppTheme::borderDefault().name()));
        auto* btnRow = new QHBoxLayout(tpmActionBar);
        btnRow->setContentsMargins(0, 0, 0, 0);
        btnRow->setSpacing(6);

        m_loadBtn = new QPushButton("Load Selected", tpmActionBar);
        m_loadBtn->setEnabled(false);
        m_loadBtn->setProperty("primary", true);
        m_loadBtn->setIcon(AppTheme::icon("folder-open"));
        btnRow->addWidget(m_loadBtn);

        m_loadAllBtn = new QPushButton("Load All Visible", tpmActionBar);
        btnRow->addWidget(m_loadAllBtn);

        m_deleteBtn = new QPushButton("Delete Selected", tpmActionBar);
        m_deleteBtn->setEnabled(false);
        m_deleteBtn->setProperty("destructive", true);
        m_deleteBtn->setIcon(AppTheme::icon("x"));
        btnRow->addWidget(m_deleteBtn);

        m_tpmRepairBtn = new QPushButton("Repair Incomplete…", tpmActionBar);
        m_tpmRepairBtn->setToolTip("Re-read the source .xlsx for TPM files saved without their detailed\n"
                                   "rows (the pre-v2.0 data-row gap) and backfill them. Scans all files;\n"
                                   "files whose source .xlsx can't be located are skipped.");
        btnRow->addWidget(m_tpmRepairBtn);

        auto* cleanupBtn = new QPushButton("Cleanup Duplicates", tpmActionBar);
        cleanupBtn->setToolTip("Remove unknown/corrupt entries and keep only 3 most recent per file");
        btnRow->addWidget(cleanupBtn);

        btnRow->addStretch();

        for (auto* b : {m_loadBtn, m_deleteBtn}) {
            b->style()->unpolish(b);
            b->style()->polish(b);
        }

        tpmLayout->addWidget(tpmActionBar);

        // Connections
        connect(refreshBtn,    &QPushButton::clicked, this, &DatabaseBrowserDialog::onRefresh);
        connect(m_filterEdit,  &QLineEdit::textChanged, this, &DatabaseBrowserDialog::onFilterChanged);
        connect(m_loadBtn,     &QPushButton::clicked, this, &DatabaseBrowserDialog::onLoad);
        connect(m_loadAllBtn,  &QPushButton::clicked, this, &DatabaseBrowserDialog::onLoadAll);
        connect(m_deleteBtn,   &QPushButton::clicked, this, &DatabaseBrowserDialog::onDelete);
        connect(cleanupBtn,    &QPushButton::clicked, this, &DatabaseBrowserDialog::onCleanup);
        connect(m_tpmRepairBtn,&QPushButton::clicked, this, &DatabaseBrowserDialog::onRepairTpm);
        connect(m_tree,        &QTreeWidget::itemSelectionChanged, this, &DatabaseBrowserDialog::onSelectionChanged);
        connect(m_tree,        &QTreeWidget::itemDoubleClicked, this, &DatabaseBrowserDialog::onItemDoubleClicked);
        connect(m_tree,        &QTreeWidget::itemExpanded, this, &DatabaseBrowserDialog::onItemExpanded);

        m_tabWidget->addTab(tpmTab, "TPM Data");
    }

    // ══════════════════════════════════════════════════════════════════════════
    // Sensory Data tab
    // ══════════════════════════════════════════════════════════════════════════
    {
        auto* sensoryTab = new QWidget;
        auto* sensoryLayout = new QVBoxLayout(sensoryTab);
        sensoryLayout->setSpacing(6);

        // Filter bar
        auto* filterRow = new QHBoxLayout;
        filterRow->addWidget(new QLabel("Search:"));
        m_sensoryFilterEdit = new QLineEdit(sensoryTab);
        m_sensoryFilterEdit->setPlaceholderText("Filter by session name, assessor, or media…");
        m_sensoryFilterEdit->setClearButtonEnabled(true);
        m_sensoryFilterEdit->setMinimumHeight(28);
        filterRow->addWidget(m_sensoryFilterEdit);
        m_sensoryFilterBtn = new QToolButton(sensoryTab);
        m_sensoryFilterBtn->setText("Version");
        m_sensoryFilterBtn->setToolButtonStyle(Qt::ToolButtonTextOnly);
        m_sensoryFilterBtn->setPopupMode(QToolButton::InstantPopup);
        m_sensoryFilterBtn->setMinimumHeight(28);
        m_sensoryFilterBtn->setToolTip("Filter by app-version era and data health");
        m_sensoryFilterMenu = new QMenu(m_sensoryFilterBtn);
        m_sensoryFilterBtn->setMenu(m_sensoryFilterMenu);
        filterRow->addWidget(m_sensoryFilterBtn);
        auto* refreshBtn = new QPushButton("Refresh", sensoryTab);
        filterRow->addWidget(refreshBtn);
        sensoryLayout->addLayout(filterRow);

        // Sensory tree (hierarchical: Test Title → Tester entries)
        m_sensoryTree = new QTreeWidget(sensoryTab);
        m_sensoryTree->setColumnCount(7);
        m_sensoryTree->setHeaderLabels({"Test Title / Tester", "Assessor", "Media", "Date", "Samples", "Version", "Health"});
        m_sensoryTree->setSelectionBehavior(QAbstractItemView::SelectRows);
        m_sensoryTree->setSelectionMode(QAbstractItemView::ExtendedSelection);
        m_sensoryTree->setAlternatingRowColors(true);
        m_sensoryTree->setRootIsDecorated(true);
        m_sensoryTree->setUniformRowHeights(true);
        m_sensoryTree->header()->setDefaultAlignment(Qt::AlignLeft | Qt::AlignVCenter);
        m_sensoryTree->header()->setSectionResizeMode(0, QHeaderView::Stretch);
        for (int c = 1; c < m_sensoryTree->columnCount(); ++c)
            m_sensoryTree->header()->setSectionResizeMode(c, QHeaderView::ResizeToContents);
        m_sensoryTree->header()->setStretchLastSection(false);
        m_sensoryTree->setSortingEnabled(true);
        m_sensoryTree->sortByColumn(3, Qt::DescendingOrder);  // newest first by date
        sensoryLayout->addWidget(m_sensoryTree);

        // Status label
        m_sensoryStatusLabel = new QLabel(sensoryTab);
        sensoryLayout->addWidget(m_sensoryStatusLabel);

        // Action bar
        auto* sensoryActionBar = new QFrame(sensoryTab);
        sensoryActionBar->setStyleSheet(QString(
            "QFrame { background-color: %1; border-top: 1px solid %2; padding: 8px 12px; }"
        ).arg(AppTheme::surfaceApp().name(), AppTheme::borderDefault().name()));
        auto* btnRow = new QHBoxLayout(sensoryActionBar);
        btnRow->setContentsMargins(0, 0, 0, 0);
        btnRow->setSpacing(6);

        m_sensoryLoadBtn = new QPushButton("Open Selected", sensoryActionBar);
        m_sensoryLoadBtn->setEnabled(false);
        m_sensoryLoadBtn->setProperty("primary", true);
        m_sensoryLoadBtn->setIcon(AppTheme::icon("folder-open"));
        btnRow->addWidget(m_sensoryLoadBtn);

        m_sensoryDeleteBtn = new QPushButton("Delete Selected", sensoryActionBar);
        m_sensoryDeleteBtn->setEnabled(false);
        m_sensoryDeleteBtn->setProperty("destructive", true);
        m_sensoryDeleteBtn->setIcon(AppTheme::icon("x"));
        btnRow->addWidget(m_sensoryDeleteBtn);

        m_sensoryRepairBtn = new QPushButton("Repair Legacy Scores", sensoryActionBar);
        m_sensoryRepairBtn->setEnabled(false);
        m_sensoryRepairBtn->setToolTip("Normalize legacy string-typed scores back to numbers (lossless).\n"
                                       "Acts on every legacy row in the database, not just the selection.");
        btnRow->addWidget(m_sensoryRepairBtn);

        m_sensoryReportBtn = new QPushButton("Generate Combined Report...", sensoryActionBar);
        m_sensoryReportBtn->setEnabled(false);
        m_sensoryReportBtn->setToolTip("Generate a PPTX report with one slide per selected session");
        btnRow->addWidget(m_sensoryReportBtn);

        btnRow->addStretch();

        for (auto* b : {m_sensoryLoadBtn, m_sensoryDeleteBtn}) {
            b->style()->unpolish(b);
            b->style()->polish(b);
        }

        sensoryLayout->addWidget(sensoryActionBar);

        // Connections
        connect(refreshBtn,          &QPushButton::clicked, this, &DatabaseBrowserDialog::onRefresh);
        connect(m_sensoryFilterEdit, &QLineEdit::textChanged, this, [this](const QString& text) {
            populateSensoryTree(text);
        });
        connect(m_sensoryLoadBtn,    &QPushButton::clicked, this, [this]() {
            m_selectedSensoryIds.clear();
            const auto items = m_sensoryTree->selectedItems();
            for (QTreeWidgetItem* item : items) {
                int id = idFromItem(item);
                if (id > 0) {
                    m_selectedSensoryIds.append(id);
                } else if (item->childCount() > 0) {
                    // Parent item: collect all children
                    for (int c = 0; c < item->childCount(); ++c) {
                        int cid = idFromItem(item->child(c));
                        if (cid > 0) m_selectedSensoryIds.append(cid);
                    }
                }
            }
            if (!m_selectedSensoryIds.isEmpty()) {
                m_sensorySelection = true;
                accept();
            }
        });
        connect(m_sensoryDeleteBtn,  &QPushButton::clicked, this, &DatabaseBrowserDialog::onSensoryDelete);
        connect(m_sensoryRepairBtn, &QPushButton::clicked, this, &DatabaseBrowserDialog::onRepairLegacyScores);
        connect(m_sensoryReportBtn, &QPushButton::clicked, this, &DatabaseBrowserDialog::onSensoryGenerateReport);
        connect(m_sensoryTree,       &QTreeWidget::itemSelectionChanged, this, &DatabaseBrowserDialog::onSensorySelectionChanged);
        connect(m_sensoryTree,       &QTreeWidget::itemDoubleClicked, this, &DatabaseBrowserDialog::onSensoryDoubleClicked);

        m_tabWidget->addTab(sensoryTab, "Sensory Data");
    }

    // ══════════════════════════════════════════════════════════════════════════
    // Detailed Sensory Data tab
    // ══════════════════════════════════════════════════════════════════════════
    {
        auto* detSensTab = new QWidget;
        auto* detSensLayout = new QVBoxLayout(detSensTab);
        detSensLayout->setSpacing(6);

        // Filter bar
        auto* filterRow = new QHBoxLayout;
        filterRow->addWidget(new QLabel("Search:"));
        m_detSensFilterEdit = new QLineEdit(detSensTab);
        m_detSensFilterEdit->setPlaceholderText("Filter by session name, tester, assessor, or media…");
        m_detSensFilterEdit->setClearButtonEnabled(true);
        m_detSensFilterEdit->setMinimumHeight(28);
        filterRow->addWidget(m_detSensFilterEdit);
        m_detSensFilterBtn = new QToolButton(detSensTab);
        m_detSensFilterBtn->setText("Version");
        m_detSensFilterBtn->setToolButtonStyle(Qt::ToolButtonTextOnly);
        m_detSensFilterBtn->setPopupMode(QToolButton::InstantPopup);
        m_detSensFilterBtn->setMinimumHeight(28);
        m_detSensFilterBtn->setToolTip("Filter by app-version era and data health");
        m_detSensFilterMenu = new QMenu(m_detSensFilterBtn);
        m_detSensFilterBtn->setMenu(m_detSensFilterMenu);
        filterRow->addWidget(m_detSensFilterBtn);
        auto* refreshBtn = new QPushButton("Refresh", detSensTab);
        filterRow->addWidget(refreshBtn);
        detSensLayout->addLayout(filterRow);

        // Tree widget
        m_detSensTree = new QTreeWidget(detSensTab);
        m_detSensTree->setColumnCount(7);
        m_detSensTree->setHeaderLabels({"Test Title / Tester", "Assessor", "Media", "Date", "Samples", "Version", "Health"});
        m_detSensTree->setSelectionBehavior(QAbstractItemView::SelectRows);
        m_detSensTree->setSelectionMode(QAbstractItemView::ExtendedSelection);
        m_detSensTree->setAlternatingRowColors(true);
        m_detSensTree->setRootIsDecorated(true);
        m_detSensTree->setUniformRowHeights(true);
        m_detSensTree->header()->setDefaultAlignment(Qt::AlignLeft | Qt::AlignVCenter);
        m_detSensTree->header()->setSectionResizeMode(0, QHeaderView::Stretch);
        for (int c = 1; c < m_detSensTree->columnCount(); ++c)
            m_detSensTree->header()->setSectionResizeMode(c, QHeaderView::ResizeToContents);
        m_detSensTree->header()->setStretchLastSection(false);
        m_detSensTree->setSortingEnabled(true);
        m_detSensTree->sortByColumn(3, Qt::DescendingOrder);  // newest first by date
        detSensLayout->addWidget(m_detSensTree);

        // Status label
        m_detSensStatusLabel = new QLabel(detSensTab);
        detSensLayout->addWidget(m_detSensStatusLabel);

        // Action bar
        auto* detSensActionBar = new QFrame(detSensTab);
        detSensActionBar->setStyleSheet(QString(
            "QFrame { background-color: %1; border-top: 1px solid %2; padding: 8px 12px; }"
        ).arg(AppTheme::surfaceApp().name(), AppTheme::borderDefault().name()));
        auto* btnRow = new QHBoxLayout(detSensActionBar);
        btnRow->setContentsMargins(0, 0, 0, 0);
        btnRow->setSpacing(6);

        m_detSensLoadBtn = new QPushButton("Open Selected", detSensActionBar);
        m_detSensLoadBtn->setEnabled(false);
        m_detSensLoadBtn->setProperty("primary", true);
        m_detSensLoadBtn->setIcon(AppTheme::icon("folder-open"));
        btnRow->addWidget(m_detSensLoadBtn);

        m_detSensDeleteBtn = new QPushButton("Delete Selected", detSensActionBar);
        m_detSensDeleteBtn->setEnabled(false);
        m_detSensDeleteBtn->setProperty("destructive", true);
        m_detSensDeleteBtn->setIcon(AppTheme::icon("x"));
        btnRow->addWidget(m_detSensDeleteBtn);

        m_detSensRepairBtn = new QPushButton("Repair Legacy Scores", detSensActionBar);
        m_detSensRepairBtn->setEnabled(false);
        m_detSensRepairBtn->setToolTip("Normalize legacy string-typed scores back to numbers (lossless).\n"
                                       "Acts on every legacy row in the database, not just the selection.");
        btnRow->addWidget(m_detSensRepairBtn);

        m_detSensReportBtn = new QPushButton("Generate Combined Report...", detSensActionBar);
        m_detSensReportBtn->setEnabled(false);
        m_detSensReportBtn->setToolTip("Generate a PPTX report with selected detailed sensory sessions");
        btnRow->addWidget(m_detSensReportBtn);

        btnRow->addStretch();

        for (auto* b : {m_detSensLoadBtn, m_detSensDeleteBtn}) {
            b->style()->unpolish(b);
            b->style()->polish(b);
        }

        detSensLayout->addWidget(detSensActionBar);

        // Connections
        connect(refreshBtn, &QPushButton::clicked, this, &DatabaseBrowserDialog::onRefresh);
        connect(m_detSensFilterEdit, &QLineEdit::textChanged, this, [this](const QString& text) {
            populateDetailedSensoryTree(text);
        });
        connect(m_detSensLoadBtn, &QPushButton::clicked, this, [this]() {
            m_selectedDetSensIds.clear();
            const auto items = m_detSensTree->selectedItems();
            for (QTreeWidgetItem* item : items) {
                int id = idFromItem(item);
                if (id > 0) {
                    m_selectedDetSensIds.append(id);
                } else if (item->childCount() > 0) {
                    for (int c = 0; c < item->childCount(); ++c) {
                        int cid = idFromItem(item->child(c));
                        if (cid > 0) m_selectedDetSensIds.append(cid);
                    }
                }
            }
            if (!m_selectedDetSensIds.isEmpty()) {
                m_detSensSelection = true;
                accept();
            }
        });
        connect(m_detSensDeleteBtn, &QPushButton::clicked, this, &DatabaseBrowserDialog::onDetailedSensoryDelete);
        connect(m_detSensRepairBtn, &QPushButton::clicked, this, &DatabaseBrowserDialog::onRepairLegacyScores);
        connect(m_detSensReportBtn, &QPushButton::clicked, this, &DatabaseBrowserDialog::onDetailedSensoryGenerateReport);
        connect(m_detSensTree, &QTreeWidget::itemSelectionChanged, this, &DatabaseBrowserDialog::onDetailedSensorySelectionChanged);
        connect(m_detSensTree, &QTreeWidget::itemDoubleClicked, this, [this](QTreeWidgetItem* item, int) {
            if (!item) return;
            m_selectedDetSensIds.clear();
            int id = idFromItem(item);
            if (id > 0) {
                m_selectedDetSensIds.append(id);
            } else if (item->childCount() > 0) {
                for (int c = 0; c < item->childCount(); ++c) {
                    int cid = idFromItem(item->child(c));
                    if (cid > 0) m_selectedDetSensIds.append(cid);
                }
            }
            if (!m_selectedDetSensIds.isEmpty()) {
                m_detSensSelection = true;
                accept();
            }
        });

        m_tabWidget->addTab(detSensTab, "Detailed Sensory Data");
    }

    // ── Close button (shared) ────────────────────────────────────────────────
    auto* bottomRow = new QHBoxLayout;
    bottomRow->addStretch();
    auto* closeBtn = new QPushButton("Close", this);
    bottomRow->addWidget(closeBtn);
    mainLayout->addLayout(bottomRow);

    connect(closeBtn, &QPushButton::clicked, this, &QDialog::reject);
    connect(m_tabWidget, &QTabWidget::currentChanged, this, &DatabaseBrowserDialog::onTabChanged);

    // Drop shadow elevation
    auto* shadow = new QGraphicsDropShadowEffect(this);
    shadow->setBlurRadius(16);
    shadow->setOffset(0, 4);
    shadow->setColor(QColor(0, 0, 0, 30));
    setGraphicsEffect(shadow);

    // Initial load
    onRefresh();
}

void DatabaseBrowserDialog::onRefresh()
{
    m_allRecords = m_db->listFiles();
    m_sensoryRecords = m_db->listSensoryRecords();
    m_detSensRecords = m_db->listDetailedSensoryRecords();

    // SP4 A4: recompute per-era / per-health bucket counts over ALL loaded rows
    // and rebuild each tab's Version ▾ dropdown. Counts are over every DB row
    // (pre-dedup) — for triage you want to see every stale/legacy row, not just
    // the de-duplicated ones the tree shows.
    {
        QMap<QString, int> era, health;
        for (const FileRecord& r : m_allRecords) {
            const CompatClass c = classifyFileRec(r);
            ++era[c.eraLabel];
            for (const QString& h : c.health) ++health[h];
        }
        rebuildVersionMenu(m_tpmFilterMenu, m_tpmFilterBtn, m_tpmActiveFilters,
                           era, health, [this] { populateTree(m_filterEdit->text()); });
    }
    {
        QMap<QString, int> era, health;
        for (const SensoryRecord& r : m_sensoryRecords) {
            const CompatClass c = classifySensoryRec(r);
            ++era[c.eraLabel];
            for (const QString& h : c.health) ++health[h];
        }
        rebuildVersionMenu(m_sensoryFilterMenu, m_sensoryFilterBtn, m_sensoryActiveFilters,
                           era, health, [this] { populateSensoryTree(m_sensoryFilterEdit->text()); });
        if (m_sensoryRepairBtn)
            m_sensoryRepairBtn->setEnabled(health.value(QStringLiteral("Legacy string scores"), 0) > 0);
    }
    {
        QMap<QString, int> era, health;
        for (const DetailedSensoryRecord& r : m_detSensRecords) {
            const CompatClass c = classifyDetailedRec(r);
            ++era[c.eraLabel];
            for (const QString& h : c.health) ++health[h];
        }
        rebuildVersionMenu(m_detSensFilterMenu, m_detSensFilterBtn, m_detSensActiveFilters,
                           era, health, [this] { populateDetailedSensoryTree(m_detSensFilterEdit->text()); });
        if (m_detSensRepairBtn)
            m_detSensRepairBtn->setEnabled(health.value(QStringLiteral("Legacy string scores"), 0) > 0);
    }

    populateTree(m_filterEdit->text());
    populateSensoryTree(m_sensoryFilterEdit->text());
    if (m_detSensFilterEdit) populateDetailedSensoryTree(m_detSensFilterEdit->text());
}

void DatabaseBrowserDialog::rebuildVersionMenu(QMenu* menu, QToolButton* btn, QSet<QString>& active,
                                               const QMap<QString, int>& eraCounts,
                                               const QMap<QString, int>& healthCounts,
                                               const std::function<void()>& repopulate)
{
    if (!menu || !btn) return;

    // Drop any checked key whose bucket no longer exists (e.g. every row of that
    // era was deleted) — otherwise the filter could get stuck hiding everything
    // against a key that has no action left to un-check it.
    for (auto it = active.begin(); it != active.end(); ) {
        if (!eraCounts.contains(*it) && !healthCounts.contains(*it)) it = active.erase(it);
        else ++it;
    }

    menu->clear();

    auto addEntry = [&](const QString& key, int count) {
        QAction* act = menu->addAction(QStringLiteral("%1  (%2)").arg(key).arg(count));
        act->setCheckable(true);
        act->setChecked(active.contains(key));   // set BEFORE connect — no spurious fire
        connect(act, &QAction::toggled, this, [this, key, &active, btn, repopulate](bool on) {
            if (on) active.insert(key); else active.remove(key);
            btn->setText(active.isEmpty() ? QStringLiteral("Version")
                                          : QStringLiteral("Version (%1)").arg(active.size()));
            repopulate();
        });
    };

    if (eraCounts.isEmpty() && healthCounts.isEmpty()) {
        QAction* empty = menu->addAction(QStringLiteral("(no records)"));
        empty->setEnabled(false);
    } else {
        for (auto it = eraCounts.constBegin(); it != eraCounts.constEnd(); ++it)
            addEntry(it.key(), it.value());
        if (!eraCounts.isEmpty() && !healthCounts.isEmpty())
            menu->addSeparator();
        for (auto it = healthCounts.constBegin(); it != healthCounts.constEnd(); ++it)
            addEntry(it.key(), it.value());
    }

    btn->setText(active.isEmpty() ? QStringLiteral("Version")
                                  : QStringLiteral("Version (%1)").arg(active.size()));
}

void DatabaseBrowserDialog::onFilterChanged(const QString& text)
{
    populateTree(text);
}

void DatabaseBrowserDialog::onTabChanged(int)
{
    // Nothing special needed; both trees are always populated
}

void DatabaseBrowserDialog::populateTree(const QString& filter)
{
    m_tree->clear();

    // F6 (v2.5.0): re-adding the same .xlsx later mints a NEW versioned row
    // (file_path, added_at) instead of overwriting the prior one, so a path can
    // now legitimately appear several times. We show EVERY version distinctly —
    // dropping older rows here would re-hide exactly the history the user asked
    // us to keep. To disambiguate, when a path has more than one version we
    // append the add-time to the displayed file name; single-version paths keep
    // the plain name. listFiles() returns rows ordered by added_at DESC.
    QVector<const FileRecord*> shown;
    QVector<CompatClass> shownClass;          // parallel to `shown`
    QHash<QString, int> versionsPerPath;

    for (const FileRecord& rec : m_allRecords) {
        if (!filter.isEmpty() &&
            !rec.fileName.contains(filter, Qt::CaseInsensitive) &&
            !rec.filePath.contains(filter, Qt::CaseInsensitive)) {
            continue;
        }
        const CompatClass klass = classifyFileRec(rec);
        if (!rowPassesVersionFilter(klass, m_tpmActiveFilters)) continue;
        ++versionsPerPath[rec.filePath];
        shown.append(&rec);
        shownClass.append(klass);
    }

    auto fmtDate = [](const QString& iso) -> QString {
        QDateTime dt = QDateTime::fromString(iso, Qt::ISODate);
        return dt.isValid() ? dt.toString("yyyy-MM-dd  HH:mm") : iso;
    };

    int totalSamples = 0;
    int versionedPaths = 0;
    QSet<QString> countedMultiPath;
    for (int i = 0; i < shown.size(); ++i) {
        const FileRecord* rec = shown[i];
        const CompatClass& klass = shownClass[i];
        const bool versioned = versionsPerPath.value(rec->filePath) > 1;
        if (versioned && !countedMultiPath.contains(rec->filePath)) {
            countedMultiPath.insert(rec->filePath);
            ++versionedPaths;
        }
        // Prefer the add-time stamp (the version identity); fall back to
        // loaded_at for legacy rows whose added_at is empty/NULL.
        const QString stampSrc = rec->addedAt.isEmpty() ? rec->loadedAt
                                                        : rec->addedAt;
        QString displayName = rec->fileName;
        if (versioned)
            displayName += QString("  (added %1)").arg(fmtDate(stampSrc));

        auto* topItem = new QTreeWidgetItem(m_tree);
        topItem->setText(0, displayName);
        topItem->setText(1, fmtDate(stampSrc));
        topItem->setText(2, rec->templateVersion);
        topItem->setText(3, QString::number(rec->sheetCount));
        topItem->setText(4, QString::number(rec->sampleCount));
        topItem->setText(5, eraDisplay(klass));
        topItem->setText(6, healthDisplay(klass));
        topItem->setData(0, Qt::UserRole, rec->id);
        totalSamples += rec->sampleCount;
    }

    QString status = QString("%1 file version%2  |  %3 total samples")
                         .arg(shown.size())
                         .arg(shown.size() == 1 ? "" : "s")
                         .arg(totalSamples);
    if (versionedPaths > 0) {
        status += QString("  (%1 file%2 with multiple versions)")
                      .arg(versionedPaths)
                      .arg(versionedPaths == 1 ? "" : "s");
    }
    m_statusLabel->setText(status);
}

void DatabaseBrowserDialog::populateSensoryTree(const QString& filter)
{
    m_sensoryTree->clear();

    // Group records by test title (parent = test, children = each tester's
    // session). Within each (test, tester) cell, only the newest row is
    // shown — older duplicates are silently hidden. listSensoryRecords()
    // returns rows ORDER BY id DESC, so the first record we see for a
    // (testTitle, testerName) pair is always the newest.
    struct Group {
        QString title;
        QVector<const SensoryRecord*> testers;     // newest per testerName
        QVector<CompatClass> testerClass;          // parallel to testers
        QSet<QString> seenTesterKeys;              // (testerName) lookup
    };
    QMap<QString, Group> groups;
    QStringList groupOrder;
    int hiddenDuplicates = 0;
    int totalSessions = 0;
    int totalSamples = 0;

    for (const SensoryRecord& rec : m_sensoryRecords) {
        if (!filter.isEmpty() &&
            !rec.testTitle.contains(filter, Qt::CaseInsensitive) &&
            !rec.testerName.contains(filter, Qt::CaseInsensitive) &&
            !rec.assessorName.contains(filter, Qt::CaseInsensitive) &&
            !rec.media.contains(filter, Qt::CaseInsensitive) &&
            !rec.sessionName.contains(filter, Qt::CaseInsensitive)) {
            continue;
        }

        const CompatClass klass = classifySensoryRec(rec);
        if (!rowPassesVersionFilter(klass, m_sensoryActiveFilters)) continue;

        const QString groupKey   = rec.testTitle.isEmpty() ? rec.sessionName : rec.testTitle;
        const QString testerKey  = rec.testerName.toLower().trimmed();
        if (!groups.contains(groupKey)) {
            groups.insert(groupKey, Group{groupKey, {}, {}, {}});
            groupOrder.append(groupKey);
        }
        Group& g = groups[groupKey];
        if (g.seenTesterKeys.contains(testerKey)) {
            ++hiddenDuplicates;
            continue;
        }
        g.seenTesterKeys.insert(testerKey);
        g.testers.append(&rec);
        g.testerClass.append(klass);
        ++totalSessions;
        totalSamples += rec.sampleCount;
    }

    for (const QString& groupKey : groupOrder) {
        const Group& g = groups[groupKey];
        auto* parentItem = new QTreeWidgetItem(m_sensoryTree);
        parentItem->setText(0, g.title);
        parentItem->setText(1, g.testers.first()->assessorName);
        parentItem->setText(2, g.testers.first()->media);
        parentItem->setText(3, g.testers.first()->date);
        parentItem->setText(4, QString::number(g.testers.size()) + " testers");
        parentItem->setData(0, Qt::UserRole, -1);
        QFont boldFont = parentItem->font(0);
        boldFont.setBold(true);
        parentItem->setFont(0, boldFont);

        for (int i = 0; i < g.testers.size(); ++i) {
            const SensoryRecord* rec = g.testers[i];
            const CompatClass& klass = g.testerClass[i];
            auto* child = new QTreeWidgetItem(parentItem);
            QString tester = rec->testerName.isEmpty() ? rec->assessorName : rec->testerName;
            if (tester.isEmpty()) tester = QStringLiteral("(unnamed)");
            child->setText(0, tester);
            child->setText(1, rec->assessorName);
            child->setText(2, rec->media);
            child->setText(3, rec->date);
            child->setText(4, QString::number(rec->sampleCount));
            child->setText(5, eraDisplay(klass));
            child->setText(6, healthDisplay(klass));
            child->setData(0, Qt::UserRole, rec->id);
        }
        parentItem->setExpanded(true);
    }

    QString status = QString("%1 sessions  |  %2 total samples")
                         .arg(totalSessions).arg(totalSamples);
    if (hiddenDuplicates > 0) {
        status += QString("  (%1 stale duplicate%2 hidden)")
                      .arg(hiddenDuplicates)
                      .arg(hiddenDuplicates == 1 ? "" : "s");
    }
    m_sensoryStatusLabel->setText(status);
}

int DatabaseBrowserDialog::idFromItem(QTreeWidgetItem* item) const
{
    return item ? item->data(0, Qt::UserRole).toInt() : -1;
}

void DatabaseBrowserDialog::onSelectionChanged()
{
    bool hasSel = !m_tree->selectedItems().isEmpty();
    m_loadBtn->setEnabled(hasSel);
    m_deleteBtn->setEnabled(hasSel);
}

void DatabaseBrowserDialog::onSensorySelectionChanged()
{
    const auto items = m_sensoryTree->selectedItems();
    int selCount = items.size();

    // Enable load if any items with valid IDs are selected
    bool hasLoadable = false;
    for (QTreeWidgetItem* item : items) {
        int id = idFromItem(item);
        if (id > 0) { hasLoadable = true; break; }
        if (item->childCount() > 0) { hasLoadable = true; break; }
    }

    m_sensoryLoadBtn->setEnabled(hasLoadable);
    m_sensoryDeleteBtn->setEnabled(selCount > 0);
    m_sensoryReportBtn->setEnabled(selCount > 0);
}

void DatabaseBrowserDialog::onItemDoubleClicked(QTreeWidgetItem* item, int)
{
    if (!item) return;
    int id = idFromItem(item);
    if (id > 0) {
        m_selectedIds.clear();
        m_selectedIds.append(id);
        m_sensorySelection = false;
        accept();
    }
}

void DatabaseBrowserDialog::onSensoryDoubleClicked(QTreeWidgetItem* item, int)
{
    if (!item) return;
    m_selectedSensoryIds.clear();
    int id = idFromItem(item);
    if (id > 0) {
        // Leaf item (individual tester)
        m_selectedSensoryIds.append(id);
    } else if (item->childCount() > 0) {
        // Parent item: load all testers under this test
        for (int c = 0; c < item->childCount(); ++c) {
            int cid = idFromItem(item->child(c));
            if (cid > 0) m_selectedSensoryIds.append(cid);
        }
    }
    if (!m_selectedSensoryIds.isEmpty()) {
        m_sensorySelection = true;
        accept();
    }
}

void DatabaseBrowserDialog::onItemExpanded(QTreeWidgetItem*)
{
    m_tree->resizeColumnToContents(0);
}

void DatabaseBrowserDialog::onLoad()
{
    m_selectedIds.clear();
    m_sensorySelection = false;
    const auto items = m_tree->selectedItems();
    for (QTreeWidgetItem* item : items) {
        int id = idFromItem(item);
        if (id > 0)
            m_selectedIds.append(id);
    }
    if (!m_selectedIds.isEmpty())
        accept();
}

void DatabaseBrowserDialog::onLoadAll()
{
    m_selectedIds.clear();
    m_sensorySelection = false;
    for (int i = 0; i < m_tree->topLevelItemCount(); ++i) {
        int id = idFromItem(m_tree->topLevelItem(i));
        if (id > 0)
            m_selectedIds.append(id);
    }
    if (!m_selectedIds.isEmpty())
        accept();
}

void DatabaseBrowserDialog::onDelete()
{
    const auto items = m_tree->selectedItems();
    if (items.isEmpty()) return;

    const QString rowList = describeRowsForDelete(items, /*versionCol=*/5);
    auto answer = QMessageBox::question(this, "Delete Files",
        QString("Delete the following %1 entr%2 from the database?\n\n%3\n\n"
                "This removes the data from the database only — the original\n"
                "Excel files on disk are not affected.")
            .arg(items.size())
            .arg(items.size() == 1 ? "y" : "ies")
            .arg(rowList),
        QMessageBox::Yes | QMessageBox::No);

    if (answer != QMessageBox::Yes) return;

    for (QTreeWidgetItem* item : items) {
        int id = idFromItem(item);
        if (id > 0)
            m_db->removeFile(id);
    }

    onRefresh();
}

static QVector<int> collectSensoryIds(const QList<QTreeWidgetItem*>& items)
{
    QVector<int> ids;
    for (QTreeWidgetItem* item : items) {
        int id = item->data(0, Qt::UserRole).toInt();
        if (id > 0) {
            ids.append(id);
        } else if (item->childCount() > 0) {
            for (int c = 0; c < item->childCount(); ++c) {
                int cid = item->child(c)->data(0, Qt::UserRole).toInt();
                if (cid > 0) ids.append(cid);
            }
        }
    }
    return ids;
}

void DatabaseBrowserDialog::onSensoryDelete()
{
    const auto items = m_sensoryTree->selectedItems();
    if (items.isEmpty()) return;

    QVector<int> ids = collectSensoryIds(items);
    if (ids.isEmpty()) return;

    const QString rowList = describeRowsForDelete(items, /*versionCol=*/5);
    auto answer = QMessageBox::question(this, "Delete Sensory Session",
        QString("Delete the following %1 sensory session%2 from the database?\n\n%3")
            .arg(ids.size())
            .arg(ids.size() == 1 ? "" : "s")
            .arg(rowList),
        QMessageBox::Yes | QMessageBox::No);

    if (answer != QMessageBox::Yes) return;

    for (int id : ids)
        m_db->removeSensorySession(id);

    onRefresh();
}

void DatabaseBrowserDialog::onSensoryGenerateReport()
{
    const auto items = m_sensoryTree->selectedItems();
    if (items.isEmpty()) return;

    // Collect all selected session IDs (including children of parent items)
    QVector<int> ids = collectSensoryIds(items);

    if (ids.isEmpty()) return;

    // Load all sessions
    QVector<SensorySession> sessions;
    for (int id : ids) {
        SensorySession sess = m_db->loadSensorySession(id);
        if (!sess.samples.isEmpty())
            sessions.append(sess);
    }

    if (sessions.isEmpty()) {
        QMessageBox::warning(this, "No Data", "Could not load selected sessions.");
        return;
    }

    // Ask for save path
    const QString sensoryDir = OutputPaths::resolveDir(ReportMode::Sensory, QString());
    QString path = QFileDialog::getSaveFileName(
        this, "Save Combined Sensory Report",
        sensoryDir + "/" + QStringLiteral("Combined_Sensory_report.pptx"),
        "PowerPoint files (*.pptx);;All files (*)");
    if (path.isEmpty()) return;

    QString errorMsg;
    if (SensoryPanel::generateCombinedPptx(sessions, path, errorMsg)) {
        QMessageBox::information(this, "Report Saved",
            QString("Combined report with %1 slide(s) saved successfully.")
                .arg(sessions.size()));
    } else {
        QMessageBox::warning(this, "Report Error",
                             "Could not generate report:\n" + errorMsg);
    }
}

void DatabaseBrowserDialog::onCleanup()
{
    auto answer = QMessageBox::question(this, "Cleanup Duplicates",
        "This will:\n"
        "  - Remove entries with 'unknown' template (empty/corrupt files)\n"
        "  - Keep only the 3 most recent versions per file name\n\n"
        "WARNING: older VERSIONS of a file are permanently deleted, not just\n"
        "exact duplicates. Since TPM files are now kept as a version history\n"
        "(each re-add of a file is a separate dated version), any file with\n"
        "more than 3 versions will lose its oldest ones. This cannot be undone.\n\n"
        "Continue?",
        QMessageBox::Yes | QMessageBox::No);

    if (answer != QMessageBox::Yes) return;

    int removed = m_db->deduplicateFiles(3);

    QMessageBox::information(this, "Cleanup Complete",
        QString("Removed %1 entries from the database.").arg(removed));

    onRefresh();
}

void DatabaseBrowserDialog::populateDetailedSensoryTree(const QString& filter)
{
    m_detSensTree->clear();

    // Same dedup contract as populateSensoryTree: newest row wins per
    // (testTitle, testerName) pair. listDetailedSensoryRecords() returns
    // rows ORDER BY id DESC.
    struct Group {
        QString title;
        QVector<const DetailedSensoryRecord*> testers;
        QVector<CompatClass> testerClass;          // parallel to testers
        QSet<QString> seenTesterKeys;
    };
    QMap<QString, Group> groups;
    QStringList groupOrder;
    int hiddenDuplicates = 0;
    int totalSessions = 0;
    int totalSamples = 0;

    for (const DetailedSensoryRecord& rec : m_detSensRecords) {
        if (!filter.isEmpty() &&
            !rec.testTitle.contains(filter, Qt::CaseInsensitive) &&
            !rec.testerName.contains(filter, Qt::CaseInsensitive) &&
            !rec.assessorName.contains(filter, Qt::CaseInsensitive) &&
            !rec.media.contains(filter, Qt::CaseInsensitive) &&
            !rec.sessionName.contains(filter, Qt::CaseInsensitive)) {
            continue;
        }

        const CompatClass klass = classifyDetailedRec(rec);
        if (!rowPassesVersionFilter(klass, m_detSensActiveFilters)) continue;

        const QString groupKey  = rec.testTitle.isEmpty() ? rec.sessionName : rec.testTitle;
        const QString testerKey = rec.testerName.toLower().trimmed();
        if (!groups.contains(groupKey)) {
            groups.insert(groupKey, Group{groupKey, {}, {}, {}});
            groupOrder.append(groupKey);
        }
        Group& g = groups[groupKey];
        if (g.seenTesterKeys.contains(testerKey)) {
            ++hiddenDuplicates;
            continue;
        }
        g.seenTesterKeys.insert(testerKey);
        g.testers.append(&rec);
        g.testerClass.append(klass);
        ++totalSessions;
        totalSamples += rec.sampleCount;
    }

    for (const QString& groupKey : groupOrder) {
        const Group& g = groups[groupKey];
        auto* parentItem = new QTreeWidgetItem(m_detSensTree);
        parentItem->setText(0, g.title);
        parentItem->setText(1, g.testers.first()->assessorName);
        parentItem->setText(2, g.testers.first()->media);
        parentItem->setText(3, g.testers.first()->date);
        parentItem->setText(4, QString::number(g.testers.size()) + " testers");
        parentItem->setData(0, Qt::UserRole, -1);
        QFont boldFont = parentItem->font(0);
        boldFont.setBold(true);
        parentItem->setFont(0, boldFont);

        for (int i = 0; i < g.testers.size(); ++i) {
            const DetailedSensoryRecord* rec = g.testers[i];
            const CompatClass& klass = g.testerClass[i];
            auto* child = new QTreeWidgetItem(parentItem);
            QString tester = rec->testerName.isEmpty() ? rec->assessorName : rec->testerName;
            if (tester.isEmpty()) tester = QStringLiteral("(unnamed)");
            child->setText(0, tester);
            child->setText(1, rec->assessorName);
            child->setText(2, rec->media);
            child->setText(3, rec->date);
            child->setText(4, QString::number(rec->sampleCount));
            child->setText(5, eraDisplay(klass));
            child->setText(6, healthDisplay(klass));
            child->setData(0, Qt::UserRole, rec->id);
        }
        parentItem->setExpanded(true);
    }

    QString status = QString("%1 sessions  |  %2 total samples")
                         .arg(totalSessions).arg(totalSamples);
    if (hiddenDuplicates > 0) {
        status += QString("  (%1 stale duplicate%2 hidden)")
                      .arg(hiddenDuplicates)
                      .arg(hiddenDuplicates == 1 ? "" : "s");
    }
    m_detSensStatusLabel->setText(status);
}

void DatabaseBrowserDialog::onDetailedSensorySelectionChanged()
{
    const auto items = m_detSensTree->selectedItems();
    bool hasLoadable = false;
    for (QTreeWidgetItem* item : items) {
        int id = idFromItem(item);
        if (id > 0) { hasLoadable = true; break; }
        if (item->childCount() > 0) { hasLoadable = true; break; }
    }
    m_detSensLoadBtn->setEnabled(hasLoadable);
    m_detSensDeleteBtn->setEnabled(!items.isEmpty());
    m_detSensReportBtn->setEnabled(!items.isEmpty());
}

void DatabaseBrowserDialog::onDetailedSensoryDelete()
{
    const auto items = m_detSensTree->selectedItems();
    if (items.isEmpty()) return;

    QVector<int> ids = collectSensoryIds(items);
    if (ids.isEmpty()) return;

    const QString rowList = describeRowsForDelete(items, /*versionCol=*/5);
    auto answer = QMessageBox::question(this, "Delete Detailed Sensory Session",
        QString("Delete the following %1 detailed sensory session%2 from the database?\n\n%3")
            .arg(ids.size())
            .arg(ids.size() == 1 ? "" : "s")
            .arg(rowList),
        QMessageBox::Yes | QMessageBox::No);

    if (answer != QMessageBox::Yes) return;

    for (int id : ids)
        m_db->removeDetailedSensorySession(id);

    onRefresh();
}

void DatabaseBrowserDialog::onDetailedSensoryGenerateReport()
{
    const auto items = m_detSensTree->selectedItems();
    if (items.isEmpty()) return;

    QVector<int> ids = collectSensoryIds(items);
    if (ids.isEmpty()) return;

    QVector<DetailedSensorySession> sessions;
    for (int id : ids) {
        DetailedSensorySession sess = m_db->loadDetailedSensorySession(id);
        if (!sess.samples.isEmpty())
            sessions.append(sess);
    }

    if (sessions.isEmpty()) {
        QMessageBox::warning(this, "No Data", "Could not load selected sessions.");
        return;
    }

    const QString detailedDir = OutputPaths::resolveDir(ReportMode::DetailedSensory, QString());
    QString path = QFileDialog::getSaveFileName(
        this, "Save Combined Detailed Sensory Report",
        detailedDir + "/" + QStringLiteral("Combined_Detailed_Sensory_report.pptx"),
        "PowerPoint files (*.pptx);;All files (*)");
    if (path.isEmpty()) return;

    QString errorMsg;
    if (DetailedSensoryPanel::generateCombinedPptx(sessions, path, errorMsg)) {
        QMessageBox::information(this, "Report Saved",
            QString("Combined detailed sensory report saved successfully."));
    } else {
        QMessageBox::warning(this, "Report Error",
                             "Could not generate report:\n" + errorMsg);
    }
}

// ── SP4 A5: repair actions ───────────────────────────────────────────────────

void DatabaseBrowserDialog::onRepairLegacyScores()
{
    // Lossless, global: normalize every string-typed score back to a number,
    // across BOTH session tables (this is the "normalize now" of what the nightly
    // maintenance job does automatically). Not scoped to the selection.
    auto answer = QMessageBox::question(this, "Repair Legacy Scores",
        "Normalize legacy string-typed scores back to numbers across the whole\n"
        "database (both Sensory and Detailed Sensory sessions)?\n\n"
        "This is lossless — \"7.5\" simply becomes 7.5 — and is the same fix the\n"
        "nightly maintenance job applies automatically. Rows that are already\n"
        "numeric are left untouched. Nothing is deleted.",
        QMessageBox::Yes | QMessageBox::No);
    if (answer != QMessageBox::Yes) return;

    const int n = m_db->normalizeLegacyScores();
    if (n < 0) {
        QMessageBox::warning(this, "Repair Failed",
            "Could not normalize legacy scores:\n" + m_db->lastError());
        return;
    }
    QMessageBox::information(this, "Repair Complete",
        n == 0 ? QStringLiteral("No legacy string-typed scores were found — nothing to repair.")
               : QStringLiteral("Normalized %1 row%2 to numeric scores.")
                     .arg(n).arg(n == 1 ? "" : "s"));
    onRefresh();   // re-classify; the "Legacy string scores" flags clear
}

void DatabaseBrowserDialog::onRepairTpm()
{
    // Backfill the pre-v2.0 "samples exist but no data_rows" gap by re-reading
    // each incomplete file's source .xlsx. runDbRepair self-detects which files
    // are incomplete, so this scans all files rather than the selection.
    auto answer = QMessageBox::question(this, "Repair Incomplete Files",
        "Scan all TPM files for ones saved without their detailed measurement\n"
        "rows (the pre-v2.0 data-row gap) and rebuild them by re-reading the\n"
        "original .xlsx?\n\n"
        "Files that are already complete are skipped. Files whose source .xlsx\n"
        "can no longer be located are reported as skipped (their database row is\n"
        "left unchanged). Nothing is deleted.",
        QMessageBox::Yes | QMessageBox::No);
    if (answer != QMessageBox::Yes) return;

    // Optional: let the user point at a folder to help locate moved .xlsx files.
    // Cancelling the picker leaves sourceDir empty -> locate by stored file_path.
    const QString sourceDir = QFileDialog::getExistingDirectory(
        this, "Optional: folder to search for source .xlsx (Cancel = use stored paths only)");

    // Audit fix (responsiveness): drive a progress dialog from the per-record
    // callback instead of freezing behind a wait cursor. runDbRepair re-reads each
    // incomplete file's .xlsx via a python subprocess, so a large DB could
    // otherwise hang the window ("Not Responding") for many seconds.
    QProgressDialog progress("Repairing incomplete files...", QString(), 0, 0, this);
    progress.setWindowModality(Qt::WindowModal);
    progress.setCancelButton(nullptr);
    progress.setMinimumDuration(0);
    progress.setValue(0);
    QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);

    DataProcessor proc;
    RepairOptions opts;
    opts.sourceDir = sourceDir;   // empty is fine — backfill locates by file_path first
    opts.dryRun    = false;
    opts.progress  = [&progress](int done, int total, const QString& fileName) {
        if (progress.maximum() != total) progress.setMaximum(total);
        progress.setValue(done);
        progress.setLabelText(QStringLiteral("Repairing incomplete files...\n%1").arg(fileName));
        QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
    };
    const RepairSummary summary = runDbRepair(*m_db, proc, opts);
    progress.close();

    QMessageBox::information(this, "Repair Complete",
        QString("Backfill scan finished:\n\n"
                "  Healed (re-read + saved):   %1\n"
                "  Already complete:           %2\n"
                "  Skipped (no source .xlsx):  %3\n"
                "  Failed:                     %4")
            .arg(summary.healed).arg(summary.alreadyComplete)
            .arg(summary.skippedNoXlsx).arg(summary.failed));
    onRefresh();
}

} // namespace DVE

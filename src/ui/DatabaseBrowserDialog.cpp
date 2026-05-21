#include "DatabaseBrowserDialog.h"
#include "SensoryPanel.h"
#include "DetailedSensoryPanel.h"
#include "../utils/AppTheme.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QHeaderView>
#include <QFrame>
#include <QGraphicsDropShadowEffect>
#include <QMessageBox>
#include <QFileDialog>
#include <QFileInfo>
#include <QMap>
#include <QSet>
#include <QDateTime>

namespace DVE {

DatabaseBrowserDialog::DatabaseBrowserDialog(DatabaseManager* db, QWidget* parent)
    : QDialog(parent)
    , m_db(db)
{
    setWindowTitle("Database Browser");
    setMinimumSize(900, 500);
    resize(1000, 600);

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
        auto* refreshBtn = new QPushButton("Refresh", tpmTab);
        filterRow->addWidget(refreshBtn);
        tpmLayout->addLayout(filterRow);

        // File tree
        m_tree = new QTreeWidget(tpmTab);
        m_tree->setColumnCount(5);
        m_tree->setHeaderLabels({"File Name", "Loaded At", "Template", "Tests", "Samples"});
        m_tree->setSelectionBehavior(QAbstractItemView::SelectRows);
        m_tree->setSelectionMode(QAbstractItemView::ExtendedSelection);
        m_tree->setAlternatingRowColors(true);
        m_tree->setRootIsDecorated(false);
        m_tree->setAnimated(true);
        m_tree->setUniformRowHeights(true);
        m_tree->header()->setDefaultAlignment(Qt::AlignLeft | Qt::AlignVCenter);
        m_tree->header()->setSectionResizeMode(QHeaderView::Interactive);
        m_tree->header()->setStretchLastSection(true);
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
        auto* refreshBtn = new QPushButton("Refresh", sensoryTab);
        filterRow->addWidget(refreshBtn);
        sensoryLayout->addLayout(filterRow);

        // Sensory tree (hierarchical: Test Title → Tester entries)
        m_sensoryTree = new QTreeWidget(sensoryTab);
        m_sensoryTree->setColumnCount(5);
        m_sensoryTree->setHeaderLabels({"Test Title / Tester", "Assessor", "Media", "Date", "Samples"});
        m_sensoryTree->setSelectionBehavior(QAbstractItemView::SelectRows);
        m_sensoryTree->setSelectionMode(QAbstractItemView::ExtendedSelection);
        m_sensoryTree->setAlternatingRowColors(true);
        m_sensoryTree->setRootIsDecorated(true);
        m_sensoryTree->setUniformRowHeights(true);
        m_sensoryTree->header()->setDefaultAlignment(Qt::AlignLeft | Qt::AlignVCenter);
        m_sensoryTree->header()->setSectionResizeMode(QHeaderView::Interactive);
        m_sensoryTree->header()->setStretchLastSection(true);
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
        auto* refreshBtn = new QPushButton("Refresh", detSensTab);
        filterRow->addWidget(refreshBtn);
        detSensLayout->addLayout(filterRow);

        // Tree widget
        m_detSensTree = new QTreeWidget(detSensTab);
        m_detSensTree->setColumnCount(5);
        m_detSensTree->setHeaderLabels({"Test Title / Tester", "Assessor", "Media", "Date", "Samples"});
        m_detSensTree->setSelectionBehavior(QAbstractItemView::SelectRows);
        m_detSensTree->setSelectionMode(QAbstractItemView::ExtendedSelection);
        m_detSensTree->setAlternatingRowColors(true);
        m_detSensTree->setRootIsDecorated(true);
        m_detSensTree->setUniformRowHeights(true);
        m_detSensTree->header()->setDefaultAlignment(Qt::AlignLeft | Qt::AlignVCenter);
        m_detSensTree->header()->setSectionResizeMode(QHeaderView::Interactive);
        m_detSensTree->header()->setStretchLastSection(true);
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
    populateTree(m_filterEdit->text());
    populateSensoryTree(m_sensoryFilterEdit->text());
    if (m_detSensFilterEdit) populateDetailedSensoryTree(m_detSensFilterEdit->text());
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

    // Show only the newest row per file_path. listFiles() already returns
    // rows ordered by loaded_at DESC, so the first record we see for each
    // path is the newest — older entries are dropped.
    QSet<QString> seenPaths;
    QVector<const FileRecord*> latest;
    int hiddenDuplicates = 0;

    for (const FileRecord& rec : m_allRecords) {
        if (!filter.isEmpty() &&
            !rec.fileName.contains(filter, Qt::CaseInsensitive) &&
            !rec.filePath.contains(filter, Qt::CaseInsensitive)) {
            continue;
        }
        if (seenPaths.contains(rec.filePath)) {
            ++hiddenDuplicates;
            continue;
        }
        seenPaths.insert(rec.filePath);
        latest.append(&rec);
    }

    auto fmtDate = [](const QString& iso) -> QString {
        QDateTime dt = QDateTime::fromString(iso, Qt::ISODate);
        return dt.isValid() ? dt.toString("yyyy-MM-dd  HH:mm") : iso;
    };

    int totalSamples = 0;
    for (const FileRecord* rec : latest) {
        auto* topItem = new QTreeWidgetItem(m_tree);
        topItem->setText(0, rec->fileName);
        topItem->setText(1, fmtDate(rec->loadedAt));
        topItem->setText(2, rec->templateVersion);
        topItem->setText(3, QString::number(rec->sheetCount));
        topItem->setText(4, QString::number(rec->sampleCount));
        topItem->setData(0, Qt::UserRole, rec->id);
        totalSamples += rec->sampleCount;
    }

    QString status = QString("%1 files  |  %2 total samples")
                         .arg(latest.size()).arg(totalSamples);
    if (hiddenDuplicates > 0) {
        status += QString("  (%1 stale duplicate%2 hidden)")
                      .arg(hiddenDuplicates)
                      .arg(hiddenDuplicates == 1 ? "" : "s");
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

        const QString groupKey   = rec.testTitle.isEmpty() ? rec.sessionName : rec.testTitle;
        const QString testerKey  = rec.testerName.toLower().trimmed();
        if (!groups.contains(groupKey)) {
            groups.insert(groupKey, Group{groupKey, {}, {}});
            groupOrder.append(groupKey);
        }
        Group& g = groups[groupKey];
        if (g.seenTesterKeys.contains(testerKey)) {
            ++hiddenDuplicates;
            continue;
        }
        g.seenTesterKeys.insert(testerKey);
        g.testers.append(&rec);
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

        for (const SensoryRecord* rec : g.testers) {
            auto* child = new QTreeWidgetItem(parentItem);
            QString tester = rec->testerName.isEmpty() ? rec->assessorName : rec->testerName;
            if (tester.isEmpty()) tester = QStringLiteral("(unnamed)");
            child->setText(0, tester);
            child->setText(1, rec->assessorName);
            child->setText(2, rec->media);
            child->setText(3, rec->date);
            child->setText(4, QString::number(rec->sampleCount));
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

    int count = items.size();
    auto answer = QMessageBox::question(this, "Delete Files",
        QString("Delete %1 entry/entries from the database?\n\n"
                "This removes the data from the database only.\n"
                "The original Excel files on disk are not affected.")
            .arg(count),
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

    auto answer = QMessageBox::question(this, "Delete Sensory Session",
        QString("Delete %1 selected sensory session(s) from the database?").arg(ids.size()),
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
    QString path = QFileDialog::getSaveFileName(
        this, "Save Combined Sensory Report", QString(),
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

        const QString groupKey  = rec.testTitle.isEmpty() ? rec.sessionName : rec.testTitle;
        const QString testerKey = rec.testerName.toLower().trimmed();
        if (!groups.contains(groupKey)) {
            groups.insert(groupKey, Group{groupKey, {}, {}});
            groupOrder.append(groupKey);
        }
        Group& g = groups[groupKey];
        if (g.seenTesterKeys.contains(testerKey)) {
            ++hiddenDuplicates;
            continue;
        }
        g.seenTesterKeys.insert(testerKey);
        g.testers.append(&rec);
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

        for (const DetailedSensoryRecord* rec : g.testers) {
            auto* child = new QTreeWidgetItem(parentItem);
            QString tester = rec->testerName.isEmpty() ? rec->assessorName : rec->testerName;
            if (tester.isEmpty()) tester = QStringLiteral("(unnamed)");
            child->setText(0, tester);
            child->setText(1, rec->assessorName);
            child->setText(2, rec->media);
            child->setText(3, rec->date);
            child->setText(4, QString::number(rec->sampleCount));
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

    auto answer = QMessageBox::question(this, "Delete Detailed Sensory Session",
        QString("Delete %1 selected detailed sensory session(s) from the database?").arg(ids.size()),
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

    QString path = QFileDialog::getSaveFileName(
        this, "Save Combined Detailed Sensory Report", QString(),
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

} // namespace DVE

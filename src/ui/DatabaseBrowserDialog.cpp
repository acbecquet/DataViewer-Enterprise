#include "DatabaseBrowserDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QHeaderView>
#include <QMessageBox>
#include <QFileInfo>
#include <QMap>
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

    // ── Filter bar ──────────────────────────────────────────────────────────
    auto* filterRow = new QHBoxLayout;
    filterRow->addWidget(new QLabel("Search:"));
    m_filterEdit = new QLineEdit(this);
    m_filterEdit->setPlaceholderText("Filter by file name...");
    m_filterEdit->setClearButtonEnabled(true);
    filterRow->addWidget(m_filterEdit);
    auto* refreshBtn = new QPushButton("Refresh", this);
    filterRow->addWidget(refreshBtn);
    mainLayout->addLayout(filterRow);

    // ── File tree ───────────────────────────────────────────────────────────
    m_tree = new QTreeWidget(this);
    m_tree->setColumnCount(5);
    m_tree->setHeaderLabels({"File Name", "Loaded At", "Template", "Tests", "Samples"});
    m_tree->setSelectionBehavior(QAbstractItemView::SelectRows);
    m_tree->setSelectionMode(QAbstractItemView::ExtendedSelection);
    m_tree->setAlternatingRowColors(true);
    m_tree->setRootIsDecorated(true);
    m_tree->setAnimated(true);
    m_tree->header()->setSectionResizeMode(0, QHeaderView::Stretch);
    m_tree->header()->setSectionResizeMode(1, QHeaderView::ResizeToContents);
    m_tree->header()->setSectionResizeMode(2, QHeaderView::ResizeToContents);
    m_tree->header()->setSectionResizeMode(3, QHeaderView::ResizeToContents);
    m_tree->header()->setSectionResizeMode(4, QHeaderView::ResizeToContents);
    mainLayout->addWidget(m_tree);

    // ── Status label ────────────────────────────────────────────────────────
    m_statusLabel = new QLabel(this);
    mainLayout->addWidget(m_statusLabel);

    // ── Button row ──────────────────────────────────────────────────────────
    auto* btnRow = new QHBoxLayout;

    m_loadBtn = new QPushButton("Load Selected", this);
    m_loadBtn->setEnabled(false);
    btnRow->addWidget(m_loadBtn);

    m_loadAllBtn = new QPushButton("Load All Visible", this);
    btnRow->addWidget(m_loadAllBtn);

    m_deleteBtn = new QPushButton("Delete Selected", this);
    m_deleteBtn->setEnabled(false);
    m_deleteBtn->setStyleSheet("QPushButton { color: #cc3333; }");
    btnRow->addWidget(m_deleteBtn);

    auto* cleanupBtn = new QPushButton("Cleanup Duplicates", this);
    cleanupBtn->setToolTip("Remove unknown/corrupt entries and keep only 3 most recent per file");
    btnRow->addWidget(cleanupBtn);

    btnRow->addStretch();

    auto* closeBtn = new QPushButton("Close", this);
    btnRow->addWidget(closeBtn);

    mainLayout->addLayout(btnRow);

    // ── Connections ─────────────────────────────────────────────────────────
    connect(refreshBtn,    &QPushButton::clicked, this, &DatabaseBrowserDialog::onRefresh);
    connect(m_filterEdit,  &QLineEdit::textChanged, this, &DatabaseBrowserDialog::onFilterChanged);
    connect(m_loadBtn,     &QPushButton::clicked, this, &DatabaseBrowserDialog::onLoad);
    connect(m_loadAllBtn,  &QPushButton::clicked, this, &DatabaseBrowserDialog::onLoadAll);
    connect(m_deleteBtn,   &QPushButton::clicked, this, &DatabaseBrowserDialog::onDelete);
    connect(cleanupBtn,    &QPushButton::clicked, this, &DatabaseBrowserDialog::onCleanup);
    connect(closeBtn,      &QPushButton::clicked, this, &QDialog::reject);
    connect(m_tree,        &QTreeWidget::itemSelectionChanged, this, &DatabaseBrowserDialog::onSelectionChanged);
    connect(m_tree,        &QTreeWidget::itemDoubleClicked, this, &DatabaseBrowserDialog::onItemDoubleClicked);
    connect(m_tree,        &QTreeWidget::itemExpanded, this, &DatabaseBrowserDialog::onItemExpanded);

    // Initial load
    onRefresh();
}

void DatabaseBrowserDialog::onRefresh()
{
    m_allRecords = m_db->listFiles();
    populateTree(m_filterEdit->text());
}

void DatabaseBrowserDialog::onFilterChanged(const QString& text)
{
    populateTree(text);
}

void DatabaseBrowserDialog::populateTree(const QString& filter)
{
    m_tree->clear();

    // Group records by file_name, most recent first per group
    // Records are already sorted by loaded_at DESC from listFiles()
    QMap<QString, QVector<const FileRecord*>> groups;
    QStringList order;  // preserve insertion order

    for (const FileRecord& rec : m_allRecords) {
        if (!filter.isEmpty() &&
            !rec.fileName.contains(filter, Qt::CaseInsensitive) &&
            !rec.filePath.contains(filter, Qt::CaseInsensitive)) {
            continue;
        }
        if (!groups.contains(rec.fileName))
            order.append(rec.fileName);
        groups[rec.fileName].append(&rec);
    }

    int totalFiles = 0;
    int totalSamples = 0;

    for (const QString& name : order) {
        const auto& recs = groups[name];
        if (recs.isEmpty()) continue;

        const FileRecord* latest = recs[0];  // most recent

        // Format date nicely
        auto fmtDate = [](const QString& iso) -> QString {
            QDateTime dt = QDateTime::fromString(iso, Qt::ISODate);
            return dt.isValid() ? dt.toString("yyyy-MM-dd  HH:mm") : iso;
        };

        // Create top-level item for the most recent version
        auto* topItem = new QTreeWidgetItem(m_tree);
        topItem->setText(0, latest->fileName);
        topItem->setText(1, fmtDate(latest->loadedAt));
        topItem->setText(2, latest->templateVersion);
        topItem->setText(3, QString::number(latest->sheetCount));
        topItem->setText(4, QString::number(latest->sampleCount));
        topItem->setData(0, Qt::UserRole, latest->id);

        // Show version count badge if there are older versions
        if (recs.size() > 1) {
            topItem->setText(0, QString("%1  (%2 versions)")
                .arg(latest->fileName).arg(recs.size()));
            topItem->setChildIndicatorPolicy(QTreeWidgetItem::ShowIndicator);
        }

        // Add older versions as children (expandable)
        for (int i = 1; i < recs.size(); ++i) {
            const FileRecord* older = recs[i];
            auto* childItem = new QTreeWidgetItem(topItem);
            childItem->setText(0, fmtDate(older->loadedAt));
            childItem->setText(1, older->filePath);
            childItem->setText(2, older->templateVersion);
            childItem->setText(3, QString::number(older->sheetCount));
            childItem->setText(4, QString::number(older->sampleCount));
            childItem->setData(0, Qt::UserRole, older->id);
            childItem->setForeground(0, QColor(0x66, 0x66, 0x66));
            childItem->setForeground(1, QColor(0x66, 0x66, 0x66));
        }

        ++totalFiles;
        totalSamples += latest->sampleCount;
    }

    int totalRecords = 0;
    for (const auto& recs : groups)
        totalRecords += recs.size();

    m_statusLabel->setText(QString("%1 unique files  |  %2 total entries  |  %3 total samples")
        .arg(totalFiles).arg(totalRecords).arg(totalSamples));
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

void DatabaseBrowserDialog::onItemDoubleClicked(QTreeWidgetItem* item, int)
{
    if (!item) return;
    int id = idFromItem(item);
    if (id > 0) {
        m_selectedIds.clear();
        m_selectedIds.append(id);
        accept();
    }
}

void DatabaseBrowserDialog::onItemExpanded(QTreeWidgetItem*)
{
    // Auto-resize first column after expanding
    m_tree->resizeColumnToContents(0);
}

void DatabaseBrowserDialog::onLoad()
{
    m_selectedIds.clear();
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
    // Only load top-level items (most recent per group)
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

} // namespace DVE

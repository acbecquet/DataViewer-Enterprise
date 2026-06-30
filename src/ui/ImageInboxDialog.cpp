#include "ImageInboxDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QSplitter>
#include <QGroupBox>
#include <QListWidget>
#include <QListWidgetItem>
#include <QLabel>
#include <QLineEdit>
#include <QComboBox>
#include <QPushButton>
#include <QTabWidget>
#include <QTreeWidget>
#include <QTreeWidgetItem>
#include <QScrollArea>
#include <QFileDialog>
#include <QFileInfo>
#include <QFileInfoList>
#include <QDir>
#include <QPixmap>
#include <QIcon>
#include <QFrame>
#include <QDesktopServices>
#include <QUrl>
#include <QDateTime>
#include <QDialogButtonBox>
#include <QFont>
#include <QToolButton>
#include <QInputDialog>
#include <QGraphicsDropShadowEffect>
#include <QStyle>
#include <algorithm>
#include "../utils/AppTheme.h"

namespace DVE {

// Static empty files vector for sensory mode (m_files is a const reference)
const QVector<FileResult> ImageInboxDialog::s_emptyFiles;

// ─── TPM Constructor ─────────────────────────────────────────────────────────
ImageInboxDialog::ImageInboxDialog(const QVector<FileResult>& files,
                                   int curFileIdx, int curSheetIdx, int curSampleIdx,
                                   const QString& watchFolder,
                                   QWidget* parent)
    : QDialog(parent)
    , m_watchFolder(watchFolder)
    , m_files(files)
    , m_sensoryMode(false)
{
    initUI(-1);

    // ── Pre-select current sample ─────────────────────────────────────────────
    {
        int fi = (curFileIdx >= 0 && curFileIdx < m_fileCombo->count()) ? curFileIdx : 0;
        m_fileCombo->setCurrentIndex(fi);
        populateSheetCombo(fi);
        int si = (curSheetIdx >= 0 && curSheetIdx < m_sheetCombo->count()) ? curSheetIdx : 0;
        if (m_sheetCombo->count() > 0) {
            m_sheetCombo->setCurrentIndex(si);
            populateSampleCombo(fi, si);
            int smi = (curSampleIdx >= 0 && curSampleIdx < m_sampleCombo->count()) ? curSampleIdx : 0;
            if (m_sampleCombo->count() > 0)
                m_sampleCombo->setCurrentIndex(smi);
        }
    }

    loadInboxImages();
    buildLibraryTree();
}

// ─── Sensory Constructor ─────────────────────────────────────────────────────
ImageInboxDialog::ImageInboxDialog(const QStringList& sessionLabels,
                                   int curSessionIdx,
                                   const QString& watchFolder,
                                   QWidget* parent)
    : QDialog(parent)
    , m_watchFolder(watchFolder)
    , m_files(s_emptyFiles)
    , m_sensoryMode(true)
    , m_sessionLabels(sessionLabels)
{
    initUI(curSessionIdx);

    // Pre-select session
    if (m_sessionCombo && curSessionIdx >= 0 && curSessionIdx < m_sessionCombo->count())
        m_sessionCombo->setCurrentIndex(curSessionIdx);

    loadInboxImages();
    buildLibraryTree();
}

// ─── Shared UI setup ─────────────────────────────────────────────────────────
void ImageInboxDialog::initUI(int /*curSessionIdx*/)
{
    setWindowTitle(m_sensoryMode ? "Image Inbox  (Sensory)" : "Image Inbox");
    resize(960, 780);
    setMinimumSize(560, 420);

    auto* shadow = new QGraphicsDropShadowEffect(this);
    shadow->setBlurRadius(16);
    shadow->setOffset(0, 4);
    shadow->setColor(QColor(0, 0, 0, 30));
    setGraphicsEffect(shadow);

    QVBoxLayout* mainVL = new QVBoxLayout(this);
    mainVL->setContentsMargins(16, 16, 16, 16);
    mainVL->setSpacing(6);

    // ── Watch folder bar ──────────────────────────────────────────────────────
    QHBoxLayout* folderHL = new QHBoxLayout();
    folderHL->setSpacing(6);

    QLabel* folderLbl = new QLabel("Watch Folder:", this);
    folderLbl->setMinimumWidth(90);

    m_folderEdit = new QLineEdit(m_watchFolder, this);
    m_folderEdit->setReadOnly(true);

    QPushButton* changeBtn = new QPushButton("Change...", this);
    changeBtn->setMinimumWidth(90);

    QPushButton* openBtn = new QPushButton("Open Folder", this);
    openBtn->setMinimumWidth(100);

    m_inboxCountLabel = new QLabel("", this);

    folderHL->addWidget(folderLbl);
    folderHL->addWidget(m_folderEdit, 1);
    folderHL->addWidget(changeBtn);
    folderHL->addWidget(openBtn);
    folderHL->addWidget(m_inboxCountLabel);

    mainVL->addLayout(folderHL);

    // ── Tab widget ────────────────────────────────────────────────────────────
    m_tabWidget = new QTabWidget(this);

    // ─────────────────────────────────────────────────────────────────────────
    // INBOX TAB
    // ─────────────────────────────────────────────────────────────────────────
    QWidget* inboxTab = new QWidget();
    QVBoxLayout* inboxTabVL = new QVBoxLayout(inboxTab);
    inboxTabVL->setContentsMargins(4, 4, 4, 4);
    inboxTabVL->setSpacing(4);

    QSplitter* inboxSplit = new QSplitter(Qt::Horizontal, inboxTab);

    // Left: icon grid
    m_inboxList = new QListWidget(inboxSplit);
    m_inboxList->setViewMode(QListWidget::IconMode);
    m_inboxList->setIconSize(QSize(96, 96));
    m_inboxList->setGridSize(QSize(116, 126));
    m_inboxList->setResizeMode(QListWidget::Adjust);
    m_inboxList->setWordWrap(true);
    m_inboxList->setSelectionMode(QAbstractItemView::ExtendedSelection);

    // Right: preview + assignment (in a scroll area so nothing overlaps)
    QScrollArea* rightScroll = new QScrollArea(inboxSplit);
    rightScroll->setWidgetResizable(true);
    rightScroll->setFrameShape(QFrame::NoFrame);
    QGroupBox* previewBox = new QGroupBox("Selected Image");
    QVBoxLayout* previewVL = new QVBoxLayout(previewBox);
    previewVL->setSpacing(5);
    rightScroll->setWidget(previewBox);

    m_dateLabel = new QLabel("", previewBox);
    m_dateLabel->setWordWrap(true);
    m_dateLabel->setAlignment(Qt::AlignCenter);
    {
        QFont f = m_dateLabel->font();
        f.setPointSize(8);
        m_dateLabel->setFont(f);
    }
    m_dateLabel->setStyleSheet("padding: 2px 0;");

    m_previewLabel = new QLabel(previewBox);
    m_previewLabel->setFixedSize(160, 160);
    m_previewLabel->setScaledContents(false);
    m_previewLabel->setAlignment(Qt::AlignCenter);
    m_previewLabel->setStyleSheet("background: #F0F0F0;");

    m_fileInfoLabel = new QLabel("", previewBox);
    m_fileInfoLabel->setWordWrap(true);
    QFont infoFont = m_fileInfoLabel->font();
    infoFont.setPointSize(8);
    m_fileInfoLabel->setFont(infoFont);
    m_fileInfoLabel->setAlignment(Qt::AlignTop | Qt::AlignLeft);
    // color handled by global QSS

    // Rename row
    QHBoxLayout* renameHL = new QHBoxLayout();
    renameHL->setSpacing(4);
    m_renameEdit = new QLineEdit(previewBox);
    m_renameEdit->setPlaceholderText("Filename (no extension)");
    m_renameEdit->setEnabled(false);
    m_renameBtn = new QToolButton(previewBox);
    m_renameBtn->setText("Rename");
    m_renameBtn->setEnabled(false);
    // button styled by global QSS
    renameHL->addWidget(m_renameEdit, 1);
    renameHL->addWidget(m_renameBtn);

    QFrame* sep = new QFrame(previewBox);
    sep->setFrameShape(QFrame::HLine);

    QLabel* assignLbl = new QLabel("Assign to:", previewBox);
    assignLbl->setStyleSheet("font-weight: 600;");

    previewVL->addWidget(m_dateLabel);
    previewVL->addWidget(m_previewLabel, 0, Qt::AlignHCenter);
    previewVL->addWidget(m_fileInfoLabel);
    previewVL->addLayout(renameHL);
    previewVL->addWidget(sep);
    previewVL->addWidget(assignLbl);

    if (m_sensoryMode) {
        // ── Sensory mode: single session combo ────────────────────────────────
        QLabel* sessionLbl = new QLabel("Session:", previewBox);
        m_sessionCombo = new QComboBox(previewBox);
        for (const QString& lbl : m_sessionLabels)
            m_sessionCombo->addItem(lbl);

        m_addBtn = new QPushButton("Add to Session", previewBox);
        m_addBtn->setProperty("primary", true);
        m_addBtn->style()->unpolish(m_addBtn);
        m_addBtn->style()->polish(m_addBtn);

        previewVL->addWidget(sessionLbl);
        previewVL->addWidget(m_sessionCombo);
        previewVL->addWidget(m_addBtn);
    } else {
        // ── TPM mode: file / sheet / sample combos ────────────────────────────
        QLabel* fileLbl = new QLabel("File:", previewBox);
        m_fileCombo = new QComboBox(previewBox);
        for (const FileResult& fr : m_files)
            m_fileCombo->addItem(fr.fileName);

        QLabel* sheetLbl = new QLabel("Sheet:", previewBox);
        m_sheetCombo = new QComboBox(previewBox);

        QLabel* sampleLbl = new QLabel("Sample:", previewBox);
        m_sampleCombo = new QComboBox(previewBox);

        m_addBtn = new QPushButton("Add to Sample", previewBox);
        m_addBtn->setProperty("primary", true);
        m_addBtn->style()->unpolish(m_addBtn);
        m_addBtn->style()->polish(m_addBtn);

        previewVL->addWidget(fileLbl);
        previewVL->addWidget(m_fileCombo);
        previewVL->addWidget(sheetLbl);
        previewVL->addWidget(m_sheetCombo);
        previewVL->addWidget(sampleLbl);
        previewVL->addWidget(m_sampleCombo);
        previewVL->addWidget(m_addBtn);
    }

    previewVL->addStretch(1);

    inboxSplit->addWidget(m_inboxList);
    inboxSplit->addWidget(rightScroll);
    inboxSplit->setStretchFactor(0, 3);
    inboxSplit->setStretchFactor(1, 2);

    inboxTabVL->addWidget(inboxSplit, 1);

    // Bottom bar for inbox tab
    QHBoxLayout* inboxBottomHL = new QHBoxLayout();
    QPushButton* selectAllBtn = new QPushButton("Select All", inboxTab);
    selectAllBtn->setMinimumWidth(80);
    QPushButton* refreshBtn = new QPushButton("Refresh", inboxTab);
    refreshBtn->setMinimumWidth(70);
    m_loadMoreBtn = new QPushButton("Load 20 More", inboxTab);
    m_loadMoreBtn->setMinimumWidth(100);
    m_loadMoreBtn->setVisible(false);
    QLabel* imgCountLbl = new QLabel("", inboxTab);
    imgCountLbl->setObjectName("imgCountLbl");

    inboxBottomHL->addWidget(selectAllBtn);
    inboxBottomHL->addWidget(refreshBtn);
    inboxBottomHL->addWidget(m_loadMoreBtn);
    inboxBottomHL->addStretch(1);
    inboxBottomHL->addWidget(imgCountLbl);
    inboxTabVL->addLayout(inboxBottomHL);

    m_tabWidget->addTab(inboxTab, "Inbox");

    // ─────────────────────────────────────────────────────────────────────────
    // LIBRARY TAB
    // ─────────────────────────────────────────────────────────────────────────
    QWidget* libraryTab = new QWidget();
    QVBoxLayout* libTabVL = new QVBoxLayout(libraryTab);
    libTabVL->setContentsMargins(4, 4, 4, 4);
    libTabVL->setSpacing(4);

    QSplitter* libSplit = new QSplitter(Qt::Horizontal, libraryTab);

    m_libraryTree = new QTreeWidget(libSplit);
    m_libraryTree->setHeaderHidden(true);

    m_libraryImageArea = new QScrollArea(libSplit);
    m_libraryImageArea->setWidgetResizable(true);

    m_libraryImagePanel = new QWidget();
    m_libraryImagePanel->setLayout(new QGridLayout());
    m_libraryImageArea->setWidget(m_libraryImagePanel);

    libSplit->addWidget(m_libraryTree);
    libSplit->addWidget(m_libraryImageArea);
    libSplit->setStretchFactor(0, 2);
    libSplit->setStretchFactor(1, 3);
    libSplit->setSizes({384, 576}); // ~40% / 60%

    libTabVL->addWidget(libSplit, 1);

    m_tabWidget->addTab(libraryTab, "Library");

    mainVL->addWidget(m_tabWidget, 1);

    // ── Dialog buttons ────────────────────────────────────────────────────────
    QDialogButtonBox* bbox = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    mainVL->addWidget(bbox);

    // ── Connect signals ───────────────────────────────────────────────────────
    connect(changeBtn,    &QPushButton::clicked, this, &ImageInboxDialog::onBrowseFolder);
    connect(openBtn,      &QPushButton::clicked, this, &ImageInboxDialog::onOpenFolder);
    connect(refreshBtn,   &QPushButton::clicked, this, &ImageInboxDialog::onRefreshInbox);
    connect(m_loadMoreBtn,&QPushButton::clicked, this, &ImageInboxDialog::onLoadMoreImages);
    connect(selectAllBtn, &QPushButton::clicked, this, &ImageInboxDialog::onSelectAll);
    connect(m_addBtn,     &QPushButton::clicked, this, &ImageInboxDialog::onAddToSample);
    connect(m_renameBtn,  &QToolButton::clicked, this, &ImageInboxDialog::onRenameImage);
    connect(m_inboxList,  &QListWidget::itemSelectionChanged,
            this, &ImageInboxDialog::onInboxItemSelectionChanged);

    if (!m_sensoryMode) {
        connect(m_fileCombo,   QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, &ImageInboxDialog::onFileComboChanged);
        connect(m_sheetCombo,  QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, &ImageInboxDialog::onSheetComboChanged);
        connect(m_sampleCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, &ImageInboxDialog::onSampleComboChanged);
    }

    connect(m_libraryTree, &QTreeWidget::itemSelectionChanged,
            this, &ImageInboxDialog::onLibraryItemSelected);
    connect(bbox, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(bbox, &QDialogButtonBox::rejected, this, &QDialog::reject);
}

// ─── Inbox loading ────────────────────────────────────────────────────────────
void ImageInboxDialog::loadInboxImages()
{
    m_inboxList->clear();
    m_allInboxFiles.clear();
    m_inboxOffset = 0;
    if (m_loadMoreBtn) m_loadMoreBtn->setVisible(false);

    if (m_watchFolder.isEmpty() || !QDir(m_watchFolder).exists()) {
        updateInboxCount();
        return;
    }

    QDir dir(m_watchFolder);
    QStringList filters = {"*.jpg", "*.jpeg", "*.png", "*.bmp", "*.webp", "*.gif"};
    m_allInboxFiles = dir.entryInfoList(filters, QDir::Files);

    // Sort by modification time, newest first
    std::sort(m_allInboxFiles.begin(), m_allInboxFiles.end(),
              [](const QFileInfo& a, const QFileInfo& b) {
                  return a.lastModified() > b.lastModified();
              });

    appendInboxPage();
    updateInboxCount();
}

void ImageInboxDialog::appendInboxPage()
{
    int end = qMin(m_inboxOffset + kPageSize, (int)m_allInboxFiles.size());
    for (int i = m_inboxOffset; i < end; ++i) {
        const QFileInfo& fi = m_allInboxFiles[i];
        QPixmap px(fi.absoluteFilePath());
        QIcon icon;
        if (!px.isNull())
            icon = QIcon(px.scaled(96, 96, Qt::KeepAspectRatio, Qt::SmoothTransformation));

        QListWidgetItem* item = new QListWidgetItem(icon, fi.fileName());
        item->setData(Qt::UserRole, fi.absoluteFilePath());
        item->setToolTip(fi.absoluteFilePath());
        m_inboxList->addItem(item);
    }
    m_inboxOffset = end;
    if (m_loadMoreBtn)
        m_loadMoreBtn->setVisible(m_inboxOffset < (int)m_allInboxFiles.size());
}

// ─── Preview ──────────────────────────────────────────────────────────────────
void ImageInboxDialog::showPreview(const QString& path)
{
    m_selectedImagePath = path;

    if (path.isEmpty()) {
        if (m_dateLabel)     m_dateLabel->clear();
        m_previewLabel->clear();
        m_fileInfoLabel->clear();
        return;
    }

    QFileInfo fi(path);
    if (m_dateLabel)
        m_dateLabel->setText(fi.lastModified().toString("yyyy-MM-dd  hh:mm"));

    QPixmap px(path);
    if (!px.isNull()) {
        m_previewLabel->setPixmap(
            px.scaled(160, 160, Qt::KeepAspectRatio, Qt::SmoothTransformation));
    } else {
        m_previewLabel->clear();
    }

    m_fileInfoLabel->setText(QString("<b>%1</b>").arg(fi.fileName()));
}

void ImageInboxDialog::onInboxItemSelectionChanged()
{
    QList<QListWidgetItem*> sel = m_inboxList->selectedItems();
    if (sel.isEmpty()) {
        showPreview({});
        if (m_renameEdit) { m_renameEdit->clear(); m_renameEdit->setEnabled(false); }
        if (m_renameBtn)  { m_renameBtn->setEnabled(false); }
        return;
    }
    QString path = sel.last()->data(Qt::UserRole).toString();
    showPreview(path);
    if (m_renameEdit && m_renameBtn) {
        bool single = (sel.size() == 1);
        m_renameEdit->setEnabled(single);
        m_renameBtn->setEnabled(single);
        if (single)
            updateRenameDefault();
        else
            m_renameEdit->clear();
    }
}

// ─── Add to Sample / Session ──────────────────────────────────────────────────
void ImageInboxDialog::onAddToSample()
{
    QList<QListWidgetItem*> sel = m_inboxList->selectedItems();
    if (sel.isEmpty()) {
        m_fileInfoLabel->setText("No images selected.");
        return;
    }

    QStringList paths;
    for (QListWidgetItem* item : sel)
        paths << item->data(Qt::UserRole).toString();

    if (m_sensoryMode) {
        // ── Sensory mode: assign to session ──────────────────────────────────
        int sessIdx = m_sessionCombo ? m_sessionCombo->currentIndex() : -1;
        if (sessIdx < 0) {
            m_fileInfoLabel->setText("No session selected.");
            return;
        }

        bool found = false;
        for (Assignment& a : m_assignments) {
            if (a.fileIdx == -1 && a.sheetIdx == -1 && a.sampleIdx == sessIdx) {
                for (const QString& p : paths)
                    if (!a.imagePaths.contains(p))
                        a.imagePaths << p;
                found = true;
                break;
            }
        }
        if (!found) {
            Assignment a;
            a.fileIdx   = -1;
            a.sheetIdx  = -1;
            a.sampleIdx = sessIdx;
            a.imagePaths = paths;
            m_assignments.append(a);
        }

        QString sessName = m_sessionCombo->currentText();
        if (sessName.isEmpty()) sessName = QString("session %1").arg(sessIdx + 1);
        m_fileInfoLabel->setText(
            QString("<span style='color:green;'>&#10003; Added %1 image(s) to <b>%2</b></span>")
                .arg(paths.size())
                .arg(sessName));

    } else {
        // ── TPM mode: assign to file/sheet/sample ────────────────────────────
        int fIdx = m_fileCombo->currentIndex();
        int sIdx = m_sheetCombo->currentIndex();
        int smIdx = m_sampleCombo->currentIndex();

        bool found = false;
        for (Assignment& a : m_assignments) {
            if (a.fileIdx == fIdx && a.sheetIdx == sIdx && a.sampleIdx == smIdx) {
                for (const QString& p : paths)
                    if (!a.imagePaths.contains(p))
                        a.imagePaths << p;
                found = true;
                break;
            }
        }
        if (!found) {
            Assignment a;
            a.fileIdx   = fIdx;
            a.sheetIdx  = sIdx;
            a.sampleIdx = smIdx;
            a.imagePaths = paths;
            m_assignments.append(a);
        }

        QString sampleName = m_sampleCombo->currentText();
        if (sampleName.isEmpty()) sampleName = QString("sample %1").arg(smIdx + 1);
        m_fileInfoLabel->setText(
            QString("<span style='color:green;'>&#10003; Added %1 image(s) to <b>%2</b></span>")
                .arg(paths.size())
                .arg(sampleName));
    }
}

// ─── Rename ──────────────────────────────────────────────────────────────────
void ImageInboxDialog::onRenameImage()
{
    QList<QListWidgetItem*> sel = m_inboxList->selectedItems();
    if (sel.size() != 1 || !m_renameEdit) return;

    QListWidgetItem* item = sel.first();
    QString oldPath = item->data(Qt::UserRole).toString();
    QString newStem = m_renameEdit->text().trimmed();
    if (newStem.isEmpty()) return;

    QFileInfo fi(oldPath);
    QString newPath = fi.absoluteDir().filePath(newStem + "." + fi.suffix());

    if (newPath == oldPath) return; // no change

    if (QFile::exists(newPath)) {
        m_fileInfoLabel->setText(
            "<span style='color:red;'>A file with that name already exists.</span>");
        return;
    }

    if (!QFile::rename(oldPath, newPath)) {
        m_fileInfoLabel->setText(
            "<span style='color:red;'>Rename failed. Check permissions.</span>");
        return;
    }

    // Update the list item in-place
    item->setText(newStem + "." + fi.suffix());
    item->setData(Qt::UserRole, newPath);
    item->setToolTip(newPath);

    // Reload preview with new path
    showPreview(newPath);
    m_fileInfoLabel->setText(
        QString("<span style='color:green;'>&#10003; Renamed to <b>%1</b></span>")
            .arg(newStem + "." + fi.suffix()));
}

// ─── Select All ──────────────────────────────────────────────────────────────
void ImageInboxDialog::onSelectAll()
{
    m_inboxList->selectAll();
}

// ─── Load more ───────────────────────────────────────────────────────────────
void ImageInboxDialog::onLoadMoreImages()
{
    appendInboxPage();
    updateInboxCount();
}

// ─── Refresh ─────────────────────────────────────────────────────────────────
void ImageInboxDialog::onRefreshInbox()
{
    loadInboxImages();
}

// ─── Browse folder ───────────────────────────────────────────────────────────
void ImageInboxDialog::onBrowseFolder()
{
    QString dir = QFileDialog::getExistingDirectory(
        this, "Select Watch Folder", m_watchFolder);
    if (dir.isEmpty()) return;

    m_watchFolder = dir;
    m_folderEdit->setText(m_watchFolder);
    loadInboxImages();
    emit watchFolderChanged(m_watchFolder);
}

// ─── Open folder in explorer ─────────────────────────────────────────────────
void ImageInboxDialog::onOpenFolder()
{
    if (!m_watchFolder.isEmpty())
        QDesktopServices::openUrl(QUrl::fromLocalFile(m_watchFolder));
}

// ─── Combo population ────────────────────────────────────────────────────────
void ImageInboxDialog::onFileComboChanged(int idx)
{
    populateSheetCombo(idx);
}

void ImageInboxDialog::onSheetComboChanged(int idx)
{
    int fIdx = m_fileCombo->currentIndex();
    populateSampleCombo(fIdx, idx);
}

void ImageInboxDialog::onSampleComboChanged(int /*idx*/)
{
    updateRenameDefault();
}

void ImageInboxDialog::updateRenameDefault()
{
    if (!m_renameEdit || m_selectedImagePath.isEmpty()) return;
    if (m_inboxList->selectedItems().size() != 1) return;

    QString stem = QFileInfo(m_selectedImagePath).completeBaseName();

    if (m_sensoryMode) {
        // In sensory mode, propose: sessionLabel-originalName
        QString sessLabel;
        if (m_sessionCombo && m_sessionCombo->currentIndex() >= 0)
            sessLabel = m_sessionCombo->currentText();
        QString proposed = (!sessLabel.isEmpty())
            ? (sessLabel + "-" + stem)
            : stem;
        m_renameEdit->setText(proposed);
        return;
    }

    // TPM mode
    int fIdx  = m_fileCombo->currentIndex();
    int sIdx  = m_sheetCombo->currentIndex();
    int smIdx = m_sampleCombo->currentIndex();

    QString sampleID;
    QString sheetName;
    if (fIdx >= 0 && fIdx < m_files.size()) {
        const FileResult& fr = m_files[fIdx];
        if (sIdx >= 0 && sIdx < fr.sheets.size()) {
            sheetName = fr.sheets[sIdx].sheetName;
            if (smIdx >= 0 && smIdx < fr.sheets[sIdx].samples.size()) {
                const SampleResult& sr = fr.sheets[sIdx].samples[smIdx];
                sampleID = sr.sampleID.isEmpty() ? sr.sampleName : sr.sampleID;
            }
        }
    }

    QString proposed = (!sampleID.isEmpty() && !sheetName.isEmpty())
        ? (sampleID + "-" + sheetName + "-" + stem)
        : stem;
    m_renameEdit->setText(proposed);
}

void ImageInboxDialog::populateSheetCombo(int fileIdx)
{
    if (!m_sheetCombo) return;
    m_sheetCombo->clear();
    if (fileIdx < 0 || fileIdx >= m_files.size()) return;
    for (const SheetResult& s : m_files[fileIdx].sheets)
        m_sheetCombo->addItem(s.sheetName);
}

void ImageInboxDialog::populateSampleCombo(int fileIdx, int sheetIdx)
{
    if (!m_sampleCombo) return;
    m_sampleCombo->clear();
    if (fileIdx < 0 || fileIdx >= m_files.size()) return;
    const FileResult& fr = m_files[fileIdx];
    if (sheetIdx < 0 || sheetIdx >= fr.sheets.size()) return;
    for (const SampleResult& s : fr.sheets[sheetIdx].samples)
        m_sampleCombo->addItem(s.sampleName.isEmpty() ? s.sampleID : s.sampleName);
}

// ─── Library tab ─────────────────────────────────────────────────────────────
void ImageInboxDialog::buildLibraryTree()
{
    m_libraryTree->clear();
    if (m_sensoryMode) return;   // no library data in sensory mode
    for (int fi = 0; fi < m_files.size(); ++fi) {
        const FileResult& fr = m_files[fi];
        QTreeWidgetItem* fileItem = nullptr;

        for (int shi = 0; shi < fr.sheets.size(); ++shi) {
            const SheetResult& sheet = fr.sheets[shi];
            for (int smi = 0; smi < sheet.samples.size(); ++smi) {
                const SampleResult& s = sheet.samples[smi];
                if (s.imagePaths.isEmpty()) continue;

                if (!fileItem) {
                    fileItem = new QTreeWidgetItem(m_libraryTree);
                    fileItem->setText(0, fr.fileName);
                    fileItem->setExpanded(true);
                }

                QString sampleLabel = s.sampleName.isEmpty() ? s.sampleID : s.sampleName;
                QTreeWidgetItem* sampleItem = new QTreeWidgetItem(fileItem);
                sampleItem->setText(0,
                    QString("%1 (%2 image%3)")
                        .arg(sampleLabel)
                        .arg(s.imagePaths.size())
                        .arg(s.imagePaths.size() == 1 ? "" : "s"));
                // Store indices for lookup
                sampleItem->setData(0, Qt::UserRole,     fi);
                sampleItem->setData(0, Qt::UserRole + 1, shi);
                sampleItem->setData(0, Qt::UserRole + 2, smi);
            }
        }
    }
}

void ImageInboxDialog::onLibraryItemSelected()
{
    QList<QTreeWidgetItem*> sel = m_libraryTree->selectedItems();
    if (sel.isEmpty()) return;
    QTreeWidgetItem* item = sel.first();

    // Only leaf items (samples) carry UserRole data
    QVariant v = item->data(0, Qt::UserRole);
    if (!v.isValid()) return;

    int fi  = item->data(0, Qt::UserRole).toInt();
    int shi = item->data(0, Qt::UserRole + 1).toInt();
    int smi = item->data(0, Qt::UserRole + 2).toInt();

    if (fi  < 0 || fi  >= m_files.size()) return;
    if (shi < 0 || shi >= m_files[fi].sheets.size()) return;
    if (smi < 0 || smi >= m_files[fi].sheets[shi].samples.size()) return;

    const SampleResult& s = m_files[fi].sheets[shi].samples[smi];

    // Rebuild library panel grid
    // Delete old layout & children
    if (QLayout* oldLayout = m_libraryImagePanel->layout()) {
        QLayoutItem* child;
        while ((child = oldLayout->takeAt(0)) != nullptr) {
            if (child->widget()) child->widget()->deleteLater();
            delete child;
        }
        delete oldLayout;
    }

    QGridLayout* grid = new QGridLayout(m_libraryImagePanel);
    grid->setSpacing(6);
    grid->setContentsMargins(6, 6, 6, 6);

    const int cols = 3;
    for (int i = 0; i < s.imagePaths.size(); ++i) {
        QPixmap px(s.imagePaths[i]);
        QLabel* thumb = new QLabel(m_libraryImagePanel);
        thumb->setFixedSize(96, 96);
        thumb->setAlignment(Qt::AlignCenter);
        thumb->setStyleSheet("background: #F0F0F0;");
        if (!px.isNull())
            thumb->setPixmap(px.scaled(96, 96, Qt::KeepAspectRatio, Qt::SmoothTransformation));
        thumb->setToolTip(QFileInfo(s.imagePaths[i]).fileName());
        grid->addWidget(thumb, i / cols, i % cols);
    }
    // Push thumbnails to top-left
    grid->setRowStretch(s.imagePaths.size() / cols + 1, 1);
    grid->setColumnStretch(cols, 1);
}

// ─── Inbox count ─────────────────────────────────────────────────────────────
void ImageInboxDialog::updateInboxCount()
{
    int shown = m_inboxList->count();
    int total = m_allInboxFiles.size();
    QString text = (total > shown)
        ? QString("Showing %1 of %2 images").arg(shown).arg(total)
        : QString("%1 image%2 in inbox").arg(shown).arg(shown == 1 ? "" : "s");

    m_inboxCountLabel->setText(text);

    // Also update the bottom label inside the inbox tab if it exists
    if (m_tabWidget) {
        QWidget* inboxTab = m_tabWidget->widget(0);
        if (inboxTab) {
            QLabel* lbl = inboxTab->findChild<QLabel*>("imgCountLbl");
            if (lbl) lbl->setText(text);
        }
    }
}

} // namespace DVE

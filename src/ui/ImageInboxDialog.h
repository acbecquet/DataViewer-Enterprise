#pragma once
#include <QDialog>
#include <QMap>
#include <QStringList>
#include <QVector>
#include <QList>
#include <QFileInfo>
#include "../pipeline/ReportData.h"

class QListWidget;
class QListWidgetItem;
class QLabel;
class QLineEdit;
class QComboBox;
class QPushButton;
class QTabWidget;
class QTreeWidget;
class QScrollArea;
class QToolButton;

namespace DVE {

class ImageInboxDialog : public QDialog {
    Q_OBJECT
public:
    // files         — currently loaded files for the sample assignment combos
    // curFileIdx/sheetIdx/sampleIdx — pre-selects target sample (can be -1)
    // watchFolder   — initial watch folder path
    // TPM mode: assign images to file/sheet/sample targets
    explicit ImageInboxDialog(const QVector<FileResult>& files,
                              int curFileIdx, int curSheetIdx, int curSampleIdx,
                              const QString& watchFolder,
                              QWidget* parent = nullptr);

    // Sensory mode: assign images to sensory session targets
    explicit ImageInboxDialog(const QStringList& sessionLabels,
                              int curSessionIdx,
                              const QString& watchFolder,
                              QWidget* parent = nullptr);

    // Changed watch folder (empty = unchanged)
    QString watchFolder() const { return m_watchFolder; }

    // Images the user chose to import, keyed by (fileIdx, sheetIdx, sampleIdx)
    // Multiple calls to "Add to Sample" accumulate here.
    struct Assignment {
        int fileIdx   = -1;
        int sheetIdx  = -1;
        int sampleIdx = -1;
        QStringList imagePaths;
    };
    QVector<Assignment> assignments() const { return m_assignments; }

signals:
    void watchFolderChanged(const QString& newPath);

private slots:
    void onBrowseFolder();
    void onOpenFolder();
    void onRefreshInbox();
    void onLoadMoreImages();
    void onInboxItemSelectionChanged();
    void onAddToSample();
    void onSelectAll();
    void onRenameImage();
    void onFileComboChanged(int idx);
    void onSheetComboChanged(int idx);
    void onSampleComboChanged(int idx);
    void onLibraryItemSelected();

private:
    void initUI(int curSessionIdx);        // shared UI setup
    void loadInboxImages();
    void appendInboxPage();
    void buildLibraryTree();
    void showPreview(const QString& path);
    void populateSheetCombo(int fileIdx);
    void populateSampleCombo(int fileIdx, int sheetIdx);
    void updateInboxCount();
    void updateRenameDefault();

    QString m_watchFolder;
    const QVector<FileResult>& m_files;
    QVector<Assignment> m_assignments;

    // Sensory mode
    bool        m_sensoryMode = false;
    QStringList m_sessionLabels;
    QComboBox*  m_sessionCombo = nullptr;
    static const QVector<FileResult> s_emptyFiles;

    // Inbox tab
    QListWidget*  m_inboxList        = nullptr;
    QLabel*       m_dateLabel        = nullptr;
    QLabel*       m_previewLabel     = nullptr;
    QLabel*       m_fileInfoLabel    = nullptr;
    QString       m_selectedImagePath;
    QComboBox*    m_fileCombo        = nullptr;
    QComboBox*    m_sheetCombo       = nullptr;
    QComboBox*    m_sampleCombo      = nullptr;
    QPushButton*  m_addBtn           = nullptr;
    QPushButton*  m_loadMoreBtn      = nullptr;
    QLabel*       m_inboxCountLabel  = nullptr;
    QLineEdit*    m_renameEdit       = nullptr;
    QToolButton*  m_renameBtn        = nullptr;

    // Pagination state
    QList<QFileInfo> m_allInboxFiles;
    int              m_inboxOffset   = 0;
    static constexpr int kPageSize   = 20;

    // Library tab
    QTreeWidget*  m_libraryTree      = nullptr;
    QScrollArea*  m_libraryImageArea = nullptr;
    QWidget*      m_libraryImagePanel = nullptr;

    // Watch folder bar
    QLineEdit*    m_folderEdit       = nullptr;

    QTabWidget*   m_tabWidget        = nullptr;
};

} // namespace DVE

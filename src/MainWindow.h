#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QComboBox>
#include <QTableWidget>
#include <QPushButton>
#include <QLabel>
#include <QVBoxLayout>
#include <QMenuBar>
#include <QMenu>
#include <QAction>
#include <QStatusBar>
#include <QString>
#include <QMap>
#include <QDebug>
#include <ExcelReader.h>

class MainWindow : public QMainWindow
{
	Q_OBJECT

public:
	explicit MainWindow(QWidget* parent = nullptr);
	~MainWindow();

private slots:
	// File operations
	void onNewFile();
	void onLoadFile();
	void onSaveFile();
	void onExit();

	// Data Operations
	void onFileSelected(int index);
	void onSheetSelected(int index);

	// Sample navigation
	void onPrevSample();
	void onNextSample();

	// Report generation
	void onGenerateTestReport();
	void onGenerateFullReport();

	// Help menu
	void onHelp();
	void onAbout();

private:
	// UI Setup
	void setupUI();
	void createMenuBar();
	void createTopFrame();
    void createCenterFrame();
	void createBottomFrame();

	// UI components - Top Frame
	QWidget *topFrame;
	QComboBox *fileDropdown;
	QComboBox *sheetDropdown;
	QPushButton *loadButton;
	QPushButton *saveButton;

	// UI Components - Center Frame
	QWidget *centerFrame;
	QWidget* leftPanel; // table and stats
	QWidget* rightPanel; // plot
	QTableWidget* dataTable;
    QWidget* statsFrame;
    QLabel* statsLabel;

	// Sample navigation
	QWidget* sampleNavFrame;
	QPushButton* prevSampleButton;
	QPushButton* nextSampleButton;
	QLabel* sampleCountLabel;

	// Plot components
	QWidget* plotFrame;
    QLabel* plotLabel; // placeholder for plotting

	// UI Components - Bottom Frame
	QWidget *bottomFrame;
    QLabel *imageLabel;
	QPushButton *loadImagesButton;

	// Menu Actions
	QAction *newAction;
	QAction *loadAction;
	QAction *saveAction;
	QAction *exitAction;
	QAction *generateTestReportAction;
	QAction *generateFullReportAction;
	QAction *helpAction;
	QAction *aboutAction;

	// Data members
	QString currentFile;
	QString currentSheet;
	QMap <QString, QStringList> fileSheets; // Maps filename to list of sheets

	// Excel Data Management
	ExcelReader* m_excelReader;
	QVector<ExcelReader::SampleData> m_currentSamples;
	int m_currentSampleIndex;

	// Helper functions
	void debugPrint(const QString& message);
	void updateFileDropdown();
	void updateSheetDropdown();

	// Excel Operations
	void loadExcelData();
	void displaySample(int sampleIndex);
	void populateTableWithSample(const ExcelReader::SampleData& sample);
	void updateSampleNavigation();
	void updateSampleStatistics(const ExcelReader::SampleData& sample);
};

#endif // MAINWINDOW_H

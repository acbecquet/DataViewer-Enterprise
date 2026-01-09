#include "MainWindow.h"
#include <QApplication>
#include <QMessageBox>
#include <QFileDialog>
#include <QHeaderView>

MainWindow::MainWindow(QWidget *parent)
	: QMainWindow(parent)
{
	debugPrint("MainWindow constructor starting...");

	setWindowTitle("DataViewer Enterprise v1.0");
	resize(1200, 800);

	setupUI();

	debugPrint("MainWindow constructor comple");
}

MainWindow::~MainWindow()
{
	debugPrint("MainWindow destructor called");
}

void MainWindow::setupUI()
{
	debugPrint("Setting up UI components...");

	// Create central widget with main layout
	QWidget* centralWidget = new QWidget(this);
	QVBoxLayout* mainLayout = new QVBoxLayout(centralWidget);
	mainLayout->setContentsMargins(10, 10, 10, 10);
	mainLayout->setSpacing(10);
	
	// Create menu bar
	createMenuBar();

	// Create top frame (file/sheet selection)
	createTopFrame();
	mainLayout->addWidget(topFrame);

	// Create center frame (data table)
	createCenterFrame();
    mainLayout->addWidget(centerFrame, 1); // Stretch factor 1

	// Create bottom frame (images)
	createBottomFrame();
	mainLayout->addWidget(bottomFrame);

	setCentralWidget(centralWidget);

	// Status bar
	statusBar()->showMessage("Ready");

	debugPrint("UI setup complete");
}

void MainWindow::createMenuBar()
{
	debugPrint("Creating menu bar...");
	QMenuBar* menuBar = new QMenuBar(this);
	setMenuBar(menuBar);

	// File Menu
	QMenu* fileMenu = menuBar->addMenu("&File");

	newAction = new QAction("&New", this);
	newAction->setShortcut(QKeySequence::New);
	connect(newAction, &QAction::triggered, this, &MainWindow::onNewFile);
	fileMenu->addAction(newAction);

	loadAction = new QAction("&Load", this);
	loadAction->setShortcut(QKeySequence::Open);
	connect(loadAction, &QAction::triggered, this, &MainWindow::onLoadFile);
	fileMenu->addAction(loadAction);

	saveAction = new QAction("&Save", this);
	saveAction->setShortcut(QKeySequence::Save);
	connect(saveAction, &QAction::triggered, this, &MainWindow::onSaveFile);
	fileMenu->addAction(saveAction);

	fileMenu->addSeparator();

	exitAction = new QAction("&Exit", this);
	exitAction->setShortcut(QKeySequence::Quit);
	connect(exitAction, &QAction::triggered, this, &MainWindow::onExit);
	fileMenu->addAction(exitAction);

	// Reports Menu
    QMenu* reportsMenu = menuBar->addMenu("&Reports");

    generateTestReportAction = new QAction("Generate &Test Report", this);
	connect(generateTestReportAction, &QAction::triggered, this, &MainWindow::onGenerateTestReport);
	reportsMenu->addAction(generateTestReportAction);

	generateFullReportAction = new QAction("Generate &Full Report", this);
	connect(generateFullReportAction, &QAction::triggered, this, &MainWindow::onGenerateFullReport);
    reportsMenu->addAction(generateFullReportAction);

	// Help Menu
	QMenu* helpMenu = menuBar->addMenu("&Help");

	helpAction = new QAction("&Help", this);
	helpAction->setShortcut(QKeySequence::HelpContents);
	connect(helpAction, &QAction::triggered, this, &MainWindow::onHelp);
	helpMenu->addAction(helpAction);

	aboutAction = new QAction("&About", this);
    connect(aboutAction, &QAction::triggered, this, &MainWindow::onAbout);
	helpMenu->addAction(aboutAction);

	debugPrint("Menu bar created with File, Reports, and Help menus");
}

void MainWindow::createTopFrame()
{
	debugPrint("Creating top frame...");

	topFrame = new QWidget(this);
	QHBoxLayout* layout = new QHBoxLayout(topFrame);
	layout->setContentsMargins(0, 0, 0, 0);

	// File dropdown
	QLabel* fileLabel = new QLabel("File:", topFrame);
	layout->addWidget(fileLabel);

	fileDropdown = new QComboBox(topFrame);
    fileDropdown->setMinimumWidth(200);
	connect(fileDropdown, QOverload<int>::of(&QComboBox::currentIndexChanged),
		this, &MainWindow::onFileSelected);
	layout->addWidget(fileDropdown);

	// Sheet dropdown
	QLabel* sheetLabel = new QLabel("Sheet:", topFrame);
	layout->addWidget(sheetLabel);

	sheetDropdown = new QComboBox(topFrame);
	sheetDropdown->setMinimumWidth(200);
	connect(sheetDropdown, QOverload<int>::of(&QComboBox::currentIndexChanged),
		this, &MainWindow::onSheetSelected);
	layout->addWidget(sheetDropdown);

	layout->addStretch();

	// Load Button
	loadButton = new QPushButton("Load File", topFrame);
	connect(loadButton, &QPushButton::clicked, this, &MainWindow::onLoadFile);
	layout->addWidget(loadButton);

	// Save button
	saveButton = new QPushButton("Save File", topFrame);
	connect(saveButton, &QPushButton::clicked, this, &MainWindow::onSaveFile);
	layout->addWidget(saveButton);

	topFrame->setLayout(layout);
	topFrame->setFixedHeight(50);

	debugPrint("Top frame created with file/sheet dropdowns and buttons");
}

void MainWindow::createCenterFrame()
{ 
	debugPrint("Creating center frame...");

	centerFrame = new QWidget(this);
	QVBoxLayout* layout = new QVBoxLayout(centerFrame);
	layout->setContentsMargins(0, 0, 0, 0);

	// Data table
	dataTable = new QTableWidget(centerFrame);
	dataTable->setColumnCount(12); // Default to 12 columns per sample
	dataTable->setRowCount(50); // Default to 50 rows, will expand as needed

	// Set headers
    QStringList headers;
    headers << "Puff #" << "Before Weight" << "After Weight" << "TPM" << "Draw Pressure" << "Sample Name"
        << "Resistance" << "Voltage" << "Consistency" << "Average TPM" << "Notes" << "Clog/Smell";
	dataTable->setHorizontalHeaderLabels(headers);

	// Table settings
	dataTable->horizontalHeader()->setStretchLastSection(true);
	dataTable->setAlternatingRowColors(true);
	dataTable->setSelectionBehavior(QAbstractItemView::SelectRows);
	dataTable->setEditTriggers(QAbstractItemView::DoubleClicked);

	layout->addWidget(dataTable);
	centerFrame->setLayout(layout);

	debugPrint("Center frame created with data table");
}

void MainWindow::createBottomFrame()
{
	debugPrint("Creating bottom frame...");

	bottomFrame = new QWidget(this);
	QHBoxLayout* layout = new QHBoxLayout(bottomFrame);
	layout->setContentsMargins(0, 0, 0, 0);

	// Load Images Button
	loadImagesButton = new QPushButton("Load Images", bottomFrame);
	layout->addWidget(loadImagesButton);

	// Image display area
	imageLabel = new QLabel(bottomFrame);
    imageLabel->setText("No images loaded");
	imageLabel->setAlignment(Qt::AlignCenter);
    imageLabel->setStyleSheet("border: 1px solid gray; background-color: #f0f0f0;");
	imageLabel->setMinimumSize(400, 150);
	layout->addWidget(imageLabel, 1);

	bottomFrame->setLayout(layout);
	bottomFrame->setFixedHeight(150);

	debugPrint("Bottom frame created with image display area");
}

void MainWindow::onNewFile()
{
	debugPrint("New File action triggered");
	QMessageBox::information(this, "New File", "New file functionality will be implemented here.");
}

void MainWindow::onLoadFile()
{
	debugPrint("Load File action triggered");

	QString fileName = QFileDialog::getOpenFileName(
		this,
		"Load Excel File",
		"",
		"Excel Files (*.xlsx *.xls);;All Files(*)"
	);

	if (fileName.isEmpty()) {
		debugPrint("Load file cancelled by user");
		return;
	}

	debugPrint("Selected file: " + fileName);
	currentFile = fileName;

	statusBar()->showMessage("File loaded: " + fileName);

	// TODO: Implement Excel file loading
	QMessageBox::information(this, "Load File", "Excel loading will be implemented next.\nSelected: " + fileName);
}

void MainWindow::onSaveFile()
{
	debugPrint("Save File action triggered");

	if (currentFile.isEmpty()) {
		debugPrint("No current file to save");
		QMessageBox::warning(this, "Save File", "No file is currently loaded");
		return;
	}

	debugPrint("Saving file: " + currentFile);
	statusBar()->showMessage("File saved: " + currentFile);

	// TODO: Implement Excel file saving
	QMessageBox::information(this, "Save File", "Excel saving will be implemented");
}

void MainWindow::onExit()
{
	debugPrint("Exit action triggered");
	QApplication::quit();
}

void MainWindow::onFileSelected(int index)
{
	if (index < 0) return;

	QString selectedFile = fileDropdown->currentText();
	debugPrint("File selected: " + selectedFile);

	currentFile = selectedFile;
	updateSheetDropdown();
}

void MainWindow::onSheetSelected(int index)
{
	if (index < 0) return;

	QString selectedSheet = sheetDropdown->currentText();
	debugPrint("Sheet selected: " + selectedSheet);

	currentSheet = selectedSheet;
	loadSheetData(selectedSheet);
}

void MainWindow::onGenerateTestReport()
{
	debugPrint("Generate Test Report action triggered");
	QMessageBox::information(this, "Generate Test Report", "Test report generation will be implemented here.");
}

void MainWindow::onGenerateFullReport()
{
    debugPrint("Generate Full Report action triggered");
    QMessageBox::information(this, "Generate Full Report", "Full report generation will be implemented here.");
}

void MainWindow::onAbout()
{
    debugPrint("About action triggered");
    QString aboutText = "DataViewer Enterprise v1.0\n"
                        "Developed by Charlie Becquet\n"
                        "C++ Qt Implementation";
    QMessageBox::information(this, "About", aboutText);
}

void MainWindow::onHelp()
{
	debugPrint("Help action triggered");
	QString helpText = "DataViewer Enterprise \n\n"
		"This program is designed to be used with TPM data according to a standardized testing template.\n\n"
		"Use File -> Load to open data files in the window \n"
		"Use Reports menu to generate powerpoint reports";
    QMessageBox::information(this, "Help", helpText);
}

void MainWindow::updateFileDropdown()
{
	debugPrint("Updating file dropdown...");
	fileDropdown->clear();

	// TODO: Populate with loaded files

	debugPrint("File dropdown updated");
}

void MainWindow::updateSheetDropdown()
{
	debugPrint("Updating sheet dropdown...");
	sheetDropdown->clear();

	if (currentFile.isEmpty()) {
		debugPrint("No current file selected");
		return;
	}

	// TODO: Load sheets from Excel file
	// For now, add placeholder sheets
	sheetDropdown->addItem("Intense Test");
	sheetDropdown->addItem("User Test Simulation");
	sheetDropdown->addItem("Temperature Cycling Test");

	debugPrint("Sheet dropdown updated");
}

void MainWindow::loadSheetData(const QString& sheetName)
{
	debugPrint("Loading sheet data for: " + sheetName);

	if (sheetName.isEmpty()) {
		debugPrint("Sheet name is empty");
		return;
	}

	// TODO: Load actual data from Excel
	// For now, clear the table
	dataTable->clearContents();

	statusBar()->showMessage("Loaded sheet: " + sheetName);
	debugPrint("Sheet data loaded");
}

void MainWindow::debugPrint(const QString& message)
{
	qDebug() << "DEBUG: " << message;
}

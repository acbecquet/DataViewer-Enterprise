#include "MainWindow.h"
#include <QApplication>
#include <QMessageBox>
#include <QFileDialog>
#include <QHeaderView>
#include <QFileInfo>

MainWindow::MainWindow(QWidget *parent)
	: QMainWindow(parent)
{
	debugPrint("MainWindow constructor starting...");

	setWindowTitle("DataViewer Enterprise v1.0");
	resize(1200, 800);

	setupUI();

	// Initialize Excel Reader
	m_excelReader = new ExcelReader();
	m_currentSampleIndex = -1;
    debugPrint("Excel reader initilaized");

	debugPrint("MainWindow constructor comple");
}

MainWindow::~MainWindow()
{
	// Clean up excel reader
	if (m_excelReader)
	{
		delete m_excelReader;
		m_excelReader = nullptr;
	}

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
	debugPrint("Creating center frame with split layout...");

	centerFrame = new QWidget(this);
	QHBoxLayout* mainLayout = new QHBoxLayout(centerFrame);
	mainLayout->setContentsMargins(0, 0, 0, 0);
	mainLayout->setSpacing(10);

	// LEFT PANEL: Table and Statistics (50% width)
	leftPanel = new QWidget(centerFrame);
	QVBoxLayout* leftLayout = new QVBoxLayout(leftPanel);
	leftLayout->setContentsMargins(0, 0, 0, 0);
	leftLayout->setSpacing(5);

	// Sample navigation controls at top
	sampleNavFrame = new QWidget(leftPanel);
	QHBoxLayout* navLayout = new QHBoxLayout(sampleNavFrame);
	navLayout->setContentsMargins(0, 0, 0, 0);

	prevSampleButton = new QPushButton("<< Previous", sampleNavFrame);
	connect(prevSampleButton, &QPushButton::clicked, this, &MainWindow::onPrevSample);
	navLayout->addWidget(prevSampleButton);

	sampleCountLabel = new QLabel("Sample 1 of 1", sampleNavFrame);
	sampleCountLabel->setAlignment(Qt::AlignCenter);
	QFont labelFont = sampleCountLabel->font();
	labelFont.setBold(true);
	sampleCountLabel->setFont(labelFont);
	navLayout->addWidget(sampleCountLabel);

	nextSampleButton = new QPushButton("Next >>", sampleNavFrame);
	connect(nextSampleButton, &QPushButton::clicked, this, &MainWindow::onNextSample);
	navLayout->addWidget(nextSampleButton);

	leftLayout->addWidget(sampleNavFrame);

	// Data table
	dataTable = new QTableWidget(leftPanel);
	dataTable->setColumnCount(12); // Default to 12 columns per sample
	dataTable->setRowCount(50); // Default to 50 rows, will expand as needed

	// Set headers
	QStringList headers;
	headers << "Puffs" << "Before Weight" << "After Weight" << "Draw Pressure" << "Resistance"
		<< "Smell" << "Clog" << "Notes" << "TPM (mg/puff)" << "TPM Power Density"
		<< "Variation in TPM (%)" << "Oil Consumed";
	dataTable->setHorizontalHeaderLabels(headers);

	// Table settings
	dataTable->horizontalHeader()->setStretchLastSection(true);
	dataTable->setAlternatingRowColors(true);
	dataTable->setSelectionBehavior(QAbstractItemView::SelectRows);
	dataTable->setEditTriggers(QAbstractItemView::DoubleClicked);

	leftLayout->addWidget(dataTable, 3);  // Table gets more space (3x weight)

	// Sample statistics frame
	statsFrame = new QWidget(leftPanel);
	QVBoxLayout* statsLayout = new QVBoxLayout(statsFrame);
	statsLayout->setContentsMargins(5, 5, 5, 5);
	statsFrame->setStyleSheet("QWidget { background-color: #f5f5f5; border: 1px solid #ccc; border-radius: 5px; }");

	QLabel* statsTitle = new QLabel("Sample Statistics:", statsFrame);
	QFont titleFont = statsTitle->font();
	titleFont.setBold(true);
	titleFont.setPointSize(titleFont.pointSize() + 1);
	statsTitle->setFont(titleFont);
	statsLayout->addWidget(statsTitle);

	statsLabel = new QLabel("No data loaded", statsFrame);
	statsLabel->setWordWrap(true);
	statsLabel->setTextFormat(Qt::RichText);
	statsLayout->addWidget(statsLabel);

	leftLayout->addWidget(statsFrame, 1);  // Stats gets less space (1x weight)

	leftPanel->setLayout(leftLayout);
	mainLayout->addWidget(leftPanel, 1);  // 50% width

	// RIGHT PANEL: Plot area (50% width)
	rightPanel = new QWidget(centerFrame);
	QVBoxLayout* rightLayout = new QVBoxLayout(rightPanel);
	rightLayout->setContentsMargins(0, 0, 0, 0);

	plotFrame = new QWidget(rightPanel);
	QVBoxLayout* plotLayout = new QVBoxLayout(plotFrame);
	plotLayout->setContentsMargins(5, 5, 5, 5);

	plotLabel = new QLabel("Plot area - To be implemented", plotFrame);
	plotLabel->setAlignment(Qt::AlignCenter);
	plotLabel->setStyleSheet("QLabel { background-color: #ffffff; border: 2px dashed #ccc; }");
	plotLabel->setMinimumHeight(400);

	plotLayout->addWidget(plotLabel);
	plotFrame->setLayout(plotLayout);
	rightLayout->addWidget(plotFrame);

	rightPanel->setLayout(rightLayout);
	mainLayout->addWidget(rightPanel, 1);  // 50% width

	centerFrame->setLayout(mainLayout);

	debugPrint("Center frame created with table (left) and plot (right) panels");
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

	QString filePath = QFileDialog::getOpenFileName(
		this,
		"Load Excel File",
		"",
		"Excel Files (*.xlsx *.xls);;All Files(*)"
	);

	if (filePath.isEmpty()) {
		debugPrint("Load file cancelled by user");
		return;
	}

	debugPrint("Selected file: " + filePath);
	
	// Load the excel file
	if (!m_excelReader->loadFile(filePath))
	{
		QString error = m_excelReader->getLastError();
		debugPrint("ERROR: Failed to load file - " + error);
		QMessageBox::critical(this, "Load Error", "Failed to load Excel file:\n" + error);
		return;
	}

	currentFile = filePath;
	debugPrint("File loaded successfully");

	// Update UI
	updateFileDropdown();
	updateSheetDropdown();

	// Load data from first sheet
	loadExcelData();

	statusBar()->showMessage("Loaded: " + QFileInfo(filePath).fileName());
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
	debugPrint("File selected, index: " + QString::number(index));
	// Currently only supporting one file at a time
}

void MainWindow::onSheetSelected(int index)
{
	debugPrint("Sheet Selected, index: " + QString::number(index));

	if (index < 0)
	{
		return;
	}

	QString sheetName = sheetDropdown->itemText(index);
	debugPrint("Sheet name: " + sheetName);

	if (!m_excelReader->selectSheet(sheetName))
	{
		QString error = m_excelReader->getLastError();
		debugPrint("ERROR: Failed to select sheet - " + error);
		QMessageBox::warning(this, "Sheet Selection", "Failed to select sheet:\n" + error);
		return;
	}

	currentSheet = sheetName;

	// Check for deprecated 8-column User Test Simulation format
	if (m_excelReader->isDeprecatedUserTestSimulation())
	{
		debugPrint("WARNING: Deprecated 8-column User Test simulation detected");
		QMessageBox::warning(
			this,
			"Deprecated Format",
			"Warning: Old User Teset Simulation format (8 columns) is deprecated.\n"
			"This sheet will not be processeed. \n\n"
			"Please conver this file to the current 12-column format."
		);

		// Clear the table
		dataTable->clearContents();
		statusBar()->showMessage("Sheet Skipped - deprecated format");
		return;
	}

	// Load data if format is valid
	loadExcelData();
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

	if (!currentFile.isEmpty())
	{
		QFileInfo fileInfo(currentFile);
		fileDropdown->addItem(fileInfo.fileName());
		debugPrint("Added file to dropdown: " + fileInfo.fileName());
	}

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

	// Get sheets from Excel reader
	QStringList sheets = m_excelReader->getSheetNames();
	debugPrint("Found " + QString::number(sheets.size()) + " sheets");

	for (const QString& sheet : sheets)
	{
		sheetDropdown->addItem(sheet);
		debugPrint("Added sheet: " + sheet);
	}

	debugPrint("Sheet dropdown updated");
}

void MainWindow::loadExcelData()
{
	debugPrint("Loading Excel data from current sheet");

	if (currentFile.isEmpty() || currentSheet.isEmpty())
	{
		debugPrint("No file or sheet selected");
		return;
	}

	// Log template info
	QString templateVersion = m_excelReader->detectTemplateVersion();
	debugPrint("Template version: " + templateVersion);
	debugPrint("Processing with 12-column standard format");

	// get all samples from current sheet
	m_currentSamples = m_excelReader->getAllSamples();
	debugPrint("Loaded " + QString::number(m_currentSamples.size()) + " samples");

	// Display first sample if available
	if (!m_currentSamples.isEmpty())
	{
		displaySample(0);
	}
	else
	{
		dataTable->clearContents();
		debugPrint("No samples found in sheet");
		QMessageBox::information(
			this,
			"No Data",
			"No sample data found in this sheet.\n"
			"The sheet may be empty or have an unexpected format."
		);
	}
}

void MainWindow::displaySample(int sampleIndex)
{
	debugPrint("Displaying sample " + QString::number(sampleIndex + 1));

	if (sampleIndex < 0 || sampleIndex >= m_currentSamples.size())
	{
		debugPrint("ERROR: Invalid sample index: " + QString::number(sampleIndex));
		return;
	}

	m_currentSampleIndex = sampleIndex;
	const ExcelReader::SampleData& sample = m_currentSamples[sampleIndex];

	// Log sample metadata
	debugPrint("Sample metadata");
	debugPrint("  Test Name: " + sample.metadata.testName);
	debugPrint("  Sample ID: " + sample.metadata.sampleID);
	debugPrint("  Date: " + sample.metadata.date);
	debugPrint("  Tester: " + sample.metadata.tester);
	debugPrint("  Voltage: " + QString::number(sample.metadata.voltage));
	debugPrint("  Viscosity: " + QString::number(sample.metadata.viscosity));
	debugPrint("  Resistance: " + QString::number(sample.metadata.resistance));
	debugPrint("  Puffing Regime: " + sample.metadata.puffingRegime);
	debugPrint("  Initial Oil Mass: " + QString::number(sample.metadata.initialOilMass));
	debugPrint("  Data rows: " + QString::number(sample.dataRows.size()));

	// Populate table
	populateTableWithSample(sample);

	// Update naviagation controls
	updateSampleNavigation();

	// Update Sample Statistics
	updateSampleStatistics(sample);

	statusBar()->showMessage("Displaying sample " + QString::number(sampleIndex + 1) +
		" of " + QString::number(m_currentSamples.size()) + " - " + sample.metadata.sampleID);
}

void MainWindow::onPrevSample()
{
	debugPrint("Previous sample button clicked");

	if (m_currentSampleIndex > 0)
	{
		displaySample(m_currentSampleIndex - 1);
	}
	else
	{
		debugPrint("Already at first sample");
	}
}

void MainWindow::onNextSample()
{
	debugPrint("Next sample button clicked");

	if (m_currentSampleIndex < m_currentSamples.size() - 1)
	{
		displaySample(m_currentSampleIndex + 1);
	}
	else
	{
		debugPrint("Already at last sample");
	}
}

void MainWindow::updateSampleNavigation()
{
	debugPrint("Updating sample navigation controls");

	if (m_currentSamples.isEmpty())
	{
		prevSampleButton->setEnabled(false);
		nextSampleButton->setEnabled(false);
		sampleCountLabel->setText("No samples");
		return;
	}

	// Update label
	sampleCountLabel->setText("Sample " + QString::number(m_currentSampleIndex + 1) +
		" of " + QString::number(m_currentSamples.size()));

	// Enable/disable buttons
	prevSampleButton->setEnabled(m_currentSampleIndex > 0);
	nextSampleButton->setEnabled(m_currentSampleIndex < m_currentSamples.size() - 1);

	debugPrint("Navigation updated: " + QString::number(m_currentSampleIndex + 1) +
		"/" + QString::number(m_currentSamples.size()));
}

void MainWindow::updateSampleStatistics(const ExcelReader::SampleData& sample)
{
	debugPrint("Updating sample statistics display");

	// Build statistics HTML
	QString statsHtml = "<table style='width:100%; font-size:10pt;'>";

	// Sample identification
	statsHtml += "<tr><td style='font-weight:bold; width:40%;'>Sample ID:</td><td>" +
		sample.metadata.sampleID + "</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Date:</td><td>" +
		sample.metadata.date + "</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Tester:</td><td>" +
		sample.metadata.tester + "</td></tr>";

	// Device parameters
	statsHtml += "<tr><td colspan='2' style='padding-top:8px; font-weight:bold; background-color:#e0e0e0;'>Device Parameters</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Media:</td><td>" +
		sample.metadata.media + "</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Viscosity:</td><td>" +
		QString::number(sample.metadata.viscosity, 'f', 0) + " cP</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Resistance:</td><td>" +
		QString::number(sample.metadata.resistance, 'f', 2) + " Ω</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Voltage:</td><td>" +
		QString::number(sample.metadata.voltage, 'f', 1) + " V</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Power:</td><td>" +
		QString::number(sample.metadata.power, 'f', 2) + " W</td></tr>";

	if (!sample.metadata.heatingTechnology.isEmpty())
	{
		statsHtml += "<tr><td style='font-weight:bold;'>Heating Tech:</td><td>" +
			sample.metadata.heatingTechnology + "</td></tr>";
	}

	// Test parameters
	statsHtml += "<tr><td colspan='2' style='padding-top:8px; font-weight:bold; background-color:#e0e0e0;'>Test Parameters</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Puffing Regime:</td><td>" +
		sample.metadata.puffingRegime + "</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Initial Oil Mass:</td><td>" +
		QString::number(sample.metadata.initialOilMass, 'f', 2) + " g</td></tr>";
	statsHtml += "<tr><td style='font-weight:bold;'>Total Puffs:</td><td>" +
		QString::number(sample.dataRows.size()) + "</td></tr>";

	statsHtml += "</table>";

	statsLabel->setText(statsHtml);
	debugPrint("Statistics updated");
}

void MainWindow::populateTableWithSample(const ExcelReader::SampleData& sample)
{
	debugPrint("Populating table with sample data");

	// Clear existing data
	dataTable->clearContents();

	// set row count based on data
	int rowCount = sample.dataRows.size();
	if (rowCount > dataTable->rowCount())
	{
		dataTable->setRowCount(rowCount);
		debugPrint("Expanded Table to " + QString::number(rowCount) + " rows");
	}

	// get column headers
	QStringList headers = m_excelReader->getColumnHeaders();
	if (!headers.isEmpty())
	{
		dataTable->setHorizontalHeaderLabels(headers);
		debugPrint("Set column headers: " + headers.join(", "));
	}

	// populate rows
	for (int row = 0; row < rowCount; row++)
	{
		const QVector<QVariant>& rowData = sample.dataRows[row];

		for (int col = 0; col < rowData.size() && col < dataTable->columnCount(); col++)
		{
			QVariant value = rowData[col];

			QTableWidgetItem* item = new QTableWidgetItem();

			if (value.isNull())
			{
				item->setText("");
			}
			else if (value.type() == QVariant::Double || value.type() == QVariant::Int)
			{
				// format numeric values
				double numValue = value.toDouble();
				item->setText(QString::number(numValue, 'f', 4));
				item->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
			}
			else
			{
				item->setText(value.toString());
			}

			dataTable->setItem(row, col, item);
		}

		// Auto resize columns to content
		dataTable->resizeColumnsToContents();

		debugPrint("Table populated with " + QString::number(rowCount) + " rows");
	}
}

void MainWindow::debugPrint(const QString& message)
{
	qDebug() << "DEBUG: " << message;
}

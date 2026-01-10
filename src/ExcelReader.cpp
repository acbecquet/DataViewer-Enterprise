#include "ExcelReader.h"
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include <QDebug>
#include <QFileInfo>

ExcelReader::ExcelReader()
	: m_document(nullptr)
	, m_worksheet(nullptr)
{
	debugPrint("ExcelReader constructor");
}

ExcelReader::~ExcelReader()
{
	closeFile();
	debugPrint("ExcelReader destructor");
}

void ExcelReader::debugPrint(const QString& message) const
{
	qDebug() << "DEBUG [ExcelReader]:" << message;
}

bool ExcelReader::loadFile(const QString& filePath)
{
	debugPrint("Loading file: " + filePath);

	// Close any existing file
	closeFile();

	// Validate file path
	QFileInfo fileInfo(filePath);
	if (!fileInfo.exists())
	{
		m_lastError = "File does not exist: " + filePath;
		debugPrint("Error: " + m_lastError);
		return false;
	}

	if (!fileInfo.isFile())
	{
		m_lastError = "Path is not a file: " + filePath;
		debugPrint("ERROR: " + m_lastError);
		return false;
	}

	// Create QXlsx Document
	QXlsx::Document* doc = new QXlsx::Document(filePath);
	if (!doc)
	{
		m_lastError = "Failed to create Excel document object";
		debugPrint("Error: " + m_lastError);
		return false;
	}

	m_document = doc;
	m_filePath = filePath;

	debugPrint("File loaded successfully");
	debugPrint("Available sheets: " + getSheetNames().join(","));

	// Auto-select first sheet
	QStringList sheets = getSheetNames();
	if (!sheets.isEmpty())
	{
		selectSheet(sheets.first());
	}

	return true;
}

void ExcelReader::closeFile()
{
    debugPrint("Closing file");

	if (m_document)
	{
        QXlsx::Document* doc = static_cast<QXlsx::Document*>(m_document);
		delete doc;
		m_document = nullptr;
	}

	m_worksheet = nullptr;
    m_filePath.clear();
	m_currentSheet.clear();
}

QStringList ExcelReader::getSheetNames() const
{
	QStringList sheetNames;

	if (!m_document)
	{
		debugPrint("WARNING: No document loaded, cannot get sheet names");
		return sheetNames;
	}

	QXlsx::Document* doc = static_cast<QXlsx::Document*>(m_document);
	sheetNames = doc->sheetNames();

	debugPrint("Found " + QString::number(sheetNames.size()) + " sheets");
	return sheetNames;
}

bool ExcelReader::selectSheet(const QString& sheetName)
{
	debugPrint("Selecting sheet: " + sheetName);

	if (!m_document)
	{
		m_lastError = "No document loaded";
		debugPrint("ERROR: " + m_lastError);
		return false;
	}

	QXlsx::Document* doc = static_cast<QXlsx::Document*>(m_document);

	// Verify sheet exists
	if (!doc->sheetNames().contains(sheetName))
	{
		m_lastError = "Sheet not found: " + sheetName;
		debugPrint("ERROR: " + m_lastError);
		return false;
	}

	// Select the sheet
	if (!doc->selectSheet(sheetName))
	{
		m_lastError = "Failed to select sheet: " + sheetName;
		debugPrint("ERROR: " + m_lastError);
		return false;
	}

	m_currentSheet = sheetName;
	m_worksheet = doc->currentWorksheet();

	debugPrint("Sheet selected successfully");
	debugPrint("Sample count: " + QString::number(getSampleCount()));

	return true;
}

QVariant ExcelReader::getCellValue(int row, int col) const
{
	if (!m_worksheet)
	{
		return QVariant();
	}

	QXlsx::Worksheet* ws = static_cast<QXlsx::Worksheet*>(m_worksheet);

	// Note: QXlsx uses 1-based idnexing, so add 1 to row and col
	QVariant value = ws->read(row + 1, col + 1);

	return value;
}

QString ExcelReader::getCellString(int row, int col) const
{
    QVariant value = getCellValue(row, col);

	if (value.isNull())
	{
		return QString();
	}

	return value.toString().trimmed();
}

double ExcelReader::getCellDouble(int row, int col) const
{
	QVariant value = getCellValue(row, col);

	if (value.isNull())
	{
		return 0.0;
	}

	bool ok = false;
	double result = value.toDouble(&ok);

	return ok ? result : 0.0;
}

bool ExcelReader::isDeprecatedUserTestSimulation() const
{
	debugPrint("Checking for deprecated 8-column user test simulation format");

	if (!m_worksheet)
	{
		return false;
	}

	// only check if this is a user test simulation sheet
	if (!m_currentSheet.contains("User Test Simulation", Qt::CaseInsensitive))
	{
		return false;
	}

	// Check row 4 (index 3) for column structure
	// In 8-column format, columns 9-12 should be empty or have nonstandard headers
	// In 12-column format, these columns will have headers
	QStringList expectedNew12ColHeaders = { "Average TPM","Consistency","Variation","Oil Consumed" };

	bool foundNew12ColIndicators = false;
	for (int col = 8; col < 12; col++)
	{
		QString headerVal = getCellString(3, col);
        for (const QString& expectedHeader : expectedNew12ColHeaders)
		{
			if (headerVal.contains(expectedHeader, Qt::CaseInsensitive))
			{
				foundNew12ColIndicators = true;
                debugPrint("Found 12-column idnicator: " + headerVal + " at column " + QString::number(col + 1));
				break;
			}
		}

		if (foundNew12ColIndicators)
		{
			break;
		}
	}

	if (!foundNew12ColIndicators)
	{
		debugPrint("WARNING: Detected deprecated 8-column User Test Simulation format");
		return true;
	}

    debugPrint("User Test Simulation is using current 12-column format");
		return false;
}

QString ExcelReader::detectTemplateVersion() const
{
	debugPrint("Detecting template version");

	if (!m_worksheet)
	{
        return "unknown";
	}

	// Check for new template indicators in sheet names
    QStringList sheets = getSheetNames();
	QStringList newTemplateIndicators = {
		"Long Puff lifetime Test",
		"Rapid Puff Lifetime Test",
		"Temperature Cycling Test #1",
		"Temperature Cycling Teset #2"
	};

	for (const QString& indicator : newTemplateIndicators)
	{
		if (sheets.contains(indicator, Qt::CaseInsensitive))
		{
			debugPrint("Detected new template (December 2025");
			return "new";
		}
	}

	// if we get here, it's likely the old template (Jan 2025)
	// but that's ok, columns 1-9 are identical, and we don't process 10-12 yet
	debugPrint("Detected old template (January 2025) - compatible with columns 1-9");
	return "old";
}

int ExcelReader::countSamples() const
{
	if (!m_worksheet)
	{
		return 0;
	}

	// Always use 12 columns per sample
	const int COLUMNS_PER_SAMPLE = 12;

	QXlsx::Worksheet* ws = static_cast<QXlsx::Worksheet*>(m_worksheet);

	// Get the used range to determine the column count
	QXlsx::CellRange range = ws->dimension();
	int totalColumns = range.columnCount();

	// calculate number of samples
	int sampleCount = totalColumns / COLUMNS_PER_SAMPLE;

	debugPrint("Total Columns: " + QString::number(totalColumns) +
		", Columns per sample: " + QString::number(COLUMNS_PER_SAMPLE) +
		", Sample count: " + QString::number(sampleCount));

	return sampleCount;
}

int ExcelReader::getSampleCount() const
{
	return countSamples();
}

ExcelReader::SampleMetadata ExcelReader::extractMetadata(int sampleIndex) const
{
	SampleMetadata metadata;

	if (!m_worksheet)
	{
		return metadata;
	}

	// Always use 12 columns per sample
	const int COLUMNS_PER_SAMPLE = 12;
	int colOffset = sampleIndex * COLUMNS_PER_SAMPLE;

	debugPrint("Extracting metadata for sample " + QString::number(sampleIndex + 1) +
		" at column offset " + QString::number(colOffset));

	// Detect template version
	QString templateVersion = detectTemplateVersion();
	debugPrint("Using template version: " + templateVersion);

	if (templateVersion == "new")
	{
		// NEW TEMPLATE (December 2025) structure
		// Row 1 (index 0): Test name (col A), Date (col C), Sample ID (col E), Heating Technology (col F)
		metadata.testName = getCellString(0, colOffset + 0);  // Column A + offset
		metadata.date = getCellString(0, colOffset + 2);      // Column C + offset
		metadata.sampleID = getCellString(0, colOffset + 4);  // Column E + offset
		metadata.heatingTechnology = getCellString(0, colOffset + 5); // Column F + offset

		// Row 2 (index 1): Media (col A), Resistance (col C), Power (col E)
		metadata.media = getCellString(1, colOffset + 0);      // Column A + offset
		metadata.resistance = getCellDouble(1, colOffset + 2); // Column C + offset
		// Power will be calculated below or read from E2

		// Row 3 (index 2): Viscosity (col A), Tester (col C), Voltage (col E), Puffing Regime (col G), Initial Oil Mass (col H)
		metadata.viscosity = getCellDouble(2, colOffset + 0);         // Column A + offset
		metadata.tester = getCellString(2, colOffset + 2);            // Column C + offset
		metadata.voltage = getCellDouble(2, colOffset + 4);           // Column E + offset
		metadata.puffingRegime = getCellString(2, colOffset + 6);     // Column G + offset
		metadata.initialOilMass = getCellDouble(2, colOffset + 7);    // Column H + offset
	}
	else
	{
		// OLD TEMPLATE (January 2025) structure
		// Row 1 (index 0): Test name (col A), Date (col D), Sample ID (col G)
		metadata.testName = getCellString(0, colOffset + 0);   // Column A + offset
		metadata.date = getCellString(0, colOffset + 3);       // Column D + offset (Date value)
		metadata.sampleID = getCellString(0, colOffset + 6);   // Column G + offset (Sample ID value)

		// No heating technology in old template
		metadata.heatingTechnology = "";

		// Row 2 (index 1): Media value (col B), Resistance value (col D), Puffing Regime value (col I)
		metadata.media = getCellString(1, colOffset + 1);        // Column B + offset (Media value)
		metadata.resistance = getCellDouble(1, colOffset + 3);   // Column D + offset (Resistance value)
		metadata.puffingRegime = getCellString(1, colOffset + 8);// Column I + offset (Puffing Regime value)

		// Row 3 (index 2): Viscosity value (col B), Tester value (col D), Voltage value (col G), Initial Oil Mass (col M or nearby)
		metadata.viscosity = getCellDouble(2, colOffset + 1);    // Column B + offset (Viscosity value)
		metadata.tester = getCellString(2, colOffset + 3);       // Column D + offset (Tester value)
		metadata.voltage = getCellDouble(2, colOffset + 6);      // Column G + offset (Voltage value)
		metadata.initialOilMass = getCellDouble(2, colOffset + 12); // Column M + offset (Initial Oil Mass value - approximate position)
	}

	// Calculate power: P = V^2 / (R + Roffset)
	// For old template, Roffset is always 0 (no heating technology field)
	double rOffset = 0.0;

	if (!metadata.heatingTechnology.isEmpty())
	{
		QString tech = metadata.heatingTechnology.trimmed().toLower();
		if (tech.contains("t51"))
		{
			rOffset = 0.25;
			debugPrint("Heating technology T51 detected, using Roffset = 0.25");
		}
		else if (tech.contains("t58g") || tech.contains("ccell3.0") || tech.contains("ccell 3.0"))
		{
			rOffset = 0.78;
			debugPrint("Heating technology T58G/CCELL3.0 detected, using Roffset = 0.78");
		}
	}
	else
	{
		debugPrint("No heating technology (old template), using Roffset = 0.0");
	}

	// Calculate power only if we have valid voltage and resistance
	if (metadata.voltage > 0 && metadata.resistance > 0)
	{
		double denominator = metadata.resistance + rOffset;
		if (denominator > 0)
		{
			metadata.power = (metadata.voltage * metadata.voltage) / denominator;
			debugPrint("Calculated power: " + QString::number(metadata.power, 'f', 4) +
				" W (V=" + QString::number(metadata.voltage) +
				", R=" + QString::number(metadata.resistance) +
				", Roffset=" + QString::number(rOffset) + ")");
		}
		else
		{
			metadata.power = 0.0;
			debugPrint("Cannot calculate power: denominator is zero");
		}
	}
	else
	{
		metadata.power = 0.0;
		debugPrint("Cannot calculate power: voltage or resistance is zero");
	}

	debugPrint("Metadata extracted - Sample ID: " + metadata.sampleID +
		", Tester: " + metadata.tester +
		", Voltage: " + QString::number(metadata.voltage) +
		", Power: " + QString::number(metadata.power));

	return metadata;
}

QVector<QVariant> ExcelReader::extractRow(int rowIndex, int startCol, int numCols) const
{
    QVector<QVariant> rowData;

	if (!m_worksheet)
	{
        return rowData;
	}

	// Extract specified number of columns
	for (int col = 0; col < numCols; col++)
	{
		QVariant value = getCellValue(rowIndex, startCol + col);
		rowData.append(value);
	}

	return rowData;
}

QStringList ExcelReader::getColumnHeaders() const
{
	QStringList headers;

	if (!m_worksheet)
	{
		return headers;
	}

	// Always 12 col per sample
	const int COLUMNS_PER_SAMPLE = 12;

	// row 4 contains column headers
	for (int col = 0; col < COLUMNS_PER_SAMPLE; col++)
	{
		QString header = getCellString(3, col);
		headers.append(header);
	}

	debugPrint("Extracted " + QString::number(headers.size()) + " column headers: " + headers.join(", "));

    return headers;
}

ExcelReader::SampleData ExcelReader::getSample(int sampleIndex)
{
	SampleData sample;

	if (!m_worksheet)
	{
		debugPrint("ERROR: No worksheet loaded");
		return sample;
	}

	if (sampleIndex < 0 || sampleIndex >= getSampleCount())
	{
		debugPrint("ERROR: Invalid sample index: " + QString::number(sampleIndex));
		return sample;
	}

	debugPrint("Extracting sample" + QString::number(sampleIndex + 1));

	const int COLUMNS_PER_SAMPLE = 12;
	int colOffset = sampleIndex * COLUMNS_PER_SAMPLE;

	sample.startColumn = colOffset;
	sample.metadata = extractMetadata(sampleIndex);

	// Extract data rows from row 5 after the header row, continue until rows are empty
	QXlsx::Worksheet* ws = static_cast<QXlsx::Worksheet*>(m_worksheet);
	QXlsx::CellRange range = ws->dimension();
	int maxRows = range.rowCount();

	int dataRowCount = 0;
	for (int row = 4; row < maxRows; row++)
	{
		QVector<QVariant> rowData = extractRow(row, colOffset, COLUMNS_PER_SAMPLE);

		// check if row is empty (all values are null or empty)
		bool isEmpty = true;
		for (const QVariant& val : rowData)
		{
            if (!val.isNull() && !val.toString().trimmed().isEmpty())
			{
				isEmpty = false;
				break;
            }
		}

		if (isEmpty)
		{
			// stop at first empty row
			break;
		}

		sample.dataRows.append(rowData);
		dataRowCount++;
	}

	debugPrint("Extracted " + QString::number(dataRowCount) + " data rows for sample " + QString::number(sampleIndex + 1));

	return sample;
}

QVector<ExcelReader::SampleData> ExcelReader::getAllSamples()
{
	QVector<SampleData> samples;

	int sampleCount = getSampleCount();
	debugPrint("Extracting all " + QString::number(sampleCount) + " samples");

	for (int i = 0; i < sampleCount; i++)
	{
		SampleData sample = getSample(i);
		samples.append(sample);
	}

	return samples;
}

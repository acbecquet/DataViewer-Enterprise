#ifndef EXCELREADER_H
#define EXCELREADER_H

#include <QString>
#include <QStringList>
#include <QVector>
#include <QVariant>
#include <QMap>

class ExcelReader
{
public:
	struct SampleMetadata
	{
		QString testName;
		QString date;
		QString sampleID;
		QString media;
		double resistance;
		double voltage;
		double power;
		double viscosity;
		QString tester;
		QString puffingRegime;
		double initialOilMass;
		QString heatingTechnology;
	};

	struct SampleData
	{
		SampleMetadata metadata;
		QVector<QVector<QVariant>> dataRows; // Row x Column Data
		int startColumn; // Starting column index for this sample (0-based)
	};


	ExcelReader();
	~ExcelReader();

	// File operations
	bool loadFile(const QString& fileetPath);
	void closeFile();
	QString getFilePath() const { return m_filePath; }

	// Sheet operations
	QStringList getSheetNames() const;
	bool selectSheet(const QString& sheetName);
	QString getCurrentSheet() const { return m_currentSheet; }

	// Data extraction
	int getSampleCount() const;
	QVector<SampleData> getAllSamples();
	SampleData getSample(int sampleIndex);

	// Column headers (row 4)
	QStringList getColumnHeaders() const;

	// Template detection
	QString detectTemplateVersion() const;
	bool isDeprecatedUserTestSimulation() const;

	// Error Handling
	QString getLastError() const { return m_lastError; }

private:
	QString m_filePath;
	QString m_currentSheet;
	QString m_lastError;
	void* m_document; // Opaque pointer to QXlsx::Document
	void* m_worksheet; // Opaque pointer to QXlsx::Worksheet

	// Helper functions
	void debugPrint(const QString& message) const;
	QVariant getCellValue(int row, int col) const;
	QString getCellString(int row, int col) const;
	double getCellDouble(int row, int col) const;

	SampleMetadata extractMetadata(int sampleIndex) const;
	QVector<QVariant> extractRow(int rowIndex, int startCol, int numCols) const;
	int countSamples() const;
};

#endif // EXCELREADER_H
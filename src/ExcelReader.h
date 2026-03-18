#ifndef EXCELREADER_H
#define EXCELREADER_H

#include <QString>
#include <QStringList>
#include <QVector>
#include <QVariant>

// ---------------------------------------------------------------------------
// ExcelReader
//
// Reads .xlsx files by invoking Python (openpyxl) as a subprocess and
// storing all sheet data in memory as QVariant 2-D arrays.  This approach
// is transparent to any server-side encryption (AIP/MIP/IRM) because
// openpyxl runs inside the authenticated user session and Windows
// handles decryption automatically.  Using data_only=True also returns
// cached computed values for formula cells instead of formula strings.
// ---------------------------------------------------------------------------
class ExcelReader
{
public:
    struct SampleMetadata
    {
        QString testName;
        QString date;
        QString sampleID;
        QString media;
        double  resistance    = 0.0;
        double  voltage       = 0.0;
        double  power         = 0.0;
        double  viscosity     = 0.0;
        QString tester;
        QString puffingRegime;
        double  initialOilMass = 0.0;
        QString heatingTechnology;
    };

    struct SampleData
    {
        SampleMetadata            metadata;
        QVector<QVector<QVariant>> dataRows;   // row × column, 0-based
        int                       startColumn = 0;
    };

    ExcelReader();
    ~ExcelReader();

    // File operations
    bool        loadFile(const QString& filePath);
    void        closeFile();
    QString     getFilePath() const { return m_filePath; }

    // Sheet operations
    QStringList getSheetNames()  const;
    bool        selectSheet(const QString& sheetName);
    QString     getCurrentSheet() const { return m_currentSheet; }

    // Data extraction
    int                    getSampleCount() const;
    QVector<SampleData>    getAllSamples();
    SampleData             getSample(int sampleIndex);
    QStringList            getColumnHeaders() const;

    // Raw sheet access — returns all cells of the current sheet as strings.
    // Row 0 is the header row; subsequent rows are data rows.
    QVector<QStringList>   getSheetRows() const;

    // Template detection
    QString detectTemplateVersion()        const;
    bool    isDeprecatedUserTestSimulation() const;

    // Error handling
    QString getLastError() const { return m_lastError; }

private:
    // ── In-memory storage ────────────────────────────────────────────────
    struct SheetData {
        QString                    name;
        QVector<QVector<QVariant>> cells;  // [row][col], 0-based
    };

    QVector<SheetData> m_sheets;
    int                m_currentSheetIdx = -1;

    QString m_filePath;
    QString m_currentSheet;
    QString m_lastError;

    // ── Private helpers ──────────────────────────────────────────────────
    const SheetData* currentSheetData() const;

    void    debugPrint(const QString& msg) const;
    QVariant getCellValue(int row, int col) const;
    QString  getCellString(int row, int col) const;
    double   getCellDouble(int row, int col) const;

    SampleMetadata      extractMetadata(int sampleIndex) const;
    QVector<QVariant>   extractRow(int rowIndex, int startCol, int numCols) const;
    int                 countSamples() const;

    // ── Python subprocess ────────────────────────────────────────────────
    static QString findPython();
    static bool    runPythonReader(const QString& pythonExe,
                                   const QString& filePath,
                                   QVector<SheetData>& out,
                                   QString& error);
};

#endif // EXCELREADER_H

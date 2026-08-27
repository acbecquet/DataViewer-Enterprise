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

    // One worksheet as raw cells (in-memory form of the Python reader output).
    struct SheetData {
        QString                    name;
        QVector<QVector<QVariant>> cells;  // [row][col], 0-based
        // W1 poka-yoke (2026-08-27): count of formula cells whose CACHED value
        // is missing (openpyxl fallback only; COM computes live and reports 0).
        // Nonzero = the workbook was saved by a cache-stripping tool and every
        // formula cell read as empty - destroyed data, not blank input.
        int                        strippedFormulaCells = 0;
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

    // v3: raw typed grid of the current sheet (0-based [row][col]); empty if
    // no sheet selected. Read-only view for the schema-driven reader.
    QVector<QVector<QVariant>> currentSheetCells() const;

    // W1 poka-yoke: stripped-formula-cell count of the current sheet (see
    // SheetData::strippedFormulaCells). 0 when no sheet is selected.
    int currentSheetStrippedFormulas() const;

    // Template detection
    QString detectTemplateVersion()        const;
    bool    isDeprecatedUserTestSimulation() const;

    // Error handling
    QString getLastError() const { return m_lastError; }

    // Convert a header-cell value to a number, tolerating units typed after
    // the digits ("800kcp", "2.09 Ohm", "100 cP"). A k/K immediately after
    // the number multiplies by 1000 (kcP = kilocentipoise). Returns 0 when
    // there is no leading number. Static so tests exercise it directly.
    static double tolerantCellDouble(const QVariant& v);

    // Convert the Python reader's JSON output into SheetData. A malformed
    // per-sheet entry (not an object / missing its name) skips that SHEET
    // with a logged error - it never aborts the whole file. Returns false
    // only for whole-document failures (unparseable JSON / a Python error
    // object). Static so tests exercise it directly, no subprocess needed.
    static bool parseSheetsJson(const QByteArray& jsonBytes,
                                QVector<SheetData>& out,
                                QString& error);

private:
    // ── In-memory storage ────────────────────────────────────────────────
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

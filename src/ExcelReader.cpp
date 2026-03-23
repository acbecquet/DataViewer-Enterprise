#include "ExcelReader.h"

#include <QDebug>
#include <QFileInfo>
#include <QProcess>
#include <QProcessEnvironment>
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonObject>
#include <QJsonValue>
#include <QTemporaryFile>
#include <QDir>
#include <QStandardPaths>

// ---------------------------------------------------------------------------
// Python script written to a temp file and executed to read the xlsx.
// * data_only=True  → openpyxl returns cached computed values, not formula text.
// * Runs inside the user's authenticated session, so Windows / OneDrive
//   AIP/MIP decryption is handled transparently.
// Output: JSON array  [ {"name": "Sheet1", "rows": [[val, ...], ...]}, ... ]
// ---------------------------------------------------------------------------
static const char* kPythonScript = R"PY(
import sys, json, os
from datetime import datetime, date

def to_val(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (datetime, date)):
        return str(v)
    # Handle COM date objects (pywintypes.TimeType)
    try:
        import pywintypes
        if isinstance(v, pywintypes.TimeType):
            return str(v)
    except (ImportError, AttributeError):
        pass
    s = str(v).strip()
    if not s:
        return None
    return s

def read_with_com(path):
    """Read Excel via COM automation — evaluates all formulas correctly."""
    import win32com.client
    abs_path = os.path.abspath(path)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(abs_path, False, True)  # UpdateLinks=False, ReadOnly=True
    result = []
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        name = ws.Name
        ur = ws.UsedRange
        if ur is None or ur.Rows.Count == 0 or ur.Columns.Count == 0:
            result.append({"name": name, "rows": []})
            continue
        # Read from A1 to end of used range to keep column indices consistent
        last_row = ur.Row + ur.Rows.Count - 1
        last_col = ur.Column + ur.Columns.Count - 1
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))
        data = rng.Value
        rows = []
        if data is not None:
            if not isinstance(data, tuple):
                rows.append([to_val(data)])
            elif len(data) > 0 and not isinstance(data[0], tuple):
                rows.append([to_val(v) for v in data])
            else:
                for row in data:
                    rows.append([to_val(v) for v in row])
        result.append({"name": name, "rows": rows})
    wb.Close(False)
    excel.Quit()
    return result

def read_with_openpyxl(path):
    """Fallback: openpyxl with data_only=True (cached formula values)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    result = []
    for name in wb.sheetnames:
        ws = wb[name]
        rows = []
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0
        if max_row > 0 and max_col > 0:
            for row in ws.iter_rows(min_row=1, max_row=max_row,
                                    min_col=1, max_col=max_col):
                rows.append([to_val(cell.value) for cell in row])
        result.append({"name": name, "rows": rows})
    return result

try:
    try:
        result = read_with_com(sys.argv[1])
    except Exception:
        result = read_with_openpyxl(sys.argv[1])
    sys.stdout.buffer.write(json.dumps(result, ensure_ascii=False).encode('utf-8'))
    sys.stdout.buffer.write(b'\n')
    sys.stdout.buffer.flush()
except Exception as e:
    import traceback
    msg = {"error": str(e), "trace": traceback.format_exc()}
    print(json.dumps(msg), file=sys.stderr)
    sys.exit(1)
)PY";

// ===========================================================================
// Construction / Destruction
// ===========================================================================

ExcelReader::ExcelReader() { debugPrint("ExcelReader constructor"); }

ExcelReader::~ExcelReader()
{
    closeFile();
    debugPrint("ExcelReader destructor");
}

void ExcelReader::debugPrint(const QString& message) const
{
#ifndef QT_NO_DEBUG
    qDebug() << "DEBUG [ExcelReader]:" << message;
#else
    Q_UNUSED(message);
#endif
}

// ===========================================================================
// File operations
// ===========================================================================

bool ExcelReader::loadFile(const QString& filePath)
{
    debugPrint("Loading file: " + filePath);
    closeFile();

    QFileInfo fi(filePath);
    if (!fi.exists() || !fi.isFile()) {
        m_lastError = "File not found: " + filePath;
        debugPrint("Error: " + m_lastError);
        return false;
    }

    // Locate Python interpreter
    QString python = findPython();
    if (python.isEmpty()) {
        m_lastError = "Python not found. Please ensure Python 3 with openpyxl "
                      "is installed and on PATH.";
        debugPrint("Error: " + m_lastError);
        return false;
    }
    debugPrint("Using Python: " + python);

    // Run the reader script
    QString runError;
    if (!runPythonReader(python, filePath, m_sheets, runError)) {
        m_lastError = runError;
        debugPrint("Python reader error: " + m_lastError);
        return false;
    }

    m_filePath = filePath;
    debugPrint("File loaded successfully");
    debugPrint("Found " + QString::number(m_sheets.size()) + " sheets: " +
               getSheetNames().join(", "));

    if (!m_sheets.isEmpty())
        selectSheet(m_sheets[0].name);

    return true;
}

void ExcelReader::closeFile()
{
    debugPrint("Closing file");
    m_sheets.clear();
    m_currentSheetIdx = -1;
    m_filePath.clear();
    m_currentSheet.clear();
}

// ===========================================================================
// Sheet operations
// ===========================================================================

QStringList ExcelReader::getSheetNames() const
{
    QStringList names;
    for (const auto& s : m_sheets)
        names << s.name;
    debugPrint("Found " + QString::number(names.size()) + " sheets");
    return names;
}

bool ExcelReader::selectSheet(const QString& sheetName)
{
    debugPrint("Selecting sheet: " + sheetName);
    for (int i = 0; i < m_sheets.size(); ++i) {
        if (m_sheets[i].name == sheetName) {
            m_currentSheetIdx = i;
            m_currentSheet    = sheetName;
            debugPrint("Sheet selected successfully");
            debugPrint("Sample count: " + QString::number(getSampleCount()));
            return true;
        }
    }
    m_lastError = "Sheet not found: " + sheetName;
    debugPrint("ERROR: " + m_lastError);
    return false;
}

const ExcelReader::SheetData* ExcelReader::currentSheetData() const
{
    if (m_currentSheetIdx < 0 || m_currentSheetIdx >= m_sheets.size())
        return nullptr;
    return &m_sheets[m_currentSheetIdx];
}

// ===========================================================================
// Cell access
// ===========================================================================

QVariant ExcelReader::getCellValue(int row, int col) const
{
    const SheetData* sd = currentSheetData();
    if (!sd) return QVariant();
    if (row < 0 || row >= sd->cells.size()) return QVariant();
    const auto& rowData = sd->cells[row];
    if (col < 0 || col >= rowData.size()) return QVariant();
    return rowData[col];
}

QString ExcelReader::getCellString(int row, int col) const
{
    QVariant v = getCellValue(row, col);
    return v.isNull() ? QString() : v.toString().trimmed();
}

double ExcelReader::getCellDouble(int row, int col) const
{
    QVariant v = getCellValue(row, col);
    if (v.isNull()) return 0.0;
    bool ok = false;
    double d = v.toDouble(&ok);
    return ok ? d : 0.0;
}

// ===========================================================================
// Sample counting / extraction
// ===========================================================================

int ExcelReader::countSamples() const
{
    const SheetData* sd = currentSheetData();
    if (!sd || sd->cells.size() < 4) return 0;

    const int COLS_PER_SAMPLE = 12;
    // Use the header row (row index 3) to determine the column count.
    int totalCols = sd->cells[3].size();
    int count = totalCols / COLS_PER_SAMPLE;

    debugPrint(QString("Total columns: %1  →  %2 sample(s)").arg(totalCols).arg(count));
    return count;
}

int ExcelReader::getSampleCount() const { return countSamples(); }

QVector<QVariant> ExcelReader::extractRow(int rowIndex, int startCol, int numCols) const
{
    QVector<QVariant> row;
    for (int c = 0; c < numCols; ++c)
        row.append(getCellValue(rowIndex, startCol + c));
    return row;
}

QStringList ExcelReader::getColumnHeaders() const
{
    QStringList headers;
    const int COLS_PER_SAMPLE = 12;
    for (int c = 0; c < COLS_PER_SAMPLE; ++c)
        headers << getCellString(3, c);
    debugPrint("Column headers: " + headers.join(", "));
    return headers;
}

// ===========================================================================
// Template detection
// ===========================================================================

QString ExcelReader::detectTemplateVersion() const
{
    debugPrint("Detecting template version");
    if (m_sheets.isEmpty()) return QStringLiteral("unknown");

    QStringList newIndicators = {
        "Long Puff lifetime Test",
        "Rapid Puff Lifetime Test",
        "Temperature Cycling Test #1",
        "Temperature Cycling Test #2"
    };
    bool hasAnyData = false;
    for (const auto& s : m_sheets) {
        for (const auto& ind : newIndicators) {
            if (s.name.contains(ind, Qt::CaseInsensitive)) {
                debugPrint("Detected new template (December 2025)");
                return QStringLiteral("new");
            }
        }
        // Check if any sheet actually has data (non-empty cells in the data area)
        if (!hasAnyData) {
            for (int r = 0; r < qMin(5, (int)s.cells.size()); ++r) {
                for (const QVariant& cell : s.cells[r]) {
                    if (cell.isValid() && !cell.isNull() && !cell.toString().trimmed().isEmpty()) {
                        hasAnyData = true;
                        break;
                    }
                }
                if (hasAnyData) break;
            }
        }
    }
    if (!hasAnyData) {
        debugPrint("Sheets exist but contain no data — empty template");
        return QStringLiteral("unknown");
    }
    debugPrint("Detected old template (January 2025)");
    return QStringLiteral("old");
}

bool ExcelReader::isDeprecatedUserTestSimulation() const
{
    debugPrint("Checking for deprecated 8-column User Test Simulation format");
    if (!m_currentSheet.contains("User Test Simulation", Qt::CaseInsensitive))
        return false;

    QStringList new12ColHeaders = { "Average TPM", "Consistency", "Variation", "Oil Consumed" };
    for (int col = 8; col < 12; ++col) {
        QString h = getCellString(3, col);
        for (const auto& exp : new12ColHeaders) {
            if (h.contains(exp, Qt::CaseInsensitive)) {
                debugPrint("Found 12-column indicator: " + h);
                return false;
            }
        }
    }
    debugPrint("Detected deprecated 8-column format");
    return true;
}

// ===========================================================================
// Metadata extraction
// ===========================================================================

ExcelReader::SampleMetadata ExcelReader::extractMetadata(int sampleIndex) const
{
    SampleMetadata meta;
    const int COLS = 12;
    int off = sampleIndex * COLS;

    debugPrint(QString("Extracting metadata for sample %1 at col offset %2")
                   .arg(sampleIndex + 1).arg(off));

    QString tv = detectTemplateVersion();

    // Template layout: each metadata field is a label-value pair occupying 2
    // consecutive columns.  Labels are at even relative offsets (0, 2, 4, 6, 8,
    // 10); values are always one column to the right (1, 3, 5, 7, 9, 11).
    // Confirmed layout from D0110 Standardized testing 3-5-26.xlsx:
    // Row 0: col 0 = test name (direct value, no label prefix)
    //        col 2/3 = Date label/value
    //        col 4/5 = Sample ID label/value
    //        col 6/7 = Heating Technology label/value
    // Row 1: col 0/1 = Media label/value
    //        col 2/3 = Resistance label/value
    //        col 6/7 = Puffing Regime label/value
    // Row 2: col 0/1 = Viscosity label/value
    //        col 2/3 = Tester label/value
    //        col 4/5 = Voltage label/value
    //        col 6/7 = Initial Oil Mass label/value
    if (tv == "new") {
        meta.testName          = getCellString(0, off + 0);  // direct value at col A
        meta.date              = getCellString(0, off + 3);
        meta.sampleID          = getCellString(0, off + 5);
        meta.heatingTechnology = getCellString(0, off + 7);
        meta.media             = getCellString(1, off + 1);
        meta.resistance        = getCellDouble(1, off + 3);
        meta.puffingRegime     = getCellString(1, off + 7);
        meta.viscosity         = getCellDouble(2, off + 1);
        meta.tester            = getCellString(2, off + 3);
        meta.voltage           = getCellDouble(2, off + 5);
        meta.initialOilMass    = getCellDouble(2, off + 7);
    } else {
        // Old template (Jan 2025) — same general pattern
        meta.testName       = getCellString(0, off + 0);
        meta.date           = getCellString(0, off + 3);
        meta.sampleID       = getCellString(0, off + 5);
        meta.media          = getCellString(1, off + 1);
        meta.resistance     = getCellDouble(1, off + 3);
        meta.puffingRegime  = getCellString(1, off + 7);
        meta.viscosity      = getCellDouble(2, off + 1);
        meta.tester         = getCellString(2, off + 3);
        meta.voltage        = getCellDouble(2, off + 5);
        meta.initialOilMass = getCellDouble(2, off + 7);
    }

    // Calculate power: P = V² / (R + Roffset)
    // Offset depends on heating technology (matches Excel SWITCH formula):
    //   CCELL3.0 / T58G → +0.78 Ω
    //   T51             → +0.25 Ω
    //   EVO / EVOMAX / SE / Competitor / anything else → 0 (straight V²/R)
    double rOffset = 0.0;
    QString tech = meta.heatingTechnology.trimmed().toUpper();
    if (tech == "CCELL3.0" || tech == "CCELL 3.0" || tech == "T58G")
        rOffset = 0.78;
    else if (tech == "T51")
        rOffset = 0.25;

    double denom = meta.resistance + rOffset;
    if (meta.voltage > 0 && denom > 0)
        meta.power = (meta.voltage * meta.voltage) / denom;

    debugPrint(QString("Metadata: sampleID=%1 voltage=%2 resistance=%3 power=%4")
                   .arg(meta.sampleID)
                   .arg(meta.voltage).arg(meta.resistance).arg(meta.power));
    return meta;
}

// ===========================================================================
// Sample data extraction
// ===========================================================================

ExcelReader::SampleData ExcelReader::getSample(int sampleIndex)
{
    SampleData sample;
    const SheetData* sd = currentSheetData();
    if (!sd) { debugPrint("ERROR: No sheet loaded"); return sample; }

    int total = countSamples();
    if (sampleIndex < 0 || sampleIndex >= total) {
        debugPrint("ERROR: Invalid sample index " + QString::number(sampleIndex));
        return sample;
    }

    const int COLS = 12;
    int off = sampleIndex * COLS;

    sample.startColumn = off;
    sample.metadata    = extractMetadata(sampleIndex);

    int dataRowCount = 0;
    for (int row = 4; row < sd->cells.size(); ++row) {
        QVector<QVariant> rowData = extractRow(row, off, COLS);

        // Stop at the first row that is entirely null / empty
        bool isEmpty = true;
        for (const QVariant& v : rowData) {
            if (!v.isNull() && !v.toString().trimmed().isEmpty()) {
                isEmpty = false;
                break;
            }
        }
        if (isEmpty) break;

        sample.dataRows.append(rowData);
        ++dataRowCount;
    }

    debugPrint(QString("Sample %1: %2 data row(s)").arg(sampleIndex + 1).arg(dataRowCount));
    return sample;
}

QVector<QStringList> ExcelReader::getSheetRows() const
{
    const SheetData* sd = currentSheetData();
    QVector<QStringList> result;
    if (!sd) return result;
    for (const QVector<QVariant>& row : sd->cells) {
        QStringList strRow;
        for (const QVariant& v : row)
            strRow << (v.isNull() ? QString() : v.toString());
        result.append(strRow);
    }
    return result;
}

QVector<ExcelReader::SampleData> ExcelReader::getAllSamples()
{
    QVector<SampleData> samples;
    int n = getSampleCount();
    debugPrint("Extracting all " + QString::number(n) + " sample(s)");
    samples.reserve(n);
    for (int i = 0; i < n; ++i)
        samples.append(getSample(i));
    return samples;
}

// ===========================================================================
// Python subprocess helpers
// ===========================================================================

QString ExcelReader::findPython()
{
    // Try common Python executable names in order
    for (const QString& exe : { QString("python"), QString("python3"), QString("py") }) {
        QProcess p;
        p.start(exe, { "--version" });
        if (p.waitForFinished(5000) && p.exitCode() == 0)
            return exe;
    }
    return {};
}

bool ExcelReader::runPythonReader(const QString& pythonExe,
                                   const QString& filePath,
                                   QVector<SheetData>& out,
                                   QString& error)
{
    // Write the script to a temp file
    QString tempDir = QStandardPaths::writableLocation(QStandardPaths::TempLocation);
    QString scriptPath = tempDir + "/dve_xlsx_reader.py";

    QFile scriptFile(scriptPath);
    if (!scriptFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
        error = "Cannot write Python script to: " + scriptPath;
        return false;
    }
    scriptFile.write(kPythonScript);
    scriptFile.close();

    // Run: python dve_xlsx_reader.py "<filePath>"
    // Force UTF-8 stdout so the Ω / other Unicode chars don't hit cp1252.
    QProcess proc;
    QProcessEnvironment env = QProcessEnvironment::systemEnvironment();
    env.insert("PYTHONIOENCODING", "utf-8");
    env.insert("PYTHONUTF8", "1");
    proc.setProcessEnvironment(env);
    proc.start(pythonExe, { scriptPath, filePath });
    if (!proc.waitForFinished(60000)) {  // 60 s timeout
        proc.kill();
        QFile::remove(scriptPath);
        error = "Python reader timed out";
        return false;
    }
    QFile::remove(scriptPath);

    if (proc.exitCode() != 0) {
        QByteArray errBytes = proc.readAllStandardError();
        error = "Python reader failed (exit " +
                QString::number(proc.exitCode()) + "): " +
                QString::fromUtf8(errBytes).trimmed();
        return false;
    }

    QByteArray jsonBytes = proc.readAllStandardOutput();
    if (jsonBytes.isEmpty()) {
        error = "Python reader returned no output";
        return false;
    }

    // Parse JSON
    QJsonParseError parseError;
    QJsonDocument doc = QJsonDocument::fromJson(jsonBytes, &parseError);
    if (parseError.error != QJsonParseError::NoError) {
        error = "JSON parse error: " + parseError.errorString();
        return false;
    }

    if (!doc.isArray()) {
        // Might be an error object from Python
        QJsonObject obj = doc.object();
        error = "Python error: " + obj["error"].toString();
        return false;
    }

    // Convert JSON → SheetData
    out.clear();
    const QJsonArray sheets = doc.array();
    for (const QJsonValue& sv : sheets) {
        QJsonObject sheetObj = sv.toObject();
        SheetData sd;
        sd.name = sheetObj["name"].toString();

        const QJsonArray rowsArr = sheetObj["rows"].toArray();
        sd.cells.reserve(rowsArr.size());

        for (const QJsonValue& rv : rowsArr) {
            const QJsonArray colsArr = rv.toArray();
            QVector<QVariant> row;
            row.reserve(colsArr.size());
            for (const QJsonValue& cv : colsArr) {
                if (cv.isNull() || cv.isUndefined()) {
                    row.append(QVariant());
                } else if (cv.isBool()) {
                    row.append(QVariant(cv.toBool()));
                } else if (cv.isDouble()) {
                    row.append(QVariant(cv.toDouble()));
                } else {
                    row.append(QVariant(cv.toString()));
                }
            }
            sd.cells.append(row);
        }
        out.append(sd);
    }

    return true;
}

#include "SopDialog.h"

#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QLabel>
#include <QSplitter>
#include <QMessageBox>
#include <QFile>
#include <QDebug>

#include "xlsxdocument.h"

namespace DVE {

SopDialog::SopDialog(const QString& xlsxPath, QWidget* parent)
    : QDialog(parent)
{
    setWindowTitle("Standard Operating Procedures");
    resize(900, 600);

    auto* splitter = new QSplitter(Qt::Horizontal, this);

    // Left: test list
    auto* leftWidget = new QWidget;
    auto* leftLayout = new QVBoxLayout(leftWidget);
    leftLayout->setContentsMargins(0, 0, 0, 0);
    auto* listLabel = new QLabel("Tests");
    QFont lf = listLabel->font();
    lf.setBold(true);
    listLabel->setFont(lf);
    leftLayout->addWidget(listLabel);

    m_testList = new QListWidget;
    leftLayout->addWidget(m_testList);
    splitter->addWidget(leftWidget);

    // Right: detail browser
    auto* rightWidget = new QWidget;
    auto* rightLayout = new QVBoxLayout(rightWidget);
    rightLayout->setContentsMargins(0, 0, 0, 0);
    auto* detailLabel = new QLabel("SOP Details");
    QFont df = detailLabel->font();
    df.setBold(true);
    detailLabel->setFont(df);
    rightLayout->addWidget(detailLabel);

    m_detail = new QTextBrowser;
    m_detail->setOpenExternalLinks(false);
    m_detail->setPlaceholderText("Select a test to view its SOP.");
    rightLayout->addWidget(m_detail);
    splitter->addWidget(rightWidget);

    splitter->setStretchFactor(0, 1);
    splitter->setStretchFactor(1, 3);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->addWidget(splitter);

    connect(m_testList, &QListWidget::currentRowChanged,
            this, &SopDialog::onTestSelected);

    loadFromExcel(xlsxPath);
}

void SopDialog::loadFromExcel(const QString& xlsxPath)
{
    if (!QFile::exists(xlsxPath)) {
        QMessageBox::warning(this, "SOP Load Error",
                             "SOP file not found:\n" + xlsxPath);
        return;
    }

    QXlsx::Document xlsx(xlsxPath);

    // Row 1 is headers, data starts at row 2
    // Columns: A=Test, B=SOP, C=Objective, D=Pass Criteria, E=Equipment,
    //          F=Quantity, G=Est Duration (1mL), H=Est Duration (2mL), I=Note
    for (int row = 2; row <= 50; ++row) {
        QString test = xlsx.read(row, 1).toString().trimmed();
        if (test.isEmpty()) break;

        SopEntry e;
        e.test           = test;
        e.sop            = xlsx.read(row, 2).toString().trimmed();
        e.objective      = xlsx.read(row, 3).toString().trimmed();
        e.passCriteria   = xlsx.read(row, 4).toString().trimmed();
        e.equipment      = xlsx.read(row, 5).toString().trimmed();
        e.quantity       = xlsx.read(row, 6).toString().trimmed();
        e.estDuration1mL = xlsx.read(row, 7).toString().trimmed();
        e.estDuration2mL = xlsx.read(row, 8).toString().trimmed();
        e.note           = xlsx.read(row, 9).toString().trimmed();
        m_entries.append(e);

        m_testList->addItem(e.test);
    }

    if (!m_entries.isEmpty())
        m_testList->setCurrentRow(0);
}

void SopDialog::onTestSelected(int row)
{
    if (row < 0 || row >= m_entries.size()) {
        m_detail->clear();
        return;
    }

    const SopEntry& e = m_entries[row];

    QString html;
    html += QStringLiteral("<h2 style='color:#1F4E79;'>%1</h2>").arg(e.test.toHtmlEscaped());

    if (!e.sop.isEmpty())
        html += QStringLiteral("<p><b>SOP:</b> %1</p>").arg(e.sop.toHtmlEscaped());

    if (!e.objective.isEmpty())
        html += QStringLiteral("<h3 style='color:#2E75B6;'>Objective</h3><p>%1</p>")
                    .arg(e.objective.toHtmlEscaped().replace("\n", "<br>"));

    if (!e.passCriteria.isEmpty())
        html += QStringLiteral("<h3 style='color:#2E75B6;'>Pass Criteria</h3><p>%1</p>")
                    .arg(e.passCriteria.toHtmlEscaped().replace("\n", "<br>"));

    if (!e.equipment.isEmpty())
        html += QStringLiteral("<h3 style='color:#2E75B6;'>Equipment</h3><p>%1</p>")
                    .arg(e.equipment.toHtmlEscaped().replace("\n", "<br>"));

    if (!e.quantity.isEmpty())
        html += QStringLiteral("<p><b>Quantity:</b> %1</p>").arg(e.quantity.toHtmlEscaped());

    // Duration table
    if (!e.estDuration1mL.isEmpty() || !e.estDuration2mL.isEmpty()) {
        html += QStringLiteral("<h3 style='color:#2E75B6;'>Estimated Duration</h3>");
        html += QStringLiteral(
            "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;'>"
            "<tr style='background:#1F4E79; color:white;'>"
            "<th>1 mL</th><th>2 mL</th></tr><tr>"
            "<td align='center'>%1</td><td align='center'>%2</td>"
            "</tr></table>")
                    .arg(e.estDuration1mL.toHtmlEscaped(), e.estDuration2mL.toHtmlEscaped());
    }

    if (!e.note.isEmpty())
        html += QStringLiteral("<h3 style='color:#2E75B6;'>Notes</h3><p><i>%1</i></p>")
                    .arg(e.note.toHtmlEscaped().replace("\n", "<br>"));

    m_detail->setHtml(html);
}

} // namespace DVE
